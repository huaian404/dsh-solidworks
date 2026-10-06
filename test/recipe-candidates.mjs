/**
 * recipe-candidates.mjs — validate and run the two new recipe candidates
 * through the plugin's real data layer and build path, without the GUI host.
 *
 *   node test/recipe-candidates.mjs
 *
 * The candidate scripts live in parts/recipes/*.vbs.txt with {{NAME}}
 * placeholders. Each one is:
 *   1. validated by validateRecipe (ASCII, placeholder/parameter agreement),
 *   2. materialised with materialize,
 *   3. executed against live SolidWorks through the plugin's own executeScript,
 *   4. measured with ext.CreateMassProperty2 and compared to a closed form.
 *
 * The volume assertion is the point: solidworks_verify only reads feature
 * kinds and body counts, so a feature that silently did nothing still verifies.
 * Here a wrong volume fails the recipe.
 */
import { readFileSync, mkdirSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { tmpdir } from 'node:os'

const plugin = await import('../lib/index.js')
const { validateRecipe, materialize } = await import('../lib/recipes.js')

const scratchDir = process.env.SW_TEST_SCRATCH ?? join(tmpdir(), 'dsh-solidworks', 'candidates')
const partTemplate = process.env.SW_TEST_TEMPLATE ?? 'C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2026\\templates\\gb_part.prtdot'
const repoRoot = resolve(import.meta.dirname, '..', '..', '..', '..')

const candidates = [
  {
    name: 'plate_linear_holes',
    description: 'Mounting plate with a linear pattern of through holes.',
    file: 'parts/recipes/plate_linear_holes.vbs.txt',
    parameters: {
      PLATE_X: { value: 0.065, description: 'plate length along X' },
      PLATE_Y: { value: 0.050, description: 'plate width along Y' },
      THK: { value: 0.012, description: 'plate thickness' },
      HOLE_D: { value: 0.009, description: 'hole diameter' },
      HX: { value: 0.050, description: 'seed hole X' },
      HY: { value: 0.012, description: 'seed hole Y' },
      PITCH: { value: 0.012, description: 'pattern pitch along X' },
      COUNT: { value: 3, description: 'number of holes' },
    },
    verify: { bodyCount: 1, featureTypes: ['Extrude', 'Cut', 'Pattern'] },
    // 65 x 50 x 12 plate, three through holes of d9
    expectedVolume: 39000 - 3 * (Math.PI * 4.5 ** 2 * 12),
  },
  {
    name: 'disc_multicontour_cut',
    description: 'Disc with a bore and 8 radial slots cut as one multi-contour sketch.',
    file: 'parts/recipes/disc_multicontour_cut.vbs.txt',
    parameters: {
      R_OUT: { value: 0.060, description: 'disc radius' },
      THK: { value: 0.012, description: 'disc thickness' },
      R_BORE: { value: 0.014, description: 'bore radius' },
      SLOT_R0: { value: 0.038, description: 'slot inner radius' },
      SLOT_R1: { value: 0.046, description: 'slot outer radius' },
      SLOT_HW: { value: 0.004, description: 'slot half width' },
      SLOTS: { value: 8, description: 'number of slots' },
    },
    verify: { bodyCount: 1, featureTypes: ['Extrude', 'Cut'] },
    // disc - bore - 8 slots (the slots all lie inside the disc)
    expectedVolume: Math.PI * 60 ** 2 * 12
      - Math.PI * 14 ** 2 * 12
      - 8 * ((46 - 38) * 8 * 12),
  },
]

const results = []
const check = (name, ok, detail = '') => {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`)
}

mkdirSync(scratchDir, { recursive: true })

const registered = new Map()
plugin.apply(
  {
    tools: { register: (tool) => registered.set(tool.name, tool) },
    logger: { info: () => {} },
    systemPrompt: { section: () => {} },
  },
  { partTemplate, scratchDir, defaultTimeoutMs: 900_000 },
)

// one extra script run after the build to read the volume back
const measureScript = `Option Explicit
Dim swApp, sw, ext, mp, v
Set swApp = CreateObject("SldWorks.Application")
Set sw = swApp.ActiveDoc
Set ext = sw.Extension
On Error Resume Next
Err.Clear
Set mp = ext.CreateMassProperty2
v = mp.Volume
WScript.Echo "VOLUME_MM3=" & FormatNumber(v * 1000000000, 1)
`
const measurePath = join(scratchDir, 'measure.vbs')
writeFileSync(measurePath, measureScript.replace(/\n/g, '\r\n'), 'latin1')

for (const candidate of candidates) {
  console.log(`\n=== ${candidate.name} ===`)
  const script = readFileSync(join(repoRoot, candidate.file), 'utf8')
  const validated = validateRecipe(
    {
      name: candidate.name,
      description: candidate.description,
      script,
      parameters: candidate.parameters,
      verify: candidate.verify,
    },
    `candidate ${candidate.name}`,
  )
  if (!validated.ok) {
    check(`${candidate.name}: validates`, false, validated.reason)
    continue
  }
  check(`${candidate.name}: validates`, true, `${Object.keys(candidate.parameters).length} parameters`)

  const built = materialize(validated.recipe, {})
  if (!built.ok) {
    check(`${candidate.name}: materialises`, false, built.reason)
    continue
  }
  check(`${candidate.name}: materialises`, true)

  const run = await registered.get('solidworks_run').execute(
    { script: built.script, timeoutMs: 900_000 },
    { signal: undefined },
  )
  const tail = (run.log ?? '').trim().split('\n').slice(-4).join(' | ')
  if (run.error || run.exitCode !== 0) {
    check(`${candidate.name}: builds`, false, run.error ?? `exit ${run.exitCode}`)
    continue
  }
  check(`${candidate.name}: builds`, /BUILD COMPLETE/.test(run.log ?? ''), tail)

  const measured = await registered.get('solidworks_run').execute(
    { script: measureScript, timeoutMs: 300_000 },
    { signal: undefined },
  )
  const match = /VOLUME_MM3=([\d.,]+)/.exec(measured.log ?? '')
  if (!match) {
    check(`${candidate.name}: volume measured`, false, (measured.log ?? '').trim().slice(0, 120))
    continue
  }
  const actual = Number(match[1].replace(',', ''))
  const expected = candidate.expectedVolume
  const errorPct = Math.abs(actual - expected) / expected * 100
  check(
    `${candidate.name}: volume matches the closed form`,
    errorPct < 0.5,
    `measured ${actual} mm^3, closed form ${expected.toFixed(1)} mm^3 (${errorPct.toFixed(3)}% off)`,
  )

  const verification = await registered.get('solidworks_verify').execute(
    { expectBodyCount: candidate.verify.bodyCount, featureTypes: candidate.verify.featureTypes },
    { signal: undefined },
  )
  check(
    `${candidate.name}: solidworks_verify`,
    verification.ok === true,
    verification.ok
      ? `bodies ${verification.bodyCount}, kinds [${verification.featureKinds.join(', ')}]`
      : (verification.mismatches ?? [verification.error]).join('; '),
  )
}

const failures = results.filter((r) => !r.ok)
const report = { when: new Date().toISOString(), pass: results.length - failures.length, fail: failures.length, results }
const reportPath = join(scratchDir, 'recipe-candidates.json')
writeFileSync(reportPath, `${JSON.stringify(report, null, 2)}\n`, 'utf8')
console.log(`\n${failures.length === 0 ? 'ALL PASS' : `${failures.length} FAILED`} — report: ${reportPath}`)
process.exit(failures.length === 0 ? 0 : 1)
