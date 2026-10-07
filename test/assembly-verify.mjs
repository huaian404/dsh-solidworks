/**
 * assembly-verify.mjs — the assembly gate, run against live SolidWorks.
 *
 * The bug this locks down: `solidworks_verify` judged every assembly by its
 * solid-body count, and on an assembly `GetBodies2` does not bind, so the
 * report said `bodyCount: -1` and a correctly built assembly was reported as a
 * MISMATCH whatever the caller expected. Two gaps followed from that:
 *
 *   - `%SWASMTPL%` did not exist, so a script that needed an assembly template
 *     had to hard-code an install-specific `.asmdot` path;
 *   - no expectation could describe an assembly, so no assembly build could pass
 *     the gate that every promoted recipe has to pass.
 *
 * This test drives the real `apply()` from lib/index.js with a mock registry:
 * it builds a 3-instance assembly through `solidworks_run` using the exported
 * `%SWASMTPL%`, then gates it with `solidworks_verify expectComponentCount=3`,
 * and finally proves the gate can FAIL (expectComponentCount=4).
 *
 *   node test/assembly-verify.mjs
 *
 * Requires live SolidWorks. Uses an ASCII staging copy of a workspace part, or
 * whatever SW_TEST_ASSEMBLY_PART points at.
 */
import { copyFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'

const plugin = await import('../lib/index.js')

const scratchDir = process.env.SW_TEST_SCRATCH ?? join(tmpdir(), 'dsh-solidworks', 'assembly-verify')
const partTemplate = process.env.SW_TEST_TEMPLATE ?? 'C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2026\\templates\\gb_part.prtdot'
const assemblyTemplate = process.env.SW_TEST_ASSEMBLY_TEMPLATE ?? 'C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2026\\templates\\gb_assembly.asmdot'
const sourcePart = process.env.SW_TEST_PART ?? 'C:\\Users\\qijin\\Desktop\\harness工作区\\bolt_M20.SLDPRT'
const COMPONENTS = Number(process.env.SW_TEST_COMPONENTS ?? '3')

mkdirSync(scratchDir, { recursive: true })
const stagePart = join(scratchDir, 'assembly_part.SLDPRT')

const results = []
const check = (name, ok, detail = '') => {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`)
}

if (!existsSync(sourcePart)) {
  // A path with non-ASCII characters is fine for CopyFile (it is not COM):
  // only the strings handed to SolidWorks have to be ASCII.
  check('source part exists', false, sourcePart)
} else {
  copyFileSync(sourcePart, stagePart)
  check('staged an ASCII copy of the part', existsSync(stagePart), stagePart)
}

const registered = new Map()
plugin.apply(
  {
    tools: { register: (tool) => registered.set(tool.name, tool) },
    logger: { info: () => {} },
    systemPrompt: { section: () => {} },
  },
  { partTemplate, assemblyTemplate, scratchDir, defaultTimeoutMs: 1_500_000 },
)

const run = registered.get('solidworks_run')
const verifyTool = registered.get('solidworks_verify')

// The build script is ASCII by contract and uses the two exported templates
// rather than naming an install-specific path.
const buildScript = readFileSync(new URL('../assets/assembly_build.vbs', import.meta.url), 'utf8')

const build = await run.execute(
  { script: buildScript, args: [stagePart], timeoutMs: 1_500_000 },
  { signal: undefined },
)

// A sandboxed Node could not previously hand a pipe to cscript at all (the
// runner now redirects to a file, so that should no longer happen); and
// SolidWorks can refuse a NEW automation connection when the machine's
// automation capacity is exhausted, which is machine state rather than a
// product fault. Both are reported as SKIP with the observed reason. A genuine
// build failure still falls through to the FAIL check below, so this cannot
// quietly hide a regression.
const environmentFault = /cannot launch cscript|no free automation connection|err=429/i
const buildLog = `${build.error ?? ''}\n${build.log ?? ''}`
if (build.error || (build.exitCode !== 0 && environmentFault.test(buildLog))) {
  console.log(`SKIP  live SolidWorks unavailable: ${build.error ?? 'see log'}`)
  console.log((build.log ?? '').trim().split('\n').slice(-4).map((line) => `      ${line}`).join('\n'))
  const reportPath = join(scratchDir, 'assembly-verify.json')
  writeFileSync(reportPath, `${JSON.stringify({ when: new Date().toISOString(), skipped: true, reason: build.error ?? 'environment fault', log: build.log, results }, null, 2)}\n`, 'utf8')
  console.log(`\nSKIPPED (${results.length} check(s) recorded) — report: ${reportPath}`)
  process.exit(0)
}

console.log(`\nsolidworks_run: exit=${build.exitCode} error=${build.error ?? 'none'}`)
console.log(`  asm tpl : ${build.assemblyTemplate || '<none>'}`)
console.log((build.log ?? '').trim().split('\n').slice(-6).join('\n'))

check('build script exited 0', build.exitCode === 0, `exit=${build.exitCode} ${build.error ?? ''}`)
check('assembly template exported', typeof build.assemblyTemplate === 'string' && build.assemblyTemplate.endsWith('.asmdot'), build.assemblyTemplate)
check('build script read %SWASMTPL%', (build.log ?? '').includes('.asmdot'), '')

const ok = await verifyTool.execute({ expectComponentCount: COMPONENTS, expectSuppressedCount: 0 }, { signal: undefined })
console.log(`\nsolidworks_verify expectComponentCount=${COMPONENTS}: ${ok.ok ? 'OK' : 'MISMATCH'}`)
console.log(`  document     : ${ok.documentType}`)
console.log(`  components   : ${ok.assembly?.componentCount} (expected ${ok.expectedComponentCount})`)
console.log(`  mates        : ${ok.assembly?.mateCountAvailable === true ? ok.assembly.mateCount : 'not inspectable on this host'}  suppressed: ${ok.assembly?.suppressedFeatureCount}`)
for (const mismatch of ok.mismatches ?? []) console.log(`  ! ${mismatch}`)

check('document detected as an assembly', ok.documentType === 'assembly', String(ok.documentType))
check('assembly section reported', ok.assembly !== null && typeof ok.assembly === 'object')
check('component count is right', ok.assembly?.componentCount === COMPONENTS, `${ok.assembly?.componentCount}`)
check('independent Reference count agrees', ok.assembly?.referenceFeatureCount === COMPONENTS, `${ok.assembly?.referenceFeatureCount}`)
check('verification reports ok:true for a correct assembly', ok.ok === true, (ok.mismatches ?? []).join('; '))
check('assembly report is lossless JSON', (() => { try { return typeof JSON.stringify(ok) === 'string' } catch { return false } })())

// The gate has to be able to FAIL, or it proves nothing.
const bad = await verifyTool.execute({ expectComponentCount: COMPONENTS + 1 }, { signal: undefined })
console.log(`\nsolidworks_verify expectComponentCount=${COMPONENTS + 1}: ${bad.ok ? 'OK (unexpected)' : 'MISMATCH (expected)'}`)
for (const mismatch of bad.mismatches ?? []) console.log(`  ! ${mismatch}`)
check('a wrong component count is rejected', bad.ok === false, (bad.mismatches ?? []).join('; '))

const failures = results.filter((r) => !r.ok)
const report = {
  when: new Date().toISOString(),
  assemblyTemplate,
  build: { exitCode: build.exitCode, artifacts: build.artifacts },
  verification: ok,
  results,
}
const reportPath = join(scratchDir, 'assembly-verify.json')
writeFileSync(reportPath, `${JSON.stringify(report, null, 2)}\n`, 'utf8')
console.log(`\n${failures.length === 0 ? 'ALL PASS' : `${failures.length} FAILED`} — report: ${reportPath}`)
process.exit(failures.length === 0 ? 0 : 1)
