/**
 * verify-active.mjs — run one build script through the plugin, then gate the
 * still-active document with the plugin's real `solidworks_verify` code.
 *
 * Why only the last document: the build ends with its final part active, and
 * re-opening the earlier saved parts fails on this host (OpenDoc6 -> Nothing,
 * and getPathName comes back Empty), so the per-part gates that need a document
 * switch are done by `solidworks_recipe`, whose build + verify run inside one
 * script. This covers the path that matters here: build -> verify, end to end,
 * with the fixed `solidworks_run` payload.
 *
 *   node test/verify-active.mjs <build.vbs> <expectBodyCount> [kind,kind,...]
 */
import { readFileSync, mkdirSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { tmpdir } from 'node:os'

const plugin = await import('../lib/index.js')

const scriptPath = process.argv[2]
const expectBodyCount = Number(process.argv[3] ?? '1')
const featureTypes = (process.argv[4] ?? '').split(',').map((s) => s.trim()).filter(Boolean)
if (!scriptPath) {
  console.error('usage: node test/verify-active.mjs <build.vbs> <expectBodyCount> [kind,kind,...]')
  process.exit(2)
}

const scratchDir = process.env.SW_TEST_SCRATCH ?? join(tmpdir(), 'dsh-solidworks', 'verifyactive')
const partTemplate = process.env.SW_TEST_TEMPLATE ?? 'C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2026\\templates\\gb_part.prtdot'
const source = readFileSync(resolve(scriptPath), 'utf8')

const registered = new Map()
plugin.apply(
  {
    tools: { register: (tool) => registered.set(tool.name, tool) },
    logger: { info: () => {} },
    systemPrompt: { section: () => {} },
  },
  { partTemplate, scratchDir, defaultTimeoutMs: 1_500_000 },
)

const run = registered.get('solidworks_run')
const verifyTool = registered.get('solidworks_verify')

console.log(`script: ${scriptPath}`)
const build = await run.execute({ script: source, timeoutMs: 1_500_000 }, { signal: undefined })
console.log(`solidworks_run: exit=${build.exitCode} error=${build.error ?? 'none'} artifacts=[${(build.artifacts ?? []).join(', ')}]`)
console.log((build.log ?? '').trim().split('\n').slice(-3).join('\n'))

const verification = await verifyTool.execute({ expectBodyCount, featureTypes }, { signal: undefined })
console.log(`\nsolidworks_verify: ${verification.ok ? 'OK' : 'MISMATCH'}`)
console.log(`  title        : ${verification.title}`)
console.log(`  solid bodies : ${verification.bodyCount} (expected ${expectBodyCount})`)
console.log(`  feature kinds: [${(verification.featureKinds ?? []).join(', ')}]`)
console.log(`  renders      : ${(verification.renders ?? []).map((r) => r.bmp).join(', ') || 'none'}`)
for (const mismatch of verification.mismatches ?? []) console.log(`  ! ${mismatch}`)

mkdirSync(scratchDir, { recursive: true })
const reportPath = join(scratchDir, 'verify-active.json')
writeFileSync(reportPath, `${JSON.stringify({ when: new Date().toISOString(), script: scriptPath, build: { exitCode: build.exitCode, artifacts: build.artifacts }, verification }, null, 2)}\n`, 'utf8')
console.log(`\nreport: ${reportPath}`)
process.exit(verification.ok === true ? 0 : 1)
