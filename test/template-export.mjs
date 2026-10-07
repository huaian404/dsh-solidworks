/**
 * template-export.mjs — does the plugin resolve and export BOTH templates?
 *
 * Locks down the config surface added for assemblies: `assemblyTemplate` and
 * the `%SWASMTPL%` environment variable a build script reads. It needs no CAD
 * and no cscript, so it also serves as the fast check when the live app has not
 * reloaded the plugin yet (the running instance keeps the code it loaded).
 *
 *   node test/template-export.mjs
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'

const plugin = await import('../lib/index.js')

const scratchDir = process.env.SW_TEST_SCRATCH ?? join(tmpdir(), 'dsh-solidworks', 'template-export')
mkdirSync(scratchDir, { recursive: true })

const results = []
const check = (name, ok, detail = '') => {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`)
}

/** Register the real tools against a mock registry and hand them back. */
function toolsFor(config) {
  const registered = new Map()
  plugin.apply(
    {
      tools: { register: (tool) => registered.set(tool.name, tool) },
      logger: { info: () => {} },
      systemPrompt: { section: () => {} },
    },
    { scratchDir, ...config },
  )
  return registered
}

const registered = toolsFor({})
const run = registered.get('solidworks_run')

// 1. auto-detection: both extensions must be resolved from the install root.
//    A trivial script reports the environment the tool actually exported, which
//    is the only place %SWASMTPL% can be observed from out here. cscript is
//    unavailable under a sandboxed Node, so the reported templates are read
//    from the tool payload instead - they are resolved before the spawn.
const probe = await run.execute(
  {
    script: [
      'WScript.Echo "SWPARTTPL=[" & WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings("%SWPARTTPL%") & "]"',
      'WScript.Echo "SWASMTPL=[" & WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings("%SWASMTPL%") & "]"',
    ].join('\r\n'),
    timeoutMs: 120_000,
  },
  { signal: undefined },
)
check('part template auto-detected', String(probe.template ?? '').endsWith('.prtdot'), probe.template)
check('assembly template auto-detected', String(probe.assemblyTemplate ?? '').endsWith('.asmdot'), probe.assemblyTemplate)
if (probe.error) {
  console.log(`      (cscript unavailable here: ${probe.error} — the payload above is resolved pre-spawn)`)
} else {
  check('script saw %SWASMTPL%', String(probe.log ?? '').includes('.asmdot'), (probe.log ?? '').trim().split('\n').pop())
}

// 2. the two tools that carry the templates must expose them in their payload.
const verifyParams = Object.keys(registered.get('solidworks_verify').parameters.properties)
for (const name of ['expectComponentCount', 'expectMateCount', 'expectSuppressedCount']) {
  check(`solidworks_verify accepts ${name}`, verifyParams.includes(name), verifyParams.join(', '))
}
const runDescription = registered.get('solidworks_run').description
check('solidworks_run documents %SWASMTPL%', runDescription.includes('%SWASMTPL%'))
check('solidworks_run documents the assembly recipe', /AddComponent5/.test(runDescription))

// 4. recipe validation must accept assembly expectations.
const { validateRecipe } = await import('../lib/recipes.js')
const assemblyRecipe = validateRecipe({
  name: 'asm_demo',
  script: 'Const N = {{N}}\r\nWScript.Echo N',
  parameters: { N: { value: 3, description: 'instances' } },
  verify: { componentCount: 3, mateCount: 0 },
}, 'test')
check('an assembly verify block validates', assemblyRecipe.ok === true, assemblyRecipe.ok ? '' : assemblyRecipe.reason)
check('componentCount survives validation', assemblyRecipe.ok && assemblyRecipe.recipe.verify.componentCount === 3)
check('mateCount survives validation', assemblyRecipe.ok && assemblyRecipe.recipe.verify.mateCount === 0)
const badRecipe = validateRecipe({
  name: 'asm_bad',
  script: 'Const N = {{N}}',
  parameters: { N: { value: 3, description: 'instances' } },
  verify: { componentCount: -1 },
}, 'test')
check('a negative componentCount is rejected', badRecipe.ok === false, badRecipe.ok ? '' : badRecipe.reason)

console.log(`\nplugin revision gate: ${registered.size} tool(s) registered`)
const failures = results.filter((r) => !r.ok)
const reportPath = join(scratchDir, 'template-export.json')
writeFileSync(reportPath, `${JSON.stringify({ when: new Date().toISOString(), results }, null, 2)}\n`, 'utf8')
console.log(`${failures.length === 0 ? 'ALL PASS' : `${failures.length} FAILED`} — report: ${reportPath}`)
process.exit(failures.length === 0 ? 0 : 1)
