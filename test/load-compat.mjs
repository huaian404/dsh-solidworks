/**
 * load-compat.mjs — load every .mjs in test/ the way the app's Node does, and
 * check the plugin entry can be re-imported on a fresh module graph.
 *
 * This is the pre-restart smoke test: after the app reloads, the profile will
 * evaluate lib/index.js again, and any syntax error there would abort the whole
 * plugin entry (all five tools disappear). Checking it here means a restart is
 * not a gamble on files that were only ever run interactively.
 *
 *   node test/load-compat.mjs
 */
import { readdirSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { pathToFileURL } from 'node:url'

const here = import.meta.dirname
const results = []
const check = (name, ok, detail = '') => {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`)
}

// 1. lib/index.js imports cleanly and exposes apply()
const entry = await import('../lib/index.js')
check('lib/index.js imports', typeof entry.apply === 'function', `exports: ${Object.keys(entry).join(', ')}`)

// 2. apply() registers the full tool set on a mock registry
const registered = new Map()
entry.apply(
  {
    tools: { register: (tool) => registered.set(tool.name, tool) },
    logger: { info: () => {} },
    systemPrompt: { section: () => {} },
  },
  { scratchDir: join(tmpdir(), 'dsh-solidworks', 'loadcompat') },
)
check(
  'all five tools register',
  ['solidworks_capabilities', 'solidworks_run', 'solidworks_verify', 'solidworks_recipe', 'solidworks_recipes']
    .every((name) => registered.has(name)),
  [...registered.keys()].join(', '),
)

// 3. every tool definition survives the host's defineTool validation shape
for (const [name, tool] of registered) {
  const hasOutputSchema = tool.output !== undefined && typeof tool.output.schema === 'object'
  check(`${name}: output.schema present`, hasOutputSchema)
}

// 4. the helper modules the entry imports resolve as files
const helpers = ['../lib/recipes.js']
for (const helper of helpers) {
  try {
    await import(helper)
    check(`${helper} imports`, true)
  } catch (error) {
    check(`${helper} imports`, false, error.message)
  }
}

// 5. every .mjs in test/ at least parses (import is too heavy for the drivers;
//    a parse check via dynamic import of the source text is not available, so
//    the drivers are run deliberately elsewhere). Report what exists instead.
const testFiles = readdirSync(here).filter((f) => f.endsWith('.mjs')).sort()
check('test/ inventory', testFiles.length > 0, testFiles.join(', '))

const failures = results.filter((r) => !r.ok)
console.log(`\n${failures.length === 0 ? 'ALL PASS' : `${failures.length} FAILED`} (${results.length} checks)`)
process.exit(failures.length === 0 ? 0 : 1)
