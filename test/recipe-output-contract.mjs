/**
 * recipe-output-contract.mjs — find non-lossless-JSON values in the
 * `solidworks_recipe` and `solidworks_verify` return payloads.
 *
 * `solidworks_run` had this bug (an AbortSignal in the result, fixed in
 * recipes-2). `solidworks_recipe` then reproduced the same symptom in the live
 * session even though the build itself succeeded. This walks the real result
 * object and reports every path whose value cannot survive JSON round-tripping,
 * so the offending field is named instead of guessed.
 *
 *   node test/recipe-output-contract.mjs
 */
import { mkdirSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'

const plugin = await import('../lib/index.js')

const scratchDir = process.env.SW_TEST_SCRATCH ?? join(tmpdir(), 'dsh-solidworks', 'recipecontract')
const partTemplate = process.env.SW_TEST_TEMPLATE ?? 'C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2026\\templates\\gb_part.prtdot'
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

/** Walk a value and report every path JSON cannot represent losslessly. */
function inspect(value, path = '$', seen = new Set(), out = []) {
  const type = typeof value
  if (value === undefined) {
    out.push({ path, problem: 'undefined' })
    return out
  }
  if (value === null) return out
  if (type === 'number') {
    if (!Number.isFinite(value)) out.push({ path, problem: `non-finite number (${value})` })
    return out
  }
  if (type === 'bigint') {
    out.push({ path, problem: `bigint (${value})` })
    return out
  }
  if (type === 'function' || type === 'symbol') {
    out.push({ path, problem: type })
    return out
  }
  if (type !== 'object') return out
  if (seen.has(value)) {
    out.push({ path, problem: 'circular reference' })
    return out
  }
  seen.add(value)
  if (Array.isArray(value)) {
    value.forEach((entry, index) => inspect(entry, `${path}[${index}]`, seen, out))
    return out
  }
  for (const [key, entry] of Object.entries(value)) inspect(entry, `${path}.${key}`, seen, out)
  return out
}

const recipeTool = registered.get('solidworks_recipe')

// Gate EVERY registered tool, not just the recipe path: an `undefined` anywhere
// in a result makes the host reject the whole payload with "value is not
// lossless JSON", and the tool that hits it first is whichever one runs next.
const failures = []
for (const [name, tool] of registered) {
  const checks = []
  // a cheap, side-effect-free probe per tool where one exists
  if (name === 'solidworks_recipes') checks.push({ args: { action: 'list' } })
  if (name === 'solidworks_capabilities') checks.push({ args: { force: false } })
  for (const check of checks) {
    let value
    try {
      value = await tool.execute(check.args, { signal: undefined })
    } catch (error) {
      failures.push(`${name}: execute threw ${error.message}`)
      continue
    }
    const problems = inspect(value)
    if (problems.length > 0) {
      for (const p of problems) failures.push(`${name}${p.path.slice(1)}: ${p.problem}`)
    }
    try {
      tool.output.render(check.args, value)
    } catch (error) {
      failures.push(`${name}: render threw ${error.message}`)
    }
    console.log(`${name} {${JSON.stringify(check.args)}}: ${problems.length === 0 ? 'lossless JSON' : `${problems.length} problem(s)`}`)
  }
}

// The recipe build path itself (this one actually builds a part).
const value = await recipeTool.execute({ name: 'disc_multicontour_cut', views: '7' }, { signal: undefined })

console.log('\n=== disc_multicontour_cut result ===')
console.log(`ok=${value.ok} exitCode=${value.exitCode} error=${value.error ?? 'none'}`)
console.log(`verification.ok=${value.verification?.ok} bodies=${value.verification?.bodyCount} kinds=[${(value.verification?.featureKinds ?? []).join(', ')}]`)

const problems = inspect(value)
for (const p of problems) failures.push(`solidworks_recipe${p.path.slice(1)}: ${p.problem}`)
console.log(problems.length === 0 ? 'result is lossless JSON' : `${problems.length} non-lossless path(s)`)

// render is part of the contract too: the host renders before it serializes
try {
  const rendered = recipeTool.output.render({ name: 'disc_multicontour_cut' }, value)
  console.log(`render ok: ${rendered.length} part(s), ${rendered[0].text.length} chars`)
} catch (error) {
  failures.push(`solidworks_recipe: render threw ${error.message}`)
}

if (failures.length > 0) {
  console.log(`\n${failures.length} FAILURE(S):`)
  for (const failure of failures) console.log(`  ! ${failure}`)
} else {
  console.log('\nALL TOOL PAYLOADS ARE LOSSLESS JSON')
}
process.exit(failures.length === 0 ? 0 : 1)
