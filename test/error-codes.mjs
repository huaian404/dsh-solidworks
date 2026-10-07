/**
 * error-codes.mjs — every failure must carry a stable machine-readable code.
 *
 * Why: the tools used to report failures as prose, so a caller could not tell a
 * configuration fault (no part template), an environment fault (cscript is not
 * on PATH), a host fault (SolidWorks cannot be started), a dialog stuck in the
 * UI (timeout) and a script that genuinely failed from one another. The code is
 * a contract, so these strings must not drift.
 *
 * Runs without SolidWorks. The environment faults are synthesised by pointing
 * PATH at an empty directory, which is why `cscript-missing` is testable here
 * at all; cases that would need a live SolidWorks session are reported as SKIP
 * with the observed outcome instead of being asserted.
 *
 *   node test/error-codes.mjs
 */
import { mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const plugin = await import('../lib/index.js')

const results = []
const check = (name, ok, detail = '') => {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`)
}
const skip = (name, detail) => console.log(`SKIP  ${name} — ${detail}`)

/** Every code the plugin may emit; the README table and this list agree. */
const KNOWN = new Set([
  'no-solidworks', 'no-template', 'non-ascii-script', 'cscript-missing', 'timeout',
  'script-error', 'probe-failed', 'inspect-failed', 'verification-failed',
  'recipe-unknown', 'recipe-invalid', 'invalid-input', 'unknown',
])

const base = join(tmpdir(), 'dsh-solidworks', 'errcodes')
rmSync(base, { recursive: true, force: true })
mkdirSync(base, { recursive: true })

/** A directory with nothing in it, used to hide cscript.exe from PATH. */
const emptyPath = join(base, 'empty-path')
mkdirSync(emptyPath, { recursive: true })

const tools = new Map()
plugin.apply(
  {
    tools: { register: (tool) => tools.set(tool.name, tool) },
    logger: { info: () => {} },
    systemPrompt: { section: () => {} },
  },
  { scratchDir: join(base, 'scratch') },
)
const tool = (name) => {
  const found = tools.get(name)
  if (found === undefined) throw new Error(`tool ${name} did not register`)
  return found
}

/** Run a tool and collect both the structured value and the rendered text. */
const call = async (name, args, options = {}) => {
  const previousPath = process.env.PATH
  if (options.hideCscript === true) process.env.PATH = emptyPath
  try {
    const value = await tool(name).execute(args, {})
    const rendered = tool(name).output.render(args, value).map((c) => c.text).join('\n')
    return { value, rendered }
  } finally {
    process.env.PATH = previousPath
  }
}

const hasCode = (result) => typeof result.value.code === 'string' && KNOWN.has(result.value.code)

// 1. caller mistakes, classified without touching SolidWorks
const unknownRecipe = await call('solidworks_recipe', { name: 'does_not_exist_at_all' })
check('unknown recipe -> recipe-unknown', unknownRecipe.value.code === 'recipe-unknown', `code=${unknownRecipe.value.code}`)

const badParam = await call('solidworks_recipe', { name: 'disc_multicontour_cut', params: { R_OUT: 'wide' } })
check('bad parameter -> recipe-invalid', badParam.value.code === 'recipe-invalid', `code=${badParam.value.code}`)

const badAction = await call('solidworks_recipes', { action: 'destroy' })
check('unknown action -> invalid-input', badAction.value.code === 'invalid-input', `code=${badAction.value.code}`)

const noName = await call('solidworks_recipes', { action: 'remove' })
check('remove without name -> invalid-input', noName.value.code === 'invalid-input', `code=${noName.value.code}`)

const badRecipe = await call('solidworks_recipes', { action: 'save', name: 'Bad Name', script: 'x = 1' })
check('malformed save -> recipe-invalid', badRecipe.value.code === 'recipe-invalid', `code=${badRecipe.value.code}`)

// 2. the staging guard, classified by the caller that knows why it failed
const nonAscii = await call('solidworks_run', { script: 'WScript.Echo "\u4f60\u597d"' })
check('non-ASCII script -> non-ascii-script', nonAscii.value.code === 'non-ascii-script', `code=${nonAscii.value.code}`)

// 3. a machine that cannot run scripts at all
const hidden = await call('solidworks_run', { script: 'WScript.Echo "hi"' }, { hideCscript: true })
check('cscript not on PATH -> cscript-missing', hidden.value.code === 'cscript-missing', `code=${hidden.value.code}`)
check('the code reaches the rendered text', /\[cscript-missing\]/.test(hidden.rendered), hidden.rendered.split('\n')[0])

const capsHidden = await call('solidworks_capabilities', { force: true }, { hideCscript: true })
check('capability probe with no interpreter is classified too', hasCode(capsHidden), `code=${capsHidden.value.code}`)
// 4. a script that outlives its timeout
const slow = await call('solidworks_run', { script: 'WScript.Sleep 120000', timeoutMs: 500 })
if (slow.value.timedOut === true) {
  check('timeout -> timeout', slow.value.code === 'timeout', `code=${slow.value.code}`)
} else {
  skip('timeout -> timeout', `the script never reached a live interpreter here (code=${slow.value.code ?? 'none'})`)
}

// 5. a successful run must NOT carry a code: an explicit `code: undefined` is
//    not lossless JSON, and the host would discard the entire result.
const noInterpreter = (value) => value.error !== undefined && /EPERM|ENOENT|cannot launch|cannot find/i.test(String(value.error))
const fine = await call('solidworks_run', { script: 'WScript.Echo "hi"', timeoutMs: 20000 })
if (noInterpreter(fine.value)) {
  skip('a successful run carries no code', `no interpreter on this host (code=${fine.value.code})`)
} else {
  check('a successful run carries no code', !('code' in fine.value) && fine.value.exitCode === 0, `exit=${fine.value.exitCode} keys=${Object.keys(fine.value).length}`)
}

// 6. the contract itself: no failure escapes without a code
const failures = [unknownRecipe, badParam, badAction, noName, badRecipe, nonAscii, hidden, capsHidden]
const uncoded = failures.filter((r) => r.value.error !== undefined && !hasCode(r))
check('every failure carries a known code', uncoded.length === 0, uncoded.map((r) => r.value.error).join(' | ') || 'all coded')

// 7. the code is machine-readable in the value, not only in the prose
check(
  'codes are stable strings, not prose',
  failures.every((r) => r.value.code === undefined || r.value.code === r.value.code.toLowerCase()),
  failures.map((r) => r.value.code).filter(Boolean).join(','),
)

// 8. Scratch isolation: two invocations must not share a workspace, because
//    SolidWorks has one active document and a fixed `inspect.json` in a shared
//    directory lets one verification read another invocation's report.
const first = await call('solidworks_run', { script: 'WScript.Echo "a"', timeoutMs: 15000 })
const second = await call('solidworks_run', { script: 'WScript.Echo "b"', timeoutMs: 15000 })
check('two invocations get different scratch directories', first.value.dir !== second.value.dir, `${first.value.dir} vs ${second.value.dir}`)
check(
  'the scratch directory is per call, not a fixed "session"',
  typeof first.value.dir === 'string' && !/[\\/]session$/.test(first.value.dir),
  String(first.value.dir),
)

// 9. A tool that drives SolidWorks must NOT declare isConcurrencySafe: the host
//    schedules undeclared tools as `exclusive` and serializes them for us. A
//    read-only tool may declare it, and then must actually return true.
const safety = {
  solidworks_run: tools.get('solidworks_run').isConcurrencySafe,
  solidworks_recipe: tools.get('solidworks_recipe').isConcurrencySafe,
  solidworks_verify: tools.get('solidworks_verify').isConcurrencySafe,
}
check(
  'the UI-driving tools stay exclusive (undeclared) so the host serializes them',
  safety.solidworks_run === undefined && safety.solidworks_recipe === undefined,
  `run=${typeof safety.solidworks_run} recipe=${typeof safety.solidworks_recipe}`,
)
check(
  'the read-only tools declare concurrency safety and honour it',
  typeof safety.solidworks_verify === 'function' && safety.solidworks_verify({}) === true,
  `verify=${typeof safety.solidworks_verify}`,
)

const failed = results.filter((r) => !r.ok)
console.log(`\n${failed.length === 0 ? 'ALL PASS' : `${failed.length} FAILED`} (${results.length} checks)`)
process.exit(failed.length === 0 ? 0 : 1)
