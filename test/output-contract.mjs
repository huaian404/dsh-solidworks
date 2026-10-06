/**
 * output-contract.mjs — regression test for the `solidworks_run` return payload.
 *
 * The bug this locks down: the tool returned `abortSignal: exec?.signal`, an
 * object (or `undefined`). The host rejects a non-lossless-JSON result, so the
 * ENTIRE result — log, artifacts, exit code — was replaced by
 * "tool solidworks_run returned invalid output: value is not lossless JSON",
 * while the script itself had actually run and saved its part. A model driving
 * SolidWorks was therefore blind: no log, no artifacts, no way to tell a build
 * failure from a serialization failure.
 *
 * The test drives the real `apply()` from lib/index.js with a mock tool
 * registry, runs a trivial script against live SolidWorks, and asserts the
 * result survives a JSON round trip. It needs no host and no GUI reload, so it
 * is the fast seam for iterating on this plugin.
 *
 *   node test/output-contract.mjs
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'

const plugin = await import('../lib/index.js')

const scratchDir = process.env.SW_TEST_SCRATCH ?? join(tmpdir(), 'dsh-solidworks', 'contracttest')
const partTemplate = process.env.SW_TEST_TEMPLATE ?? 'C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2026\\templates\\gb_part.prtdot'

const registered = new Map()
const ctx = {
  tools: { register: (tool) => registered.set(tool.name, tool) },
  logger: { info: () => {} },
  systemPrompt: { section: () => {} },
}
mkdirSync(scratchDir, { recursive: true })

const results = []
const check = (name, ok, detail = '') => {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`)
}

// A result is only usable if it survives the host's serialization gate.
const lossless = (value) => {
  try {
    const text = JSON.stringify(value)
    if (text === undefined) return { ok: false, reason: 'not serializable' }
    return { ok: true, text }
  } catch (error) {
    return { ok: false, reason: error.message }
  }
}

plugin.apply(ctx, { partTemplate, scratchDir, defaultTimeoutMs: 600_000 })

check('tools registered', registered.size >= 5, [...registered.keys()].join(', '))

const run = registered.get('solidworks_run')
if (run === undefined) {
  check('solidworks_run present', false)
} else {
  const value = await run.execute(
    {
      script: [
        'WScript.Echo "solidworks_run output contract probe"',
        'Dim swApp',
        'Set swApp = CreateObject("SldWorks.Application")',
        'swApp.Visible = True',
        'WScript.Echo "revision = " & swApp.RevisionNumber',
        'WScript.Echo "docs open = " & swApp.GetDocumentCount',
      ].join('\r\n'),
      timeoutMs: 600_000,
    },
    { signal: undefined },
  )

  const json = lossless(value)
  check('result is lossless JSON', json.ok, json.ok ? `${json.text.length} bytes` : json.reason)
  check('no AbortSignal in the payload', value?.abortSignal === undefined)
  check('aborted flag is a JSON boolean', value?.aborted === undefined || typeof value.aborted === 'boolean', `aborted=${value?.aborted}`)
  check('exit code reported', value?.exitCode === 0, `exitCode=${value?.exitCode}`)
  check('script log reaches the caller', typeof value?.log === 'string' && value.log.includes('output contract probe'))
  check('template resolved', typeof value?.template === 'string' && value.template.length > 0, value?.template)
  check('artifacts listed as an array', Array.isArray(value?.artifacts), `[${value?.artifacts ?? ''}]`)
  check('probe caught the SolidWorks revision', typeof value?.log === 'string' && /revision = \d/.test(value.log))
}

const failures = results.filter((r) => !r.ok)
const report = { when: new Date().toISOString(), pass: results.length - failures.length, fail: failures.length, results }
const reportPath = join(scratchDir, 'output-contract.json')
writeFileSync(reportPath, `${JSON.stringify(report, null, 2)}\n`, 'utf8')
console.log(`\n${failures.length === 0 ? 'ALL PASS' : `${failures.length} FAILED`} — report: ${reportPath}`)
process.exit(failures.length === 0 ? 0 : 1)
