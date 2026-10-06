/**
 * rebuild-via-plugin.mjs — run the project's part build through the plugin's own
 * code path (spawn cscript, ANSI staging, %SWPARTTPL%/%SWOUTDIR%), then gate each
 * saved part with the real `solidworks_verify` implementation.
 *
 * This is the pre-reload proof: it exercises exactly the code the GUI will load
 * once the app is reloaded, so the only remaining difference is which revision
 * the host has mounted. `solidworks_run`'s output-contract fix is what makes the
 * run's log, exit code and artifacts observable here at all.
 *
 *   node test/rebuild-via-plugin.mjs [parts/build_final.vbs]
 */
import { readFileSync, mkdirSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { tmpdir } from 'node:os'

const plugin = await import('../lib/index.js')

const scratchDir = process.env.SW_TEST_SCRATCH ?? join(tmpdir(), 'dsh-solidworks', 'rebuild')
const partTemplate = process.env.SW_TEST_TEMPLATE ?? 'C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2026\\templates\\gb_part.prtdot'
const repoRoot = resolve(import.meta.dirname, '..', '..', '..', '..')
const scriptPath = process.argv[2] ?? join(repoRoot, 'parts', 'build_final.vbs')
const source = readFileSync(scriptPath, 'utf8')

const pluginLog = []
const registered = new Map()
plugin.apply(
  {
    tools: { register: (tool) => registered.set(tool.name, tool) },
    logger: { info: (message) => pluginLog.push(message) },
    systemPrompt: { section: () => {} },
  },
  { partTemplate, scratchDir, defaultTimeoutMs: 1_500_000 },
)

const run = registered.get('solidworks_run')
const verifyTool = registered.get('solidworks_verify')
if (run === undefined || verifyTool === undefined) {
  console.error(`tools did not register (got: ${[...registered.keys()].join(', ')})`)
  process.exit(1)
}

console.log(`build script : ${scriptPath}`)
for (const line of pluginLog.filter((l) => l.includes('[solidworks]'))) console.log(`plugin       : ${line}`)

const build = await run.execute({ script: source, timeoutMs: 1_500_000 }, { signal: undefined })
console.log(`\nsolidworks_run: exit=${build.exitCode} timedOut=${build.timedOut} error=${build.error ?? 'none'}`)
console.log(`template     : ${build.template || '(none)'}`)
console.log(`out dir      : ${build.dir}`)
console.log(`artifacts    : ${(build.artifacts ?? []).join(', ') || 'none'}`)
try {
  console.log(`json-safe    : ${JSON.stringify(build).length} bytes serialized, aborted=${build.aborted}`)
} catch (error) {
  console.error(`json-safe    : NOT SERIALIZABLE — ${error.message}`)
}
console.log('\n---------------- build log ----------------')
console.log((build.log ?? '').trim())

if (build.error || build.exitCode !== 0) {
  console.error('\nbuild failed')
  process.exit(1)
}

/**
 * Make one part the active document.
 *
 * The build already leaves every part loaded (it saved three documents without
 * closing them), so this activates by title rather than opening by path:
 * `OpenDoc6` returns a type mismatch / Nothing through this late binding, and a
 * file path is one more non-ASCII trap to avoid.
 */
async function activatePart(name) {
  // Notes from getting this to work:
  //   * ActivateDoc3 by bare title returns Nothing here ("slotted_disc"), and
  //     OpenDoc6 type-mismatches, so the opened documents are enumerated and the
  //     matching one is activated by its full path.
  //   * The elements of GetDocuments are OBJECTS: `docs(i).GetTitle` on an
  //     un-Set element is a type mismatch that On Error Resume Next swallows, so
  //     the loop silently reports zero documents. `Set d = docs(i)` is required.
  //   * After a fresh build only the LAST part is still open (the next NewDoc
  //     closes the previous ones), so the earlier parts are opened by name,
  //     building the path from %SWOUTDIR% to keep every literal ASCII.
  const stem = name.replace(/\.SLDPRT$/i, '')
  const activated = await run.execute(
    {
      script: `Option Explicit
Dim WSH, OUTDIR, swApp, docs, d, i, m, errs, hit, errs2
Set WSH = CreateObject("WScript.Shell")
OUTDIR = WSH.ExpandEnvironmentStrings("%SWOUTDIR%")
If Right(OUTDIR, 1) <> "\\" Then OUTDIR = OUTDIR & "\\"
Set swApp = CreateObject("SldWorks.Application")
On Error Resume Next
hit = ""
WScript.Echo "  probe: docCount=" & swApp.GetDocumentCount & " outDir=" & OUTDIR
docs = swApp.GetDocuments
If IsArray(docs) Then
    WScript.Echo "  probe: GetDocuments array " & LBound(docs) & ".." & UBound(docs)
    For i = LBound(docs) To UBound(docs)
        Set d = docs(i)
        WScript.Echo "  probe: [" & i & "] " & d.GetTitle & " | " & d.GetPathName
        If InStr(LCase(d.GetPathName), LCase("${stem}")) > 0 Then hit = d.GetPathName
    Next
Else
    WScript.Echo "  probe: GetDocuments -> " & TypeName(docs)
End If
WScript.Echo "  probe: hit=" & hit
If hit = "" Then
    Err.Clear
    Set m = swApp.OpenDoc6(OUTDIR & "${name}", 1, 0, "", errs2, errs)
    If m Is Nothing Then
        WScript.Echo "ACTIVATE FAILED: not open and OpenDoc6 returned Nothing errs=" & errs2
    Else
        WScript.Echo "OPENED " & m.GetTitle
    End If
Else
    Err.Clear
    Set m = swApp.ActivateDoc3(hit, False, 0, errs)
    If m Is Nothing Then
        WScript.Echo "ACTIVATE FAILED on " & hit & " errs=" & errs
    Else
        WScript.Echo "ACTIVATED " & m.GetTitle
    End If
End If
`,
      timeoutMs: 300_000,
    },
    { signal: undefined },
  )
  return { ok: /ACTIVATED/.test(activated.log ?? ''), log: (activated.log ?? '').trim() }
}

const gates = [
  { part: 'pulley_vbelt.SLDPRT', expectBodyCount: 1, featureTypes: ['Extrude', 'Revolve', 'Cut', 'Chamfer'] },
  { part: 'bracket_plate.SLDPRT', expectBodyCount: 1, featureTypes: ['Extrude', 'Cut', 'Pattern'] },
  { part: 'slotted_disc.SLDPRT', expectBodyCount: 1, featureTypes: ['Extrude', 'Cut'] },
]

const results = []
for (const gate of gates) {
  const opened = await activatePart(gate.part)
  const verification = opened.ok
    ? await verifyTool.execute({ expectBodyCount: gate.expectBodyCount, featureTypes: gate.featureTypes }, { signal: undefined })
    : { ok: false, error: opened.log, mismatches: ['document did not open'] }
  results.push({ part: gate.part, opened: opened.ok, verification })
  console.log(`\n=== ${gate.part} ===`)
  console.log(`  ${verification.ok ? 'OK' : 'MISMATCH'}  bodies=${verification.bodyCount}  kinds=[${(verification.featureKinds ?? []).join(', ')}]`)
  console.log(`  renders: ${(verification.renders ?? []).map((r) => r.bmp).join(', ') || 'none'}`)
  for (const mismatch of verification.mismatches ?? []) console.log(`  ! ${mismatch}`)
  if (verification.error) console.log(`  ! ${verification.error}`)
}

const failed = results.filter((r) => r.verification.ok !== true)
const report = {
  when: new Date().toISOString(),
  pluginRevision: 'recipes-2 (workspace lib, pre-reload)',
  script: scriptPath,
  exitCode: build.exitCode,
  artifacts: build.artifacts,
  results,
}
mkdirSync(scratchDir, { recursive: true })
const reportPath = join(scratchDir, 'rebuild-via-plugin.json')
writeFileSync(reportPath, `${JSON.stringify(report, null, 2)}\n`, 'utf8')
console.log(`\n${failed.length === 0 ? 'ALL PARTS VERIFIED' : `${failed.length} PART(S) FAILED VERIFY`} — report: ${reportPath}`)
process.exit(failed.length === 0 ? 0 : 1)
