// Harness that drives the plugin's tools without the DSH host, to validate the
// plugin end-to-end against a real SolidWorks session.
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { apply } from '../lib/index.js'

const tools = new Map()
const prompts = []
const ctx = {
  tools: { register: (definition) => tools.set(definition.name, definition) },
  systemPrompt: { section: (section) => prompts.push(section) },
  logger: { info: (m) => console.log('[log]', m) },
}

apply(ctx, {
  scratchDir: join(tmpdir(), 'dsh-solidworks', 'plugintest'),
})

console.log('registered tools:', [...tools.keys()].join(', '))
console.log('prompt sections :', prompts.map((p) => p.name).join(', '))

const call = async (name, args) => {
  const tool = tools.get(name)
  if (!tool) throw new Error(`no tool ${name}`)
  const started = Date.now()
  const value = await tool.execute(args ?? {}, {})
  const rendered = tool.output.render(args ?? {}, value).map((c) => c.text).join('\n')
  return { value, rendered, ms: Date.now() - started }
}

// 1) capabilities
const caps = await call('solidworks_capabilities', { force: true })
console.log('\n===== solidworks_capabilities =====')
console.log(caps.rendered)
console.log(`(sweptCut=${caps.value?.capabilities?.sweptCut}, probed=${caps.value?.capabilities?.sweptCutProbed}, ${caps.ms}ms)`)

// 2) run a script that builds a disc and renders it, proving the escape hatch
const script = `Option Explicit
Dim swApp, sw, ext, sm, fm, planes(2), f, n, cc, fe, outDir
Set swApp = CreateObject("SldWorks.Application")
swApp.Visible = True
Do While swApp.GetDocumentCount > 0
    Set sw = swApp.ActiveDoc
    If IsEmpty(sw) Or TypeName(sw) = "Nothing" Then Exit Do
    swApp.CloseDoc sw.GetTitle
Loop
swApp.NewDocument CreateObject("WScript.Shell").ExpandEnvironmentStrings("%SWPARTTPL%"), 0, 0, 0
WScript.Sleep 2500
Set sw = swApp.ActiveDoc
Set ext = sw.Extension
Set sm = sw.SketchManager
Set fm = sw.FeatureManager
n = 0
Set f = sw.FirstFeature
Do While Not f Is Nothing
    If f.GetTypeName2 = "RefPlane" Then
        If n < 3 Then planes(n) = f.Name : n = n + 1
    End If
    Set f = f.GetNextFeature
Loop
WScript.Echo "planes ok"
ext.SelectByID2 planes(0), "PLANE", 0, 0, 0, False, 0, Nothing, 0
sw.InsertSketch2 True
Set cc = sm.CreateCircleByRadius(0.0, 0.0, 0.0, 0.05)
sw.InsertSketch2 True
Set fe = fm.FeatureExtrusion2(True, False, False, 0, 0, 0.02, 0.01, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
If IsEmpty(fe) Or TypeName(fe) = "Nothing" Then WScript.Echo "EXTRUDE FAILED" : WScript.Quit 0
fe.Name = "Disc"
WScript.Echo "disc built"
outDir = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%SWOUTDIR%")
WScript.Echo "bmp=" & sw.SaveBMP(outDir & "plugin_disc.bmp", 1400, 900)
WScript.Quit 0
`
const ran = await call('solidworks_run', { script })
console.log('\n===== solidworks_run =====')
console.log(ran.rendered)
console.log(`(exit=${ran.value?.exitCode}, ${ran.ms}ms)`)

// 3) verify against expectations
const verified = await call('solidworks_verify', { expectBodyCount: 1, featureTypes: ['Extrude'], views: '7' })
console.log('\n===== solidworks_verify =====')
console.log(verified.rendered)

// 4) verify with a deliberately wrong expectation, to prove mismatches surface
const bad = await call('solidworks_verify', { expectBodyCount: 3, featureTypes: ['SweepBoss'] })
console.log('\n===== solidworks_verify (intentionally wrong) =====')
console.log(bad.rendered)

// 5) non-ASCII script must be rejected with a clear reason
const nonAscii = await call('solidworks_run', { script: 'WScript.Echo "中文注释"' })
console.log('\n===== solidworks_run (non-ASCII) =====')
console.log(nonAscii.rendered)

console.log('\nALL TOOL CALLS COMPLETED')
