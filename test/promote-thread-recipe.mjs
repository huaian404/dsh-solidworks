// Proves the promotion loop closes: build a thread with the core+swept-boss
// strategy through solidworks_run, then require solidworks_verify to accept it.
// This is the recipe the plugin routes to when capabilities.sweptCut is false.
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { apply } from '../lib/index.js'

const tools = new Map()
apply({
  tools: { register: (definition) => tools.set(definition.name, definition) },
  systemPrompt: { section: () => {} },
  logger: { info: () => {} },
}, { scratchDir: join(tmpdir(), 'dsh-solidworks', 'promotetest') })

const call = async (name, args) => {
  const tool = tools.get(name)
  const value = await tool.execute(args ?? {}, {})
  return { value, rendered: tool.output.render(args ?? {}, value).map((c) => c.text).join('\n') }
}

// --- the recipe under promotion: a trapezoidal thread ---------------------
const THREAD_RECIPE = `Option Explicit
Const CORE_R = 0.024
Const MAJ_R = 0.025
Const TIP_R = 0.026
Const TOTAL_L = 0.25
Const PIT = 0.005
Const REVS = 40
Const W_CREST = 0.00125
Const W_ROOT = 0.00225

Dim swApp, sw, ext, sm, fm, planes(2), f, n, cc, fe, fTh, ok
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

' 1. core cylinder along +Z (extruded circle: the axis is unambiguous)
ext.SelectByID2 planes(1), "PLANE", 0, 0, 0, False, 0, Nothing, 0
sw.InsertSketch2 True
Set cc = sm.CreateCircleByRadius(0.0, 0.0, 0.0, CORE_R)
sw.InsertSketch2 True
Set fe = fm.FeatureExtrusion2(True, False, False, 0, 0, TOTAL_L, 0.01, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
If IsEmpty(fe) Or TypeName(fe) = "Nothing" Then WScript.Echo "FATAL core" : WScript.Quit 0
fe.Name = "Core"
WScript.Echo "core ok"

' 2. helix from a MAJ_R base circle on the same plane
sw.ClearSelection2 True
ext.SelectByID2 planes(1), "PLANE", 0, 0, 0, False, 0, Nothing, 0
sw.InsertSketch2 True
Set cc = sm.CreateCircleByRadius(0.0, 0.0, 0.0, MAJ_R)
sw.InsertSketch2 True
Set f = sw.FeatureByPositionReverse(0)
f.Name = "HelixBase"
sw.ClearSelection2 True
ext.SelectByID2 "HelixBase", "SKETCH", 0, 0, 0, False, 0, Nothing, 0
sw.InsertHelix False, False, False, True, 0, 0.0, PIT, REVS, 0.0, 0.0
Set f = sw.FeatureByPositionReverse(0)
If f.GetTypeName2 <> "Helix" Then WScript.Echo "FATAL helix" : WScript.Quit 0
f.Name = "ThreadHelix"
WScript.Echo "helix ok"

' 3. trapezoidal profile on that plane, spanning the helix start point
sw.ClearSelection2 True
ext.SelectByID2 planes(1), "PLANE", 0, 0, 0, False, 0, Nothing, 0
sw.InsertSketch2 True
Call sm.CreateLine(TIP_R, -W_CREST, 0.0, CORE_R, -W_ROOT, 0.0)
Call sm.CreateLine(CORE_R, -W_ROOT, 0.0, CORE_R, W_ROOT, 0.0)
Call sm.CreateLine(CORE_R, W_ROOT, 0.0, TIP_R, W_CREST, 0.0)
Call sm.CreateLine(TIP_R, W_CREST, 0.0, TIP_R, -W_CREST, 0.0)
sw.InsertSketch2 True
Set f = sw.FeatureByPositionReverse(0)
f.Name = "ThreadProfile"
If f.GetSpecificFeature2.GetSketchContourCount < 1 Then WScript.Echo "FATAL profile" : WScript.Quit 0
WScript.Echo "profile ok"

' 4. sweep it as a BOSS with Merge = True (swept CUT is broken on this host)
sw.ClearSelection2 True
ext.SelectByID2 "ThreadProfile", "SKETCH", 0, 0, 0, False, 1, Nothing, 0
ext.SelectByID2 "ThreadHelix", "REFERENCECURVES", 0, 0, 0, True, 4, Nothing, 0
Set fTh = fm.InsertProtrusionSwept4(True, False, 0, False, False, 0, 0, False, 0.01, 0.01, 0, 1, True, False, False, 0.0, False, False, 0.01, 1)
If IsEmpty(fTh) Or TypeName(fTh) = "Nothing" Then WScript.Echo "FATAL sweep" : WScript.Quit 0
fTh.Name = "Thread"
WScript.Echo "thread ok"

sw.ClearSelection2 True
sw.ForceRebuild3 False
WScript.Quit 0
`

console.log('===== promoting the thread recipe =====')
const build = await call('solidworks_run', { script: THREAD_RECIPE })
console.log(build.rendered)

if (build.value?.exitCode !== 0) {
  console.error('BUILD FAILED')
  process.exit(1)
}

const verified = await call('solidworks_verify', {
  expectBodyCount: 1,
  featureTypes: ['Extrude', 'Helix', 'SweepBoss'],
  views: '7',
})
console.log('\n===== gate: solidworks_verify =====')
console.log(verified.rendered)

if (!verified.value?.ok) {
  console.error('\nPROMOTION REJECTED: the recipe did not produce the expected model')
  process.exit(1)
}

console.log('\nPROMOTION ACCEPTED: recipe is safe to freeze as a primitive')
console.log('renders:', (verified.value.renders ?? []).map((r) => r.bmp).join(', '))
