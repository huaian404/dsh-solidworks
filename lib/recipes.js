/**
 * Recipe storage — the data half of the plugin.
 *
 * A recipe is a validated build script plus the shape it must produce. It is
 * stored as JSON, never as code, so a new recipe is written with a tool call
 * instead of a source edit. That distinction is the whole point:
 *
 *   - a broken recipe fails only its own substitution and is skipped with a
 *     warning; `apply()` still returns and the plugin's three stable tools
 *     keep working;
 *   - a broken source edit aborts the entire plugin entry at composition time,
 *     which is exactly how this plugin was once un-enableable.
 *
 * Recipes are immutable-by-construction: every substitution is validated
 * before the script runs.
 *
 * @module dsh-solidworks/recipes
 */

import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

/** Recipe names become file names and parameter labels, so keep them strict. */
export const RECIPE_NAME = /^[a-z0-9][a-z0-9_-]{0,30}$/
/** Parameter placeholders look like `%CORE_R%` in the script body. */
const PLACEHOLDER = /\{\{([A-Za-z_][A-Za-z0-9_]*)\}\}/g
const MAX_SCRIPT_BYTES = 64 * 1024

/**
 * Ships with the plugin. The script is parameterised: a parameter's `value` is
 * spliced into the script wherever `%NAME%` appears, so the body must be
 * coefficient-free — every numeric literal lives in the parameter defaults.
 *
 * The thread recipe is the one this plugin routes to when the swept-cut route
 * is unavailable: a core cylinder at the minor diameter plus a merged swept
 * boss. It scored bodyCount 1 / [Extrude, Helix, SweepBoss] in
 * test/promote-thread-recipe.mjs.
 */
export const SHIPPED_RECIPES = [
  {
    name: 'threaded_shaft',
    description: 'Trapezoidal threaded shaft (leadscrew): core cylinder + helix + merged swept boss. Builds one solid body.',
    parameters: {
      CORE_R: { value: 0.024, description: 'core (minor) radius in metres, e.g. 0.024 = d48' },
      MAJ_R: { value: 0.025, description: 'helix / major radius in metres, e.g. 0.025 = d50' },
      TIP_R: { value: 0.026, description: 'profile outer radius; must exceed CORE_R' },
      TOTAL_L: { value: 0.25, description: 'total shaft length in metres' },
      PIT: { value: 0.005, description: 'thread pitch in metres' },
      REVS: { value: 40, description: 'helix revolutions; PIT x REVS = threaded length' },
      W_CREST: { value: 0.00125, description: 'tangential half-width at the crest' },
      W_ROOT: { value: 0.00225, description: 'tangential half-width at the root' },
    },
    verify: { bodyCount: 1, featureTypes: ['Extrude', 'Helix', 'SweepBoss'] },
    script: `Option Explicit
Const CORE_R = {{CORE_R}}
Const MAJ_R = {{MAJ_R}}
Const TIP_R = {{TIP_R}}
Const TOTAL_L = {{TOTAL_L}}
Const PIT = {{PIT}}
Const REVS = {{REVS}}
Const W_CREST = {{W_CREST}}
Const W_ROOT = {{W_ROOT}}

Dim swApp, sw, ext, sm, fm, planes(2), f, n, cc, fe, fTh, guard, act
Set swApp = CreateObject("SldWorks.Application")
swApp.Visible = True
guard = 0
Do While swApp.GetDocumentCount > 0 And guard < 30
    Set act = swApp.ActiveDoc
    If IsEmpty(act) Or TypeName(act) = "Nothing" Then Exit Do
    swApp.CloseDoc act.GetTitle
    guard = guard + 1
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
WScript.Echo "planes = [" & planes(0) & "] [" & planes(1) & "] [" & planes(2) & "]"

' 1. core cylinder along +Z (extruded circle: the axis is unambiguous)
ext.SelectByID2 planes(1), "PLANE", 0, 0, 0, False, 0, Nothing, 0
sw.InsertSketch2 True
Set cc = sm.CreateCircleByRadius(0.0, 0.0, 0.0, CORE_R)
If IsEmpty(cc) Or TypeName(cc) = "Nothing" Then WScript.Echo "FATAL core circle" : WScript.Quit 0
sw.InsertSketch2 True
Set fe = fm.FeatureExtrusion2(True, False, False, 0, 0, TOTAL_L, 0.01, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
If IsEmpty(fe) Or TypeName(fe) = "Nothing" Then WScript.Echo "FATAL core extrude" : WScript.Quit 0
fe.Name = "Core"
WScript.Echo "core ok: d" & CORE_R * 2000 & " x " & TOTAL_L * 1000 & " mm"

' 2. helix from a base circle on the same plane, so it is coaxial with the core
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
WScript.Echo "helix ok: R" & MAJ_R * 1000 & " pitch " & PIT * 1000 & " x " & REVS

' 3. trapezoidal profile spanning the helix start point
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

' 4. sweep as a BOSS with Merge = True: the swept CUT route is broken on this host
sw.ClearSelection2 True
ext.SelectByID2 "ThreadProfile", "SKETCH", 0, 0, 0, False, 1, Nothing, 0
ext.SelectByID2 "ThreadHelix", "REFERENCECURVES", 0, 0, 0, True, 4, Nothing, 0
Set fTh = fm.InsertProtrusionSwept4(True, False, 0, False, False, 0, 0, False, 0.01, 0.01, 0, 1, True, False, False, 0.0, False, False, 0.01, 1)
If IsEmpty(fTh) Or TypeName(fTh) = "Nothing" Then WScript.Echo "FATAL sweep" : WScript.Quit 0
fTh.Name = "Thread"
WScript.Echo "thread ok"

sw.ClearSelection2 True
sw.ForceRebuild3 False
WScript.Echo "build complete"
`,
  },
{
    name: "plate_linear_holes",
    description: "Mounting plate with a linear pattern of through holes (extrude + blind cut from the +Z face + FeatureLinearPattern4).",
    parameters: {
      PLATE_X: { value: 0.065, description: "plate length along X (m)" },
      PLATE_Y: { value: 0.05, description: "plate width along Y (m)" },
      THK: { value: 0.012, description: "plate thickness (m)" },
      HOLE_D: { value: 0.009, description: "hole diameter (m)" },
      HX: { value: 0.05, description: "seed hole X (m)" },
      HY: { value: 0.012, description: "seed hole Y (m)" },
      PITCH: { value: 0.012, description: "pattern pitch along X (m)" },
      COUNT: { value: 3, description: "number of holes" },
    },
    verify: { bodyCount: 1, featureTypes: ["Extrude", "Cut", "Pattern"] },
    script: "Option Explicit\n' Recipe: mounting plate with a linear pattern of holes.\n' Routes: extrude, blind cut from the +Z face, FeatureLinearPattern4.\n' Verified: plate {{PLATE_X}} x {{PLATE_Y}} x {{THK}}, hole d{{HOLE_D}},\n' {{COUNT}} holes at {{PITCH}} pitch -> exact volume (see FINDINGS.md 3.4).\nDim WSH, OUTDIR, swApp, sw, ext, sm, fm, f, fe, fc, cc, ok, pn\nConst PLATE_X = {{PLATE_X}}\nConst PLATE_Y = {{PLATE_Y}}\nConst THK = {{THK}}\nConst HOLE_D = {{HOLE_D}}\nConst HX = {{HX}}\nConst HY = {{HY}}\nConst PITCH = {{PITCH}}\nConst COUNT = {{COUNT}}\nSet WSH = CreateObject(\"WScript.Shell\")\nOUTDIR = WSH.ExpandEnvironmentStrings(\"%SWOUTDIR%\")\nIf Right(OUTDIR, 1) <> \"\\\" Then OUTDIR = OUTDIR & \"\\\"\n\nSub L(s) : WScript.Echo s : End Sub\nFunction IsNothing(o)\n    On Error Resume Next\n    Dim tn : tn = TypeName(o)\n    If Err.Number <> 0 Then IsNothing = True : Exit Function\n    IsNothing = (tn = \"Nothing\") Or (tn = \"Empty\")\nEnd Function\n\nOn Error Resume Next\nSet swApp = CreateObject(\"SldWorks.Application\")\nswApp.Visible = True\nDim g, a\ng = 0\nDo While swApp.GetDocumentCount > 0 And g < 30\n    Set a = swApp.ActiveDoc\n    If Not IsNothing(a) Then swApp.CloseDoc a.GetTitle Else swApp.CloseAllDocuments True\n    WScript.Sleep 300\n    g = g + 1\nLoop\nswApp.NewDocument WSH.ExpandEnvironmentStrings(\"%SWPARTTPL%\"), 0, 0, 0\nWScript.Sleep 2000\nSet sw = swApp.ActiveDoc\nSet ext = sw.Extension\nSet sm = sw.SketchManager\nSet fm = sw.FeatureManager\npn = \"\" : g = 0\nSet f = sw.FirstFeature\nDo While Not f Is Nothing\n    If f.GetTypeName2 = \"RefPlane\" Then\n        If g = 0 Then pn = f.Name\n        g = g + 1\n    End If\n    Set f = f.GetNextFeature\nLoop\n\n' plate on the front plane, extruded +Z\nsw.ClearSelection2 True\next.SelectByID2 pn, \"PLANE\", 0, 0, 0, False, 0, Nothing, 0\nsw.InsertSketch2 True\nCall sm.CreateLine(0.0, 0.0, 0.0, PLATE_X, 0.0, 0.0)\nCall sm.CreateLine(PLATE_X, 0.0, 0.0, PLATE_X, PLATE_Y, 0.0)\nCall sm.CreateLine(PLATE_X, PLATE_Y, 0.0, 0.0, PLATE_Y, 0.0)\nCall sm.CreateLine(0.0, PLATE_Y, 0.0, 0.0, 0.0, 0.0)\nsw.InsertSketch2 True\nSet f = sw.FeatureByPositionReverse(0)\nf.Name = \"PlateSketch\"\next.SelectByID2 \"PlateSketch\", \"SKETCH\", 0, 0, 0, False, 0, Nothing, 0\nSet fe = fm.FeatureExtrusion2(True, False, False, 0, 0, THK, 0.01, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)\nIf IsNothing(fe) Then L \"FATAL plate\" : WScript.Quit 0\nfe.Name = \"Plate\"\nL \"plate ok\"\n\n' seed hole from the top (+Z) face\nsw.ClearSelection2 True\nok = ext.SelectByRay(HX, HY, THK + 0.01, 0.0, 0.0, -1.0, 0.0001, 2, False, 0, 0)\nIf ok = 0 Then L \"FATAL hole face ray\" : WScript.Quit 0\nsw.InsertSketch2 True\nSet cc = sm.CreateCircleByRadius(HX, HY, 0.0, HOLE_D / 2)\nsw.InsertSketch2 True\nSet f = sw.FeatureByPositionReverse(0)\nf.Name = \"HoleSketch\"\next.SelectByID2 \"HoleSketch\", \"SKETCH\", 0, 0, 0, False, 0, Nothing, 0\nSet fc = fm.FeatureCut4(True, False, False, 1, 0, THK * 4, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, False, False, False, 0, 0, False, False)\nIf IsNothing(fc) Then L \"FATAL hole\" : WScript.Quit 0\nfc.Name = \"Hole\"\nL \"hole ok\"\n\n' linear pattern: direction edge (mark 1), seed feature (mark 4)\nsw.ClearSelection2 True\nok = ext.SelectByID2(\"\", \"EDGE\", PLATE_X, HY, THK / 2, False, 1, Nothing, 0)\nIf ok = 0 Then L \"FATAL direction edge\" : WScript.Quit 0\nok = ext.SelectByID2(\"Hole\", \"BODYFEATURE\", 0, 0, 0, True, 4, Nothing, 0)\nIf ok = 0 Then L \"FATAL seed feature\" : WScript.Quit 0\nSet fc = fm.FeatureLinearPattern4(COUNT, PITCH, 1, 0.0, False, False, \"\", \"\", False, False, False, False, False, False, False, False, False, False, 0.0, 0.0)\nIf IsNothing(fc) Then L \"FATAL linear pattern\" : WScript.Quit 0\nfc.Name = \"HolePattern\"\nL \"pattern ok\"\n\nsw.ClearSelection2 True\nsw.ForceRebuild3 False\nL \"save -> \" & sw.SaveAs3(OUTDIR & \"plate_linear_holes.SLDPRT\", 0, 1)\nL \"BUILD COMPLETE\"\n",
  },
  {
    name: "disc_multicontour_cut",
    description: "Disc with a central bore and N radial slots removed by one multi-contour sketch and one cut (no axis/pattern API needed).",
    parameters: {
      R_OUT: { value: 0.06, description: "disc radius (m)" },
      THK: { value: 0.012, description: "disc thickness (m)" },
      R_BORE: { value: 0.014, description: "bore radius (m)" },
      SLOT_R0: { value: 0.038, description: "slot inner radius (m)" },
      SLOT_R1: { value: 0.046, description: "slot outer radius (m)" },
      SLOT_HW: { value: 0.004, description: "slot half width (m)" },
      SLOTS: { value: 8, description: "number of radial slots" },
    },
    verify: { bodyCount: 1, featureTypes: ["Extrude", "Cut"] },
    script: "Option Explicit\n' Recipe: slotted disc, built as one multi-contour cut (no axis API, no pattern\n' API - InsertAxis2 terminates the script on this host, see FINDINGS.md 3.5).\n' One sketch holds the bore plus {{SLOTS}} rotated slot rectangles; a single\n' FeatureCut4 removes all of them. Verified exact on 8 slots / d{{R_BORE}}.\nDim WSH, OUTDIR, swApp, sw, ext, sm, fm, f, fe, fc, cc, ok, pn\nConst R_OUT = {{R_OUT}}\nConst THK = {{THK}}\nConst R_BORE = {{R_BORE}}\nConst SLOT_R0 = {{SLOT_R0}}\nConst SLOT_R1 = {{SLOT_R1}}\nConst SLOT_HW = {{SLOT_HW}}\nConst SLOTS = {{SLOTS}}\nConst PI = 3.14159265358979\nSet WSH = CreateObject(\"WScript.Shell\")\nOUTDIR = WSH.ExpandEnvironmentStrings(\"%SWOUTDIR%\")\nIf Right(OUTDIR, 1) <> \"\\\" Then OUTDIR = OUTDIR & \"\\\"\n\nSub L(s) : WScript.Echo s : End Sub\nFunction IsNothing(o)\n    On Error Resume Next\n    Dim tn : tn = TypeName(o)\n    If Err.Number <> 0 Then IsNothing = True : Exit Function\n    IsNothing = (tn = \"Nothing\") Or (tn = \"Empty\")\nEnd Function\n\nOn Error Resume Next\nSet swApp = CreateObject(\"SldWorks.Application\")\nswApp.Visible = True\nDim g, a, i, ang\ng = 0\nDo While swApp.GetDocumentCount > 0 And g < 30\n    Set a = swApp.ActiveDoc\n    If Not IsNothing(a) Then swApp.CloseDoc a.GetTitle Else swApp.CloseAllDocuments True\n    WScript.Sleep 300\n    g = g + 1\nLoop\nswApp.NewDocument WSH.ExpandEnvironmentStrings(\"%SWPARTTPL%\"), 0, 0, 0\nWScript.Sleep 2000\nSet sw = swApp.ActiveDoc\nSet ext = sw.Extension\nSet sm = sw.SketchManager\nSet fm = sw.FeatureManager\npn = \"\" : g = 0\nSet f = sw.FirstFeature\nDo While Not f Is Nothing\n    If f.GetTypeName2 = \"RefPlane\" Then\n        If g = 0 Then pn = f.Name\n        g = g + 1\n    End If\n    Set f = f.GetNextFeature\nLoop\n\n' disc\nsw.ClearSelection2 True\next.SelectByID2 pn, \"PLANE\", 0, 0, 0, False, 0, Nothing, 0\nsw.InsertSketch2 True\nSet cc = sm.CreateCircleByRadius(0.0, 0.0, 0.0, R_OUT)\nsw.InsertSketch2 True\nSet f = sw.FeatureByPositionReverse(0)\nf.Name = \"DiscSketch\"\next.SelectByID2 \"DiscSketch\", \"SKETCH\", 0, 0, 0, False, 0, Nothing, 0\nSet fe = fm.FeatureExtrusion2(True, False, False, 0, 0, THK, 0.01, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)\nIf IsNothing(fe) Then L \"FATAL disc\" : WScript.Quit 0\nfe.Name = \"Disc\"\nL \"disc ok\"\n\n' bore + slots in ONE sketch on the +Z face\nsw.ClearSelection2 True\nok = ext.SelectByRay(0.0, 0.0, THK + 0.01, 0.0, 0.0, -1.0, 0.0001, 2, False, 0, 0)\nIf ok = 0 Then L \"FATAL cut face ray\" : WScript.Quit 0\nsw.InsertSketch2 True\nSet cc = sm.CreateCircleByRadius(0.0, 0.0, 0.0, R_BORE)\nFor i = 0 To SLOTS - 1\n    ang = i * 2 * PI / SLOTS\n    Call sm.CreateLine(SLOT_R0 * Cos(ang) + SLOT_HW * Sin(ang), SLOT_R0 * Sin(ang) - SLOT_HW * Cos(ang), 0.0, SLOT_R1 * Cos(ang) + SLOT_HW * Sin(ang), SLOT_R1 * Sin(ang) - SLOT_HW * Cos(ang), 0.0)\n    Call sm.CreateLine(SLOT_R1 * Cos(ang) + SLOT_HW * Sin(ang), SLOT_R1 * Sin(ang) - SLOT_HW * Cos(ang), 0.0, SLOT_R1 * Cos(ang) - SLOT_HW * Sin(ang), SLOT_R1 * Sin(ang) + SLOT_HW * Cos(ang), 0.0)\n    Call sm.CreateLine(SLOT_R1 * Cos(ang) - SLOT_HW * Sin(ang), SLOT_R1 * Sin(ang) + SLOT_HW * Cos(ang), 0.0, SLOT_R0 * Cos(ang) - SLOT_HW * Sin(ang), SLOT_R0 * Sin(ang) + SLOT_HW * Cos(ang), 0.0)\n    Call sm.CreateLine(SLOT_R0 * Cos(ang) - SLOT_HW * Sin(ang), SLOT_R0 * Sin(ang) + SLOT_HW * Cos(ang), 0.0, SLOT_R0 * Cos(ang) + SLOT_HW * Sin(ang), SLOT_R0 * Sin(ang) - SLOT_HW * Cos(ang), 0.0)\nNext\nsw.InsertSketch2 True\nSet f = sw.FeatureByPositionReverse(0)\nf.Name = \"CutSketch\"\nIf IsNothing(f) Then L \"FATAL cut sketch\" : WScript.Quit 0\nsw.ClearSelection2 True\next.SelectByID2 \"CutSketch\", \"SKETCH\", 0, 0, 0, False, 4, Nothing, 0\nSet fc = fm.FeatureCut4(True, False, False, 1, 0, THK * 4, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, False, False, False, 0, 0, False, False)\nIf IsNothing(fc) Then L \"FATAL multi-contour cut\" : WScript.Quit 0\nfc.Name = \"BoreAndSlots\"\nL \"cut ok\"\n\nsw.ClearSelection2 True\nsw.ForceRebuild3 False\nL \"save -> \" & sw.SaveAs3(OUTDIR & \"slotted_disc.SLDPRT\", 0, 1)\nL \"BUILD COMPLETE\"\n",
  },
]

/** Parameter values must be finite numbers: they are spliced into VBScript. */
function isFiniteNumber(value) {
  return typeof value === 'number' && Number.isFinite(value)
}

/**
 * Validate one recipe object (from disk or from a tool call).
 *
 * @param raw - untrusted recipe candidate.
 * @param source - where it came from, for diagnostics.
 * @returns `{ ok: true, recipe }` with defaults applied, or `{ ok: false, reason }`.
 */
export function validateRecipe(raw, source = 'recipe') {
  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) {
    return { ok: false, reason: `${source}: must be a JSON object` }
  }
  if (typeof raw.name !== 'string' || !RECIPE_NAME.test(raw.name)) {
    return { ok: false, reason: `${source}: "name" must match ${RECIPE_NAME} (lowercase letters, digits, _ or -)` }
  }
  if (typeof raw.script !== 'string' || raw.script.trim().length === 0) {
    return { ok: false, reason: `${source}: "script" must be a non-empty string` }
  }
  if (Buffer.byteLength(raw.script, 'utf8') > MAX_SCRIPT_BYTES) {
    return { ok: false, reason: `${source}: "script" exceeds ${MAX_SCRIPT_BYTES} bytes` }
  }
  const nonAscii = [...raw.script].find((ch) => ch.charCodeAt(0) > 0x7e)
  if (nonAscii !== undefined) {
    return { ok: false, reason: `${source}: "script" contains non-ASCII text (${JSON.stringify(nonAscii)}); WSH reads .vbs as ANSI, so keep script and comments ASCII` }
  }

  // `%NAME%` is the environment-variable convention the scripts already use
  // (%SWPARTTPL%, %SWOUTDIR%), so parameter placeholders are `{{NAME}}`.
  const envStyle = [...raw.script.matchAll(/%([A-Za-z_][A-Za-z0-9_]*)%/g)]
    .map((m) => m[1])
    .filter((token) => !['SWPARTTPL', 'SWOUTDIR', 'TEMP', 'TMP', 'USERPROFILE'].includes(token))
  if (envStyle.length > 0) {
    return { ok: false, reason: `${source}: script uses %${envStyle[0]}%, which is not a parameter; parameter placeholders are written {{${envStyle[0]}}}` }
  }

  const parameters = raw.parameters ?? {}
  if (typeof parameters !== 'object' || parameters === null || Array.isArray(parameters)) {
    return { ok: false, reason: `${source}: "parameters" must be an object` }
  }
  const normalisedParameters = {}
  for (const [key, spec] of Object.entries(parameters)) {
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) {
      return { ok: false, reason: `${source}: parameter "${key}" is not a valid identifier` }
    }
    const entry = typeof spec === 'object' && spec !== null ? spec : { value: spec }
    if (!isFiniteNumber(entry.value)) {
      return { ok: false, reason: `${source}: parameter "${key}" needs a finite numeric "value"` }
    }
    if (entry.description !== undefined && typeof entry.description !== 'string') {
      return { ok: false, reason: `${source}: parameter "${key}" description must be a string` }
    }
    normalisedParameters[key] = { value: entry.value, description: entry.description ?? '' }
  }

  // Every placeholder must resolve, and every declared parameter must be used,
  // so a typo cannot silently produce an unparameterised script.
  const used = new Set()
  for (const match of raw.script.matchAll(PLACEHOLDER)) used.add(match[1])
  for (const key of used) {
    if (!(key in normalisedParameters)) {
      return { ok: false, reason: `${source}: script uses {{${key}}} but declares no such parameter` }
    }
  }
  for (const key of Object.keys(normalisedParameters)) {
    if (!used.has(key)) {
      return { ok: false, reason: `${source}: parameter "${key}" never appears as {{${key}}} in the script` }
    }
  }

  const verify = raw.verify ?? {}
  if (typeof verify !== 'object' || verify === null || Array.isArray(verify)) {
    return { ok: false, reason: `${source}: "verify" must be an object` }
  }
  if (verify.bodyCount !== undefined && (!Number.isInteger(verify.bodyCount) || verify.bodyCount < 0)) {
    return { ok: false, reason: `${source}: verify.bodyCount must be a non-negative integer` }
  }
  if (verify.featureTypes !== undefined) {
    if (!Array.isArray(verify.featureTypes) || verify.featureTypes.some((t) => typeof t !== 'string')) {
      return { ok: false, reason: `${source}: verify.featureTypes must be an array of strings` }
    }
  }
  // An assembly recipe asserts components and mates instead of a body count,
  // because a body count is not meaningful for an assembly and reporting one
  // was what made an assembly-shaped recipe impossible to record.
  for (const key of ['componentCount', 'mateCount']) {
    if (verify[key] !== undefined && (!Number.isInteger(verify[key]) || verify[key] < 0)) {
      return { ok: false, reason: `${source}: verify.${key} must be a non-negative integer` }
    }
  }

  if (raw.description !== undefined && typeof raw.description !== 'string') {
    return { ok: false, reason: `${source}: "description" must be a string` }
  }

  return {
    ok: true,
    recipe: {
      name: raw.name,
      description: raw.description ?? '',
      parameters: normalisedParameters,
      verify: {
        ...(verify.bodyCount === undefined ? {} : { bodyCount: verify.bodyCount }),
        ...(verify.featureTypes === undefined ? {} : { featureTypes: verify.featureTypes }),
        ...(verify.componentCount === undefined ? {} : { componentCount: verify.componentCount }),
        ...(verify.mateCount === undefined ? {} : { mateCount: verify.mateCount }),
      },
      script: raw.script,
      shipped: raw.shipped === true,
      source,
    },
  }
}

/** Read shipped recipes, skipping any that fail validation. */
export function loadShipped(log = () => {}) {
  const out = []
  for (const raw of SHIPPED_RECIPES) {
    const result = validateRecipe({ ...raw, shipped: true }, `shipped recipe ${raw.name}`)
    if (result.ok) out.push(result.recipe)
    else log(`skipping ${result.reason}`)
  }
  return out
}

/**
 * Read user recipes from a directory. Each file is one recipe; an unreadable or
 * invalid file is skipped with a warning and never affects the plugin.
 */
export function loadUserRecipes(dir, log = () => {}) {
  if (!dir || !existsSync(dir)) return []
  const out = []
  let entries = []
  try {
    entries = readdirSync(dir).filter((name) => name.endsWith('.json'))
  } catch (error) {
    log(`cannot read recipe dir ${dir}: ${error instanceof Error ? error.message : String(error)}`)
    return []
  }
  for (const entry of entries) {
    const path = join(dir, entry)
    try {
      const parsed = JSON.parse(readFileSync(path, 'utf8'))
      const result = validateRecipe(parsed, `recipe file ${entry}`)
      if (result.ok) out.push(result.recipe)
      else log(`skipping ${result.reason}`)
    } catch (error) {
      log(`skipping recipe file ${entry}: ${error instanceof Error ? error.message : String(error)}`)
    }
  }
  return out
}

/**
 * Merge shipped then user recipes; a user recipe with a shipped name wins,
 * which is how a shipped default gets tuned without a code change.
 */
export function mergeRecipes(shipped, user, log = () => {}) {
  const byName = new Map()
  for (const recipe of shipped) byName.set(recipe.name, recipe)
  for (const recipe of user) {
    if (byName.has(recipe.name) && byName.get(recipe.name).shipped) {
      log(`recipe "${recipe.name}" overrides the shipped version`)
    }
    byName.set(recipe.name, recipe)
  }
  return [...byName.values()].sort((a, b) => a.name.localeCompare(b.name))
}

/** Persist one recipe as `<name>.json`; returns the path written. */
export function saveRecipeFile(dir, recipe) {
  mkdirSync(dir, { recursive: true })
  const path = join(dir, `${recipe.name}.json`)
  const payload = {
    name: recipe.name,
    description: recipe.description,
    parameters: recipe.parameters,
    verify: recipe.verify,
    script: recipe.script,
  }
  writeFileSync(path, `${JSON.stringify(payload, null, 2)}\n`, 'utf8')
  return path
}

/** Remove a saved recipe; returns true when a file was deleted. */
export function removeRecipeFile(dir, name) {
  const path = join(dir, `${name}.json`)
  if (!existsSync(path)) return false
  rmSync(path, { force: true })
  return true
}

/**
 * Splice parameter values into a script. Values are validated as finite
 * numbers before this runs, so the resulting source is arithmetic-safe.
 *
 * @param recipe - validated recipe.
 * @param overrides - optional `{ PARAM: value }` overrides of the defaults.
 * @returns `{ ok: true, script, values }` or `{ ok: false, reason }`.
 */
export function materialize(recipe, overrides = {}) {
  const values = {}
  for (const [key, spec] of Object.entries(recipe.parameters)) values[key] = spec.value
  for (const [key, value] of Object.entries(overrides ?? {})) {
    if (!(key in values)) {
      return { ok: false, reason: `unknown parameter "${key}" (recipe declares: ${Object.keys(values).join(', ') || 'none'})` }
    }
    if (!isFiniteNumber(value)) {
      return { ok: false, reason: `parameter "${key}" must be a finite number` }
    }
    values[key] = value
  }
  let script = recipe.script
  for (const [key, value] of Object.entries(values)) {
    script = script.split(`{{${key}}}`).join(String(value))
  }
  // Only `{{NAME}}` marks a parameter. `%SWPARTTPL%` and friends are Windows
  // environment variables the script expands itself at run time, so they must
  // survive substitution untouched.
  const leftover = script.match(/\{\{\s*[A-Za-z_][A-Za-z0-9_]*\s*\}\}/)
  if (leftover !== null) {
    return { ok: false, reason: `script still contains an unresolved ${leftover[0]} after substitution` }
  }
  return { ok: true, script, values }
}

/** One-line-per-recipe catalogue for tool descriptions and listings. */
export function describeRecipes(recipes, includeParameters = true) {
  if (recipes.length === 0) return '(none)'
  return recipes.map((recipe) => {
    const params = includeParameters
      ? Object.entries(recipe.parameters).map(([k, v]) => `${k}=${v.value}`).join(' ')
      : ''
    const verify = [
      recipe.verify.bodyCount === undefined ? '' : `bodyCount=${recipe.verify.bodyCount}`,
      recipe.verify.featureTypes?.length ? `kinds=[${recipe.verify.featureTypes.join(',')}]` : '',
      recipe.verify.componentCount === undefined ? '' : `componentCount=${recipe.verify.componentCount}`,
      recipe.verify.mateCount === undefined ? '' : `mateCount=${recipe.verify.mateCount}`,
    ].filter(Boolean).join(' ')
    return [
      `- ${recipe.name}${recipe.shipped ? ' (shipped)' : ''}: ${recipe.description || 'no description'}`,
      params ? `    parameters: ${params}` : '',
      verify ? `    verify: ${verify}` : '',
    ].filter(Boolean).join('\n')
  }).join('\n')
}
