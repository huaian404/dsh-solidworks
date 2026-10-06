Option Explicit

' =====================================================================
'  SolidWorks capability / workspace probe
'
'  NON-DESTRUCTIVE: it may open one throwaway scratch part to exercise the
'  part-level API, and it closes everything it opened. It never enters a
'  sketch, never creates a feature on your work, and never saves.
'
'  Usage:  cscript //nologo probe.vbs <outJson>
'  Env:    SWPARTTPL  part template (.prtdot)
'
'  "ok" semantics: true = interface exposes the member and it returned;
'  false = missing member or a VBScript runtime error (the "broken on this
'  host" signal used to route modelling strategies); null = not determined.
'
'  capabilities.sweptCut drives the thread strategy: when false, threads must
'  be modelled as a core body plus a merged swept BOSS, not a swept cut.
' =====================================================================

Dim fso, outJson, t0
Set fso = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count < 1 Then
    WScript.Echo "usage: cscript //nologo probe.vbs <outJson>"
    WScript.Quit 2
End If
outJson = WScript.Arguments(0)
t0 = Timer()

Sub E(s)
    WScript.Echo "[" & FormatNumber(Timer() - t0, 2) & "s] " & s
End Sub

Function JStr(s)
    Dim r
    r = CStr(s)
    r = Replace(r, "\", "\\")
    r = Replace(r, """", "\""")
    r = Replace(r, vbCrLf, " ")
    r = Replace(r, vbCr, " ")
    r = Replace(r, vbLf, " ")
    r = Replace(r, vbTab, " ")
    JStr = """" & r & """"
End Function

Function JBool(b)
    If b Then JBool = "true" Else JBool = "false"
End Function

Dim apiLines, apiFirst
apiLines = ""
apiFirst = True

Sub Api(name, val, detail)
    Dim ok
    If IsObject(val) Then
        ok = Not (TypeName(val) = "Nothing")
    ElseIf VarType(val) = vbEmpty Then
        ok = False
    Else
        ok = True
    End If
    If Not apiFirst Then apiLines = apiLines & "," & vbCrLf
    apiFirst = False
    apiLines = apiLines & "    " & JStr(name) & ": { ""ok"": " & JBool(ok) & _
               ", ""detail"": " & JStr(detail) & " }"
    E "api " & name & " ok=" & ok & " (" & detail & ")"
End Sub

Sub WriteOut(path, text)
    Dim f
    On Error Resume Next
    Err.Clear
    Set f = fso.CreateTextFile(path, True, False)
    If Err.Number = 0 Then
        f.Write text
        f.Close
        E "wrote " & path
    Else
        E "cannot write " & path & ": " & Err.Description
    End If
End Sub

' =====================================================================
'  Part-level capability probes
'
'  Three routes fail SILENTLY on this host, and each one is worth probing
'  because the symptom is indistinguishable from "my script is wrong":
'
'    filletOptions   FeatureFillet3 returns Nothing unless Options includes
'                    the uniform-radius bit; Options 0/1/4 fail with Err=0.
'    negZFaceCut     a FeatureCut4 whose sketch sits on a face whose outward
'                    normal is -Z returns Nothing for every Flip/Dir/T1
'                    combination, while the same cut on the +Z face works.
'    shell / axis    InsertFeatureShell and InsertAxis2 do not produce the
'                    feature here. They are run in a CHILD process so that a
'                    procedure-level error aborting the probe cannot lose the
'                    measurements taken above.
'
'  VBScript error-handling fact this file depends on (measured - see
'  parts/FINDINGS.md 3.5): `On Error Resume Next` at file level does NOT apply
'  inside a Sub or Function. An error raised inside a procedure terminates THAT
'  PROCEDURE; statements after the call in the caller keep running, and the
'  script itself survives. A build written as one Sub per part therefore loses
'  every feature after the first failing call, silently. Every procedure below
'  starts with its own On Error Resume Next for that reason.
' =====================================================================
Const PROBE_X = 0.065
Const PROBE_Y = 0.050
Const PROBE_T = 0.012
Const PROBE_R = 1
Const PROBE_FIL = 0.003
Const PROBE_HD = 0.009

Dim probeOK, probeDetail, abortProbed, abortShellOK, abortAxisOK, abortNote
Dim abortShellDetail, abortAxisDetail, abortProbeRun, probeDims, probeVolBase
Dim filletOptions(3), filletOK(3), filletDetail(3), negZOK, negZDetail
Dim probeFilletOk, probeFilletOption, probesRun
Dim fProbe, ix

Function PVol()
    Dim mp2, v2
    PVol = -1
    On Error Resume Next
    Err.Clear
    Set mp2 = ext.CreateMassProperty2
    If IsEmpty(mp2) Or TypeName(mp2) = "Nothing" Then Exit Function
    Err.Clear
    v2 = mp2.Volume
    If Err.Number <> 0 Or Not IsNumeric(v2) Then Exit Function
    PVol = v2 * 1000000000
End Function

Sub FreshPart()
    ' procedure-local: a file-level On Error Resume Next does NOT apply in here
    On Error Resume Next
    Set sw = swApp.ActiveDoc
    Set ext = sw.Extension
    Set sm = sw.SketchManager
    Set fm = sw.FeatureManager
    Dim pl, ffp, n : n = 0
    pl = ""
    Set ffp = sw.FirstFeature
    Do While Not ffp Is Nothing
        If ffp.GetTypeName2 = "RefPlane" Then
            If n = 0 Then pl = ffp.Name
            n = n + 1
        End If
        Set ffp = ffp.GetNextFeature
    Loop
    probeDims = (pl <> "")
    If Not probeDims Then
        probeDetail = "no reference plane in the scratch part"
        Exit Sub
    End If
    sw.ClearSelection2 True
    ext.SelectByID2 pl, "PLANE", 0, 0, 0, False, 0, Nothing, 0
    sw.InsertSketch2 True
    Call sm.CreateLine(0.0, 0.0, 0.0, PROBE_X, 0.0, 0.0)
    Call sm.CreateLine(PROBE_X, 0.0, 0.0, PROBE_X, PROBE_Y, 0.0)
    Call sm.CreateLine(PROBE_X, PROBE_Y, 0.0, 0.0, PROBE_Y, 0.0)
    Call sm.CreateLine(0.0, PROBE_Y, 0.0, 0.0, 0.0, 0.0)
    sw.InsertSketch2 True
    Set fProbe = sw.FeatureByPositionReverse(0)
    fProbe.Name = "ProbeSketch"
    ext.SelectByID2 "ProbeSketch", "SKETCH", 0, 0, 0, False, 0, Nothing, 0
    Set fProbe = fm.FeatureExtrusion2(True, False, False, 0, 0, PROBE_T, 0.01, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
    If IsEmpty(fProbe) Or TypeName(fProbe) = "Nothing" Then
        probeDims = False
        probeDetail = "scratch block extrude returned Nothing"
        Exit Sub
    End If
    fProbe.Name = "ProbeBlock"
    probeVolBase = PVol()
    probeDetail = "block " & (PROBE_X * 1000) & "x" & (PROBE_Y * 1000) & "x" & (PROBE_T * 1000) & " mm, V=" & FormatNumber(probeVolBase, 1) & " mm^3"
End Sub

' Fillet Options matrix on one scratch part.
Sub ProbeFilletOptions()
    On Error Resume Next
    Dim m, s, vNow
    probeFilletOk = False
    probeFilletOption = -1
    filletOptions(0) = 0 : filletOptions(1) = 1 : filletOptions(2) = 2 : filletOptions(3) = 3
    For ix = 0 To 3
        filletOK(ix) = False
        filletDetail(ix) = "not attempted"
    Next
    For ix = 0 To 3
        m = 0
        sw.ClearSelection2 True
        If ext.SelectByID2("", "EDGE", PROBE_X, 0.0, PROBE_T / 2, False, 0, Nothing, 0) Then m = m + 1
        If ext.SelectByID2("", "EDGE", 0.0, PROBE_Y, PROBE_T / 2, True, 0, Nothing, 0) Then m = m + 1
        If m = 0 Then
            filletDetail(ix) = "no edge selected"
        Else
            Err.Clear
            Set fProbe = fm.FeatureFillet3(filletOptions(ix), PROBE_FIL, 0.0, 0.0, 0, 0, 0, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
            If IsEmpty(fProbe) Or TypeName(fProbe) = "Nothing" Then
                filletDetail(ix) = "Options=" & filletOptions(ix) & " -> Nothing (err=" & Err.Number & " " & Err.Description & "), edges=" & m
            Else
                fProbe.Name = "ProbeFillet" & filletOptions(ix)
                vNow = PVol()
                filletOK(ix) = True
                filletDetail(ix) = "Options=" & filletOptions(ix) & " -> " & fProbe.GetTypeName2 & ", dV=" & FormatNumber(vNow - probeVolBase, 1) & " mm^3, edges=" & m
                If Not probeFilletOk Then
                    probeFilletOk = True
                    probeFilletOption = filletOptions(ix)
                End If
            End If
        End If
        E "  fillet options " & filletOptions(ix) & ": " & filletDetail(ix)
    Next
End Sub

' Cut on the -Z facing face of the same scratch block.
Sub ProbeNegZCut()
    On Error Resume Next
    Dim okc, m
    negZOK = False
    negZDetail = "not attempted"
    sw.ClearSelection2 True
    okc = ext.SelectByRay(PROBE_X / 2, PROBE_Y / 2, -0.010, 0.0, 0.0, 1.0, 0.0001, 2, False, 0, 0)
    If okc = 0 Then
        negZDetail = "could not select the -Z face (SelectByRay=0)"
        E "  negZ cut: " & negZDetail
        Exit Sub
    End If
    sw.InsertSketch2 True
    Set fProbe = sw.FeatureByPositionReverse(0)
    If IsEmpty(fProbe) Or TypeName(fProbe) = "Nothing" Then
        negZDetail = "no sketch was created on the -Z face"
        E "  negZ cut: " & negZDetail
        Exit Sub
    End If
    fProbe.Name = "NegZSketch"
    Call sm.CreateCircleByRadius(PROBE_X / 2, PROBE_Y / 2, 0.0, PROBE_HD / 2)
    sw.InsertSketch2 True
    sw.ClearSelection2 True
    ext.SelectByID2 "NegZSketch", "SKETCH", 0, 0, 0, False, 0, Nothing, 0
    Err.Clear
    Set fProbe = fm.FeatureCut4(True, False, False, 1, 0, PROBE_T * 4, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, False, False, False, 0, 0, False, False)
    If IsEmpty(fProbe) Or TypeName(fProbe) = "Nothing" Then
        negZDetail = "FeatureCut4 on the -Z face -> Nothing (err=" & Err.Number & " " & Err.Description & ")"
    Else
        negZOK = True
        negZDetail = "FeatureCut4 on the -Z face -> " & fProbe.GetTypeName2 & ", dV=" & FormatNumber(PVol() - probeVolBase, 1) & " mm^3"
    End If
    E "  negZ cut: " & negZDetail
End Sub

' InsertFeatureShell / InsertAxis2: run in a CHILD probe process.
'
' Why a child: an error raised inside this Sub would abort this Sub (VBScript
' does not inherit On Error Resume Next into a procedure), so the probe could
' not report what the call did. Running it in a child keeps that failure away
' from the measurements already taken, and the child's markers say exactly
' where it stopped.
Sub ProbeScriptTerminatingCalls()
    On Error Resume Next
    Dim tmp, childPath, lineArr, res, marker, childCmd, mode, keepgoing
    abortProbed = True
    abortShellOK = False
    abortAxisOK = False
    abortShellDetail = "not attempted"
    abortAxisDetail = "not attempted"
    abortNote = ""

    tmp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%SWOUTDIR%")
    If Right(tmp, 1) <> "\" Then tmp = tmp & "\"
    childPath = tmp & "probe_abort_child.vbs"

    ' Fresh scratch part for the child; if it cannot be created, skip.
    Err.Clear
    swApp.NewDocument tpl, 0, 0, 0
    WScript.Sleep 1800
    If Err.Number <> 0 Or IsEmpty(swApp.ActiveDoc) Or TypeName(swApp.ActiveDoc) = "Nothing" Then
        abortShellDetail = "skipped: no scratch part for the child probe"
        abortAxisDetail = abortShellDetail
        E "  abort probes skipped (scratch part unavailable)"
        Exit Sub
    End If

    abortProbeRun = 0
    Do While abortProbeRun < 2
        marker = tmp & "probe_abort_marker" & abortProbeRun & ".txt"
        On Error Resume Next
        If fso.FileExists(marker) Then fso.DeleteFile marker, True
        If abortProbeRun = 0 Then mode = "shell" Else mode = "axis"
        Set res = fso.CreateTextFile(childPath, True, False)
        res.WriteLine "Option Explicit"
        res.WriteLine "Dim swApp, sw, ext, sm, fm, m, errs, f, pl, ffp, n"
        res.WriteLine "Set swApp = CreateObject(""SldWorks.Application"")"
        res.WriteLine "On Error Resume Next"
        res.WriteLine "Set sw = swApp.ActiveDoc"
        res.WriteLine "Set ext = sw.Extension"
        res.WriteLine "Set sm = sw.SketchManager"
        res.WriteLine "Set fm = sw.FeatureManager"
        res.WriteLine "pl = """" : n = 0"
        res.WriteLine "Set ffp = sw.FirstFeature"
        res.WriteLine "Do While Not ffp Is Nothing"
        res.WriteLine "    If ffp.GetTypeName2 = ""RefPlane"" Then"
        res.WriteLine "        If n = 0 Then pl = ffp.Name"
        res.WriteLine "        n = n + 1"
        res.WriteLine "    End If"
        res.WriteLine "    Set ffp = ffp.GetNextFeature"
        res.WriteLine "Loop"
        res.WriteLine "sw.ClearSelection2 True"
        res.WriteLine "ext.SelectByID2 pl, ""PLANE"", 0, 0, 0, False, 0, Nothing, 0"
        res.WriteLine "sw.InsertSketch2 True"
        res.WriteLine "Call sm.CreateLine(0.0, 0.0, 0.0, 0.05, 0.0, 0.0)"
        res.WriteLine "Call sm.CreateLine(0.05, 0.0, 0.0, 0.05, 0.04, 0.0)"
        res.WriteLine "Call sm.CreateLine(0.05, 0.04, 0.0, 0.0, 0.04, 0.0)"
        res.WriteLine "Call sm.CreateLine(0.0, 0.04, 0.0, 0.0, 0.0, 0.0)"
        res.WriteLine "sw.InsertSketch2 True"
        res.WriteLine "Set f = sw.FeatureByPositionReverse(0)"
        res.WriteLine "f.Name = ""AbortProbeSketch"""
        res.WriteLine "ext.SelectByID2 ""AbortProbeSketch"", ""SKETCH"", 0, 0, 0, False, 0, Nothing, 0"
        res.WriteLine "Set f = fm.FeatureExtrusion2(True, False, False, 0, 0, 0.012, 0.01, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)"
        res.WriteLine "WScript.Echo ""ABORT_PROBE_START " & mode & """"
        If abortProbeRun = 0 Then
            res.WriteLine "sw.ClearSelection2 True"
            res.WriteLine "ext.SelectByRay 0.025, 0.02, 0.012 + 0.01, 0.0, 0.0, -1.0, 0.0001, 2, False, 0, 0"
            res.WriteLine "Err.Clear"
            res.WriteLine "Set f = sw.InsertFeatureShell(0.004, False)"
            res.WriteLine "WScript.Echo ""SHELL type="" & TypeName(f) & "" err="" & Err.Number & "" "" & Err.Description"
            ' a returned object is NOT success: this call can return an object AND
            ' set Err (measured: Object + err 424 "object missing"). Feature only
            ' when Err is clean.
            res.WriteLine "If Err.Number = 0 Then"
            res.WriteLine "    If TypeName(f) = ""Nothing"" Or TypeName(f) = ""Empty"" Then"
            res.WriteLine "        WScript.Echo ""SHELL FEATURE no"""
            res.WriteLine "    Else"
            res.WriteLine "        WScript.Echo ""SHELL FEATURE yes "" & f.GetTypeName2"
            res.WriteLine "    End If"
            res.WriteLine "Else"
            res.WriteLine "    WScript.Echo ""SHELL FEATURE no (err set)"""
            res.WriteLine "End If"
        Else
            res.WriteLine "sw.ClearSelection2 True"
            res.WriteLine "ext.SelectByID2 pl, ""PLANE"", 0, 0, 0, False, 1, Nothing, 0"
            res.WriteLine "ext.SelectByID2 pl, ""PLANE"", 0, 0, 0, True, 1, Nothing, 0"
            res.WriteLine "Err.Clear"
            res.WriteLine "Set f = fm.InsertAxis2(True)"
            res.WriteLine "WScript.Echo ""AXIS type="" & TypeName(f) & "" err="" & Err.Number & "" "" & Err.Description"
            res.WriteLine "If Err.Number = 0 Then"
            res.WriteLine "    If TypeName(f) = ""Nothing"" Or TypeName(f) = ""Empty"" Then"
            res.WriteLine "        WScript.Echo ""AXIS FEATURE no"""
            res.WriteLine "    Else"
            res.WriteLine "        WScript.Echo ""AXIS FEATURE yes "" & f.GetTypeName2"
            res.WriteLine "    End If"
            res.WriteLine "Else"
            res.WriteLine "    WScript.Echo ""AXIS FEATURE no (err set)"""
            res.WriteLine "End If"
        End If
        res.WriteLine "WScript.Echo ""ABORT_PROBE_END " & mode & """"
        res.Close

        childCmd = "cmd /c cscript.exe //nologo """ & childPath & """ > """ & marker & """ 2>&1"
        Err.Clear
        WScript.Sleep 300
        CreateObject("WScript.Shell").Run childCmd, 0, True
        WScript.Sleep 300

        On Error Resume Next
        lineArr = ""
        keepgoing = fso.FileExists(marker)
        If keepgoing Then
            Set res = fso.OpenTextFile(marker, 1)
            Do While Not res.AtEndOfStream
                lineArr = lineArr & res.ReadLine & " | "
            Loop
            res.Close
        End If
        If abortProbeRun = 0 Then
            abortShellOK = (InStr(lineArr, "SHELL FEATURE yes") > 0)
            abortShellDetail = lineArr
        Else
            abortAxisOK = (InStr(lineArr, "AXIS FEATURE yes") > 0)
            abortAxisDetail = lineArr
        End If
        E "  abort probe " & mode & ": " & lineArr
        abortProbeRun = abortProbeRun + 1
    Loop
End Sub

' Close the scratch document this probe created, by title: never CloseAllDocuments.
Sub CloseScratchDoc(expectTitle)
    On Error Resume Next
    Dim nmNow
    On Error Resume Next
    Err.Clear
    Set sw = swApp.ActiveDoc
    If IsEmpty(sw) Or TypeName(sw) = "Nothing" Then Exit Sub
    nmNow = sw.GetTitle
    If Len(expectTitle) > 0 And LCase(nmNow) <> LCase(expectTitle) Then
        E "scratch part '" & expectTitle & "' is not active (active: " & nmNow & "); leaving it open"
        Exit Sub
    End If
    Err.Clear
    swApp.CloseDoc nmNow
    If Err.Number = 0 Then
        E "scratch part closed (" & nmNow & ")"
    Else
        E "could not close scratch part: " & Err.Description
    End If
End Sub


' =====================================================================
Dim swApp, sw, ext, sm, fm, tpl, docCount, revision, hadWork
Dim f0, tn, bds, mp, d18, d15, bodyCount, scratchPart
Dim sweptCutOk, sweptCutProbed, revolveDefOk, scratchName
Dim errMsg
On Error Resume Next

sweptCutOk = False
sweptCutProbed = False
revolveDefOk = False
bodyCount = -1
errMsg = ""
scratchName = ""

Set swApp = CreateObject("SldWorks.Application")
If Err.Number <> 0 Then
    WriteOut outJson, "{""version"": 1, ""error"": " & JStr("cannot create SldWorks.Application: " & Err.Description) & "}"
    WScript.Quit 1
End If
swApp.Visible = True

Err.Clear
revision = swApp.RevisionNumber
docCount = swApp.GetDocumentCount
E "SolidWorks " & revision & " ; docs open = " & docCount

' --- app level --------------------------------------------------------
Err.Clear
Dim up
up = swApp.GetUserPreferenceStringValue(8)
Api "app_user_preference_string", up, "GetUserPreferenceStringValue(8) -> " & TypeName(up)

' --- template ---------------------------------------------------------
tpl = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%SWPARTTPL%")
If Len(tpl) > 0 Then
    Api "template_exists", fso.FileExists(tpl), tpl
Else
    Api "template_exists", False, "SWPARTTPL not set (launch through run-solidworks-vbs.ps1)"
End If

' --- part level, on a throwaway scratch part --------------------------
hadWork = (docCount > 0)
If hadWork Then E "user documents are open; probing on a scratch part instead of touching them"

If Len(tpl) > 0 Then
    Err.Clear
    swApp.NewDocument tpl, 0, 0, 0
    WScript.Sleep 2500
    Set sw = swApp.ActiveDoc
    If Not IsEmpty(sw) And TypeName(sw) <> "Nothing" Then
        scratchPart = True
        scratchName = sw.GetTitle
        E "scratch part opened: " & scratchName

        Set ext = sw.Extension
        Api "doc_extension", ext, "ModelDoc2.Extension -> " & TypeName(ext)

        Err.Clear
        Set f0 = sw.FirstFeature
        Api "doc_first_feature", f0, "FirstFeature -> " & TypeName(f0) & " err=" & Err.Number
        If Not IsEmpty(f0) And TypeName(f0) <> "Nothing" Then
            Err.Clear
            tn = f0.GetTypeName2
            Api "feature_get_type_name2", tn, "-> " & tn & " err=" & Err.Number
        End If

        Err.Clear
        bds = sw.GetBodies2(0, True)
        Api "doc_get_bodies2", bds, "GetBodies2 -> " & TypeName(bds) & " err=" & Err.Number

        Err.Clear
        mp = sw.GetMassProperties(1)
        Api "doc_get_mass_properties", mp, "GetMassProperties -> " & TypeName(mp) & " err=" & Err.Number

        Err.Clear
        Set sm = sw.SketchManager
        Api "doc_sketch_manager", sm, "SketchManager -> " & TypeName(sm)

        Err.Clear
        Set fm = sw.FeatureManager
        Api "doc_feature_manager", fm, "FeatureManager -> " & TypeName(fm)

        If Not IsEmpty(fm) And TypeName(fm) <> "Nothing" Then
            ' the swept-cut route: CreateDefinition(18) = swFmSweepCut
            Err.Clear
            Set d18 = fm.CreateDefinition(CLng(18))
            sweptCutProbed = True
            sweptCutOk = (Not IsEmpty(d18)) And (TypeName(d18) <> "Nothing")
            Api "fm_create_definition_18_sweepcut", d18, "CreateDefinition(18) -> " & TypeName(d18) & " err=" & Err.Number & " " & Err.Description

            ' the definition-based revolve route (CreateDefinition(15))
            Err.Clear
            Set d15 = fm.CreateDefinition(CLng(15))
            revolveDefOk = (Not IsEmpty(d15)) And (TypeName(d15) <> "Nothing")
            Api "fm_create_definition_15_revolve", d15, "CreateDefinition(15) -> " & TypeName(d15) & " err=" & Err.Number & " " & Err.Description
        End If

        ' --- the three silent-failure routes ------------------------------
        probeDims = True
        probeVolBase = -1
        Call FreshPart()
        probesRun = probeDims
        If probeDims Then
            E "capability probes on a " & probeDetail
            Err.Clear
            Call ProbeFilletOptions()
            Err.Clear
            Call ProbeNegZCut()
        Else
            E "capability probes skipped: " & probeDetail
            probeFilletOk = False
            probeFilletOption = -1
            negZOK = False
            negZDetail = "skipped: " & probeDetail
            For ix = 0 To 3
                filletOK(ix) = False
                filletDetail(ix) = "skipped: " & probeDetail
            Next
        End If
    Else
        Api "scratch_part", "", "NewDocument produced no ActiveDoc"
    End If
Else
    Api "scratch_part", "", "skipped: no template"
End If

' --- child-process probes for the calls that ABORT the script -----------
' These run last, in a child cscript: if InsertFeatureShell or InsertAxis2
' kills that child, the parent's measurements are already written and a
' missing completion marker is itself the result.
If Len(tpl) > 0 Then
    Err.Clear
    Call ProbeScriptTerminatingCalls()
End If

' --- clean up only the scratch parts this probe opened ------------------
' CloseAllDocuments would also close whatever the user has open, so close
' exactly the documents this probe created, by title, and only while each is
' still the active one.
If scratchPart Then
    Call CloseScratchDoc(scratchName)
    WScript.Sleep 600
End If
If abortProbeRun > 0 Then
    ' the child probes created one more scratch part each; still-active ones
    ' are the leftovers
    Dim leftover, tries
    tries = 0
    Do While tries < 3
        Err.Clear
        Set leftover = swApp.ActiveDoc
        If IsEmpty(leftover) Or TypeName(leftover) = "Nothing" Then Exit Do
        If InStr(leftover.GetTitle, "Part") = 0 And InStr(leftover.GetTitle, "part") = 0 Then Exit Do
        Call CloseScratchDoc(leftover.GetTitle)
        WScript.Sleep 500
        tries = tries + 1
    Loop
End If

' --- assemble ---------------------------------------------------------
Dim body, filletLines, j
filletLines = ""
For j = 0 To 3
    filletLines = filletLines & "        " & JStr("options" & filletOptions(j) & "_" & filletOK(j)) & ": " & JStr(filletDetail(j)) & "," & vbCrLf
Next
body = "{""version"": 1" & vbCrLf & _
       "  , ""revision"": " & JStr(revision) & vbCrLf & _
       "  , ""docCountBefore"": " & docCount & vbCrLf & _
       "  , ""userDocsWereOpen"": " & JBool(hadWork) & vbCrLf & _
       "  , ""probedOnScratchPart"": " & JBool(scratchPart) & vbCrLf & _
       "  , ""probeRanPartLevelProbes"": " & JBool(probesRun) & vbCrLf & _
       "  , ""template"": " & JStr(tpl) & vbCrLf & _
       "  , ""capabilities"": {" & vbCrLf & _
       "      ""sweptCut"": " & JBool(sweptCutOk) & vbCrLf & _
       "    , ""sweptCutProbed"": " & JBool(sweptCutProbed) & vbCrLf & _
       "    , ""revolveViaCreateDefinition"": " & JBool(revolveDefOk) & vbCrLf & _
       "    , ""filletOptionsBit2Required"": " & JBool(probeFilletOk) & vbCrLf & _
       "    , ""filletWorkingOption"": " & probeFilletOption & vbCrLf & _
       "    , ""filletOptionsProbed"": " & JBool(probesRun) & vbCrLf & _
       "    , ""negZFaceCut"": " & JBool(negZOK) & vbCrLf & _
       "    , ""negZFaceCutProbed"": " & JBool(probesRun) & vbCrLf & _
       "    , ""insertFeatureShellCreatesFeature"": " & JBool(abortShellOK) & vbCrLf & _
       "    , ""insertAxis2CreatesFeature"": " & JBool(abortAxisOK) & vbCrLf & _
       "    , ""shellAxisProbed"": " & JBool(abortProbed) & vbCrLf & _
       "    }" & vbCrLf & _
       "  , ""probeDetails"": {" & vbCrLf & _
       "      ""scratchBlock"": " & JStr(probeDetail) & vbCrLf & _
       "    , ""filletOptions"": {" & vbCrLf & filletLines & "        ""_end"": """" }" & vbCrLf & _
       "    , ""negZFaceCut"": " & JStr(negZDetail) & vbCrLf & _
       "    , ""insertFeatureShell"": " & JStr(abortShellDetail) & vbCrLf & _
       "    , ""insertAxis2"": " & JStr(abortAxisDetail) & vbCrLf & _
       "    }" & vbCrLf & _
       "  , ""api"": {" & vbCrLf & apiLines & vbCrLf & "    }" & vbCrLf & _
       "  , ""notes"": " & JStr("sweptCut=false routes threads to core+swept-boss; fillet needs the uniform-radius Options bit; a -Z face cannot host a cut sketch; shell/axis do not create the feature here; a file-level On Error Resume Next does not protect code inside a Sub") & vbCrLf & _
       "  }"
WriteOut outJson, body
E "probe done"
WScript.Quit 0
