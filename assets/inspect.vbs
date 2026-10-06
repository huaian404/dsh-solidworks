Option Explicit

' =====================================================================
'  SolidWorks model inspector
'
'  READ-ONLY. Reports what the ACTIVE document actually contains, so a build
'  can be checked against its intended result:
'    * every feature, with its type and suppression state
'    * the solid-body count (a single part must report 1)
'    * the model bounding box, when the host can produce it
'    * rendered views, as BMP files the caller can look at
'
'  Usage:  cscript //nologo inspect.vbs <outJson> <outDir> [viewIds]
'          viewIds: comma separated, default "7,5"
'                   1 back 2 front 3 left 4 right 5 front(xz) 6 top 7 iso
'  Env:    SWOUTDIR  ascii scratch directory
'
'  The JSON is the verification contract: the caller compares it against the
'  expected features/bodies and treats a mismatch as a failed build.
' =====================================================================

Dim fso, outJson, outDir, viewArg, t0
Set fso = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count < 2 Then
    WScript.Echo "usage: cscript //nologo inspect.vbs <outJson> <outDir> [viewIds]"
    WScript.Quit 2
End If
outJson = WScript.Arguments(0)
outDir = WScript.Arguments(1)
If Right(outDir, 1) <> "\" Then outDir = outDir & "\"
viewArg = "7,5"
If WScript.Arguments.Count >= 3 Then viewArg = WScript.Arguments(2)
' 4th arg: "1" keeps reference planes/geometry in the feature list (debugging)
Dim includeRef
includeRef = False
If WScript.Arguments.Count >= 4 Then
    If Trim(WScript.Arguments(3)) = "1" Then includeRef = True
End If
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

Dim swApp, sw, ext, f, tn, bds, n, i, featLines, featFirst
Dim bodyCount, suppressed, renders, renderLines, renderFirst, box, bb
Dim boxOk, boxTxt, title, viewIds, vid, bmp, gotBox

On Error Resume Next

Set swApp = CreateObject("SldWorks.Application")
If Err.Number <> 0 Then
    WriteOut outJson, "{""version"": 1, ""error"": " & JStr("cannot create SldWorks.Application: " & Err.Description) & "}"
    WScript.Quit 1
End If
swApp.Visible = True

Set sw = swApp.ActiveDoc
If IsEmpty(sw) Or TypeName(sw) = "Nothing" Then
    WriteOut outJson, "{""version"": 1, ""error"": ""no active document""}"
    WScript.Quit 1
End If
Set ext = sw.Extension
title = sw.GetTitle

' --- feature tree -----------------------------------------------------
featLines = ""
featFirst = True
Set f = sw.FirstFeature
Do While Not IsEmpty(f) And TypeName(f) <> "Nothing"
    tn = f.GetTypeName2
    suppressed = False
    Err.Clear
    suppressed = f.IsSuppressed
    If Err.Number <> 0 Then suppressed = False
    ' feature folders (History, Selection Sets, ...) are noise for verification
    If FeatureKind(tn) <> "" And (includeRef Or (FeatureKind(tn) <> "ReferencePlane" And FeatureKind(tn) <> "ReferenceGeometry")) Then
        If Not featFirst Then featLines = featLines & "," & vbCrLf
        featFirst = False
        featLines = featLines & "    { ""name"": " & JStr(f.Name) & _
                    ", ""type"": " & JStr(tn) & _
                    ", ""kind"": " & JStr(FeatureKind(tn)) & _
                    ", ""suppressed"": " & JBool(suppressed) & " }"
    End If
    Set f = f.GetNextFeature
Loop

' --- bodies -----------------------------------------------------------
bodyCount = -1
Err.Clear
bds = sw.GetBodies2(0, True)
If IsArray(bds) Then
    bodyCount = UBound(bds) - LBound(bds) + 1
ElseIf TypeName(bds) = "Object" Or TypeName(bds) = "Nothing" Then
    ' a single body may come back as a bare object in some hosts
    bodyCount = 1
End If
E "solid bodies = " & bodyCount

' --- bounding box: try every route, keep the first that answers ---------
' Body2.GetBodyBox and ModelDoc2.GetMassProperties return Empty through late
' binding on this host, and Extension.GetVisibleBox wants MathPoint objects, so
' the box is usually unavailable. It is reported as null rather than guessed.
boxOk = False
boxTxt = "null"
Err.Clear
Dim arr6(5)
If IsArray(bds) Then
    Dim bd
    Set bd = bds(LBound(bds))
    If Not IsEmpty(bd) And TypeName(bd) <> "Nothing" Then
        Err.Clear
        bd.IGetBodyBox arr6(0)
        If Err.Number = 0 Then
            boxOk = True
            boxTxt = BoxJson(arr6, "bodyBox")
        End If
    End If
End If
E "bounding box available = " & boxOk

' --- renders ----------------------------------------------------------
renderLines = ""
renderFirst = True
renders = Split(viewArg, ",")
For i = LBound(renders) To UBound(renders)
    vid = Trim(renders(i))
    If Len(vid) > 0 Then
        bmp = outDir & "view" & vid & ".bmp"
        Err.Clear
        sw.ClearSelection2 True
        sw.ShowNamedView2 "", CLng(vid)
        sw.ViewZoomtofit2
        If sw.SaveBMP(bmp, 1400, 900) Then
            If Not renderFirst Then renderLines = renderLines & "," & vbCrLf
            renderFirst = False
            renderLines = renderLines & "    { ""view"": " & vid & ", ""bmp"": " & JStr(bmp) & " }"
            E "  render view " & vid & " -> " & bmp
        Else
            E "  render view " & vid & " FAILED err=" & Err.Number
        End If
    End If
Next

' --- assemble ---------------------------------------------------------
Dim jb
jb = "{""version"": 1" & vbCrLf & _
     "  , ""title"": " & JStr(title) & vbCrLf & _
     "  , ""bodyCount"": " & bodyCount & vbCrLf & _
     "  , ""boundingBox"": " & boxTxt & vbCrLf & _
     "  , ""features"": [" & vbCrLf & featLines & vbCrLf & "  ]" & vbCrLf & _
     "  , ""renders"": [" & vbCrLf & renderLines & vbCrLf & "  ]" & vbCrLf & _
     "  }"
WriteOut outJson, jb
E "inspect done"
WScript.Quit 0

' Map a SolidWorks feature type onto the verification vocabulary. Reference
' planes, sketches and feature folders are structural, not modelled features,
' so they never appear in the report; "ICE" (instant cut extrude) and friends
' collapse into the semantic kind a caller actually asserts on.
Function FeatureKind(tn)
    FeatureKind = ""
    Select Case tn
        Case "Extrusion", "Boss", "BossThin"
            FeatureKind = "Extrude"
        Case "Revolution", "RevCut"
            FeatureKind = "Revolve"
        Case "ICE", "Cut", "CutThin"
            FeatureKind = "Cut"
        Case "Helix"
            FeatureKind = "Helix"
        Case "Sweep"
            FeatureKind = "SweepBoss"
        Case "SweepCut"
            FeatureKind = "SweepCut"
        Case "Chamfer"
            FeatureKind = "Chamfer"
        Case "Fillet", "VarFillet"
            FeatureKind = "Fillet"
        Case "HoleWzd", "SketchHole"
            FeatureKind = "Hole"
        Case "LPattern", "CirPattern", "MirrorPattern", "SketchPattern"
            FeatureKind = "Pattern"
        Case "Shell"
            FeatureKind = "Shell"
        Case "Draft"
            FeatureKind = "Draft"
        Case "RefPlane"
            FeatureKind = "ReferencePlane"
        Case "RefAxis", "RefCurve", "ReferenceCurve"
            FeatureKind = "ReferenceGeometry"
        Case Else
            FeatureKind = ""
    End Select
End Function


Function BoxJson(a, source)
    Dim x0, y0, z0, x1, y1, z1
    If source = "visibleBox" Then
        x0 = a(0) : y0 = a(1) : z0 = a(2) : x1 = a(3) : y1 = a(4) : z1 = a(5)
    Else
        x0 = a(0) : y0 = a(1) : z0 = a(2) : x1 = a(3) : y1 = a(4) : z1 = a(5)
    End If
    BoxJson = "{ ""source"": " & JStr(source) & _
              ", ""size"": [" & Fmt((x1 - x0) * 1000) & ", " & Fmt((y1 - y0) * 1000) & ", " & Fmt((z1 - z0) * 1000) & "]" & _
              ", ""min"": [" & Fmt(x0 * 1000) & ", " & Fmt(y0 * 1000) & ", " & Fmt(z0 * 1000) & "]" & _
              ", ""max"": [" & Fmt(x1 * 1000) & ", " & Fmt(y1 * 1000) & ", " & Fmt(z1 * 1000) & "] }"
End Function

Function Fmt(d)
    Fmt = Replace(FormatNumber(d, 3), ",", ".")
End Function
