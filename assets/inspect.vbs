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
'    * for an ASSEMBLY: the component count, the mate count and the number
'      of suppressed features, which is what a verification gate needs
'
'  Usage:  cscript //nologo inspect.vbs <outJson> <outDir> [options]
'          options, either positional or key=value:
'            views=7,5        1 back 2 front 3 left 4 right 5 front(xz) 6 top 7 iso
'            includeReference=1   also list reference planes/axes (debugging)
'  Env:    SWOUTDIR  ascii scratch directory
'
'  The JSON is the verification contract: the caller compares it against the
'  expected features/bodies and treats a mismatch as a failed build.
'
'  MEASURED BINDING FACTS on this host (SW 2026 SP2.1, cscript late binding),
'  which is why the assembly section is shaped the way it is:
'    * ModelDoc2.GetType DOES bind: 1 = part, 2 = assembly, 3 = drawing.
'    * AssemblyDoc.GetComponents(True) binds and its LENGTH is correct, but
'      every element comes back Empty - no Component2 property (Name2,
'      GetPathName, IsSuppressed, Transform2) can be read at all, and
'      RootComponent.GetChildren has the same limitation. Component names,
'      paths, suppressed/fixed state and transforms are therefore NOT
'      reported, and a missing component cannot be detected by path.
'    * MATES ARE INVISIBLE. The tree carries one "MateGroup" container whose
'      GetFirstSubFeature is Nothing; mate features never appear as top-level
'      features (verified after inserting two components, on an assembly that
'      had been saved from a real mate attempt); and AssemblyDoc.GetMates,
'      AssemblyDoc.GetMateCount and FeatureManager.GetMates all fail with 438
'      "object doesn't support this property or method". mateCount is therefore
'      reported as null with mateCountAvailable = false rather than as 0 or 1,
'      because either number would be a guess.
'    * What IS reliable: the component count, one Reference feature per
'      instance, and IsSuppressed per feature.
'    * The CORRECT document must be ACTIVE. Neither OpenDoc nor NewDocument
'      makes the document it returns active - a part opened while an assembly
'      is in front stays behind it - and ActivateDoc3 fails with 13
'      "type mismatch", so once an assembly is active you cannot switch back
'      to a part from VBScript. Build (or open) the document you intend to
'      verify LAST, and treat the reported title as the only proof that the
'      right document was measured.
' =====================================================================

Dim fso, outJson, outDir, viewArg, t0
Set fso = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count < 2 Then
    WScript.Echo "usage: cscript //nologo inspect.vbs <outJson> <outDir> [views=..] [includeReference=1]"
    WScript.Quit 2
End If
outJson = WScript.Arguments(0)
outDir = WScript.Arguments(1)
If Right(outDir, 1) <> "\" Then outDir = outDir & "\"
viewArg = "7,5"
Dim includeRef
includeRef = False
Dim sawViews, sawInclude
sawViews = False
sawInclude = False
t0 = Timer()
' Options are parsed as key=value, with the value inspected explicitly. The
' earlier form compared `Left(arg, n) = "literal"`, which has two traps that
' both bit: a literal length that is off by one (Left(argLow, 9) against the
' 10-character "includeref") and a fall-through branch that then treated the
' option as a legacy positional view id. Anything that is not key=value is kept
' as a positional fallback for the historical call shape.
Dim argi, argText, argLow, argKey, argVal, eqPos
For argi = 2 To WScript.Arguments.Count - 1
    argText = Trim(WScript.Arguments(argi))
    argLow = LCase(argText)
    eqPos = InStr(argLow, "=")
    If eqPos > 1 Then
        argKey = Left(argLow, eqPos - 1)
        argVal = Mid(argText, eqPos + 1)
        If argKey = "views" Then
            viewArg = argVal
            sawViews = True
        ElseIf argKey = "includereference" Or argKey = "includeref" Then
            includeRef = (argVal = "1") Or (LCase(argVal) = "true") Or (LCase(argVal) = "yes")
            sawInclude = True
        End If
    ElseIf argLow = "1" Then
        ' legacy positional form: arg 3 = includeReference
        includeRef = True
    ElseIf Len(argText) > 0 And Not sawViews Then
        ' legacy positional form: arg 2 = viewIds
        viewArg = argText
    End If
Next
' A view id must be numeric. Reporting the bad id as if it were a view is what
' made the earlier defect silent: the render was written to "view<option>.bmp"
' and the JSON became unparseable downstream.
Dim vi, vidTok, viewTokens
viewTokens = Split(viewArg, ",")
For vi = LBound(viewTokens) To UBound(viewTokens)
    vidTok = Trim(viewTokens(vi))
    If Len(vidTok) = 0 Then
        WriteOut outJson, "{""version"": 1, ""error"": " & JStr("empty view id in views=" & viewArg) & "}"
        WScript.Quit 2
    End If
    If Not IsNumeric(vidTok) Then
        WriteOut outJson, "{""version"": 1, ""error"": " & JStr("non-numeric view id """ & vidTok & """ in views=" & viewArg) & "}"
        WScript.Quit 2
    End If
Next

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
Dim docType, docTypeName, asmLines, compCount, mateFeatureCount, refCount
Dim suppressedCount, isAsm

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

' --- document type ----------------------------------------------------
' GetType is the semantic switch (1 part / 2 assembly / 3 drawing) and does
' bind here. The TypeName of the late-bound object is kept as a cross-check,
' because a document that reports neither is exactly what a caller needs to
' see rather than a guessed "part".
docType = -1
Err.Clear
docType = sw.GetType
If Err.Number <> 0 Then docType = -1
Select Case docType
    Case 1
        docTypeName = "part"
    Case 2
        docTypeName = "assembly"
    Case 3
        docTypeName = "drawing"
    Case Else
        docTypeName = "unknown"
End Select
isAsm = (docType = 2)
E "document type = " & docType & " (" & docTypeName & "), late-bound as " & TypeName(sw)

' --- feature tree -----------------------------------------------------
' The loop counts what an assembly report needs while it walks: one
' "Reference" feature per component instance, the feature types that name a
' mate, and the features SolidWorks marks suppressed (an unresolved component
' is suppressed, so this count is the only missing-reference signal that
' survives the binding limits recorded in the header).
featLines = ""
featFirst = True
mateFeatureCount = 0
refCount = 0
suppressedCount = 0
Set f = sw.FirstFeature
Do While Not IsEmpty(f) And TypeName(f) <> "Nothing"
    tn = f.GetTypeName2
    suppressed = False
    Err.Clear
    suppressed = f.IsSuppressed
    If Err.Number <> 0 Then suppressed = False
    If suppressed Then suppressedCount = suppressedCount + 1
    If tn = "Reference" Then refCount = refCount + 1
    ' "MateGroup" is the container and is present even in an assembly with no
    ' mate at all, so it is deliberately NOT counted as one.
    If MateKind(tn) <> "" Then mateFeatureCount = mateFeatureCount + 1
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

' --- assembly metrics -------------------------------------------------
' GetComponents(True) binds and its length is authoritative when the
' document is an assembly: an ASSEMBLY WITH NO COMPONENT AT ALL makes it
' fail (error 424), which is itself the report - 0, not a failure. The
' feature-tree Reference count is carried alongside as an independent
' cross-check, so a caller can see the two agree instead of trusting one.
asmLines = "null"
If isAsm Then
    compCount = -1
    Err.Clear
    Dim comps
    comps = sw.GetComponents(True)
    If Err.Number = 0 Then
        If IsArray(comps) Then
            compCount = UBound(comps) - LBound(comps) + 1
        ElseIf Not IsEmpty(comps) Then
            compCount = 1
        End If
    Else
        compCount = 0
    End If
    E "assembly components = " & compCount & ", Reference features = " & refCount
    If mateFeatureCount > 0 Then
        E "mate features = " & mateFeatureCount & " (unexpected: mates are normally invisible here)"
    Else
        E "mate features = none visible (GetMates / GetMateCount / GetFirstSubFeature do not bind)"
    End If
    E "suppressed features = " & suppressedCount
    ' mateCount is null unless a mate feature was actually seen in the tree:
    ' reporting 0 would claim "no mates" from a route that cannot see them
    ' either way, and the empty MateGroup must not be mistaken for one.
    asmLines = "{ ""componentCount"": " & compCount & _
               ", ""referenceFeatureCount"": " & refCount & _
               ", ""mateCount"": " & MateCountJson(mateFeatureCount) & _
               ", ""mateCountAvailable"": " & JBool(mateFeatureCount > 0) & _
               ", ""suppressedFeatureCount"": " & suppressedCount & _
               ", ""componentDetailsAvailable"": false }"
End If

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
     "  , ""documentType"": " & JStr(docTypeName) & vbCrLf & _
     "  , ""documentTypeCode"": " & docType & vbCrLf & _
     "  , ""title"": " & JStr(title) & vbCrLf & _
     "  , ""bodyCount"": " & bodyCount & vbCrLf & _
     "  , ""boundingBox"": " & boxTxt & vbCrLf & _
     "  , ""assembly"": " & asmLines & vbCrLf & _
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


' A mate feature's type name, or "" when the feature is not a mate. The
' "MateGroup" container is excluded on purpose: SolidWorks puts it in every
' assembly, including one with no mate at all, so counting it reported "1 mate"
' for an assembly that had none. The per-type names are kept distinct from the
' semantic FeatureKind vocabulary above because the verification contract for a
' part must not change just because mates exist.
Function MateKind(tn)
    MateKind = ""
    Select Case tn
        Case "Mate", "Coincident", "Concentric", "Parallel", "Perpendicular", _
             "Tangent", "Distance", "Angle", "Lock", "Gear", "Screw", _
             "Universal", "Symmetry", "Width", "Ratio", "RackPinion", _
             "LinearCoupler", "LimitDistance", "LimitAngle", "Hinge", "Slot"
            MateKind = tn
        Case Else
            MateKind = ""
    End Select
End Function

' null when nothing was measurable, the number otherwise. A count of 0 from a
' route that cannot see mates at all is a false claim of "no mates", so the
' absence of evidence is reported as null instead.
Function MateCountJson(n)
    If n <= 0 Then
        MateCountJson = "null"
    Else
        MateCountJson = CStr(n)
    End If
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
