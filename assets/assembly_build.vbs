Option Explicit
' =====================================================================
'  assembly_build.vbs - the assembly build script used by
'  test/assembly-verify.mjs, and the reference example for driving an
'  assembly through dsh-solidworks.
'
'  It proves the two things the plugin now supplies:
'    * %SWASMTPL%     the assembly template (.asmdot), so no script has to
'                     hard-code an install-specific path;
'    * %SWPARTTPL%    the part template, unused here but exported with it.
'
'  It stages its part into %SWOUTDIR% first, because a string handed to
'  SolidWorks must be ASCII: FileSystemObject.CopyFile is not a COM call,
'  so it can read a non-ASCII source path, but OpenDoc cannot.
'
'  When run through solidworks_run, pass the part path as argument 0.
' =====================================================================

Const N_COMPONENTS = 3

Dim WSH, FSO, OUTDIR, ASMTPL, PARTSRC, PARTDST
Set WSH = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
OUTDIR = WSH.ExpandEnvironmentStrings("%SWOUTDIR%")
ASMTPL = WSH.ExpandEnvironmentStrings("%SWASMTPL%")
If Len(OUTDIR) = 0 Then OUTDIR = WSH.ExpandEnvironmentStrings("%TEMP%") & "\dsh-solidworks\out\"
If Right(OUTDIR, 1) <> "\" Then OUTDIR = OUTDIR & "\"
If Not FSO.FolderExists(OUTDIR) Then FSO.CreateFolder OUTDIR
PARTDST = OUTDIR & "asm_build_part.SLDPRT"

If WScript.Arguments.Count >= 1 Then
    PARTSRC = WScript.Arguments(0)
Else
    PARTSRC = PARTDST
End If

Dim swApp, part, asm, comp, i, refCount, f, tn, attempt

Sub E(s)
    WScript.Echo s
End Sub

Sub Main()
    On Error Resume Next
    E "asm template   = " & ASMTPL
    If Len(ASMTPL) = 0 Then
        E "FATAL: %SWASMTPL% is empty - configure assemblyTemplate or install an .asmdot"
        WScript.Quit 1
    End If
    If Not FSO.FileExists(ASMTPL) Then
        E "FATAL: assembly template missing: " & ASMTPL
        WScript.Quit 1
    End If

    If LCase(PARTSRC) <> LCase(PARTDST) Then
        Err.Clear
        FSO.CopyFile PARTSRC, PARTDST, True
        E "staged part    = " & PARTDST & " (copy err=" & Err.Number & ")"
    End If
    If Not FSO.FileExists(PARTDST) Then
        E "FATAL: no part to instance at " & PARTDST
        WScript.Quit 1
    End If

    ' Attaching to SolidWorks needs a retry: a first COM launch takes 30-60 s and
    ' `CreateObject` reports "ActiveX component can't create object" until the
    ' class factory is registered, which looks identical to "SolidWorks is not
    ' installed". Measured under a fresh process on a busy session: activation
    ' failed once at ~35 s and succeeded on the next attempt.
    attempt = 0
    Do While attempt < 12
        Err.Clear
        Set swApp = CreateObject("SldWorks.Application")
        If Err.Number = 0 Then Exit Do
        attempt = attempt + 1
        ' 429 = "ActiveX component can't create object" because SolidWorks has no
        ' free automation connection. That is machine state, not a script bug: an
        ' always-on Harness session holds one connection for its own tool calls,
        ' and a burst of extra clients (or earlier processes that have not been
        ' reaped) can exhaust the rest. It usually clears on its own, so a few
        ' spaced retries are worth it; beyond that the caller must act.
        E "CreateObject attempt " & attempt & " failed (err=" & Err.Number & "), retrying in 4 s"
        WScript.Sleep 4000
        If Err.Number = 429 And attempt >= 5 Then Exit Do
    Loop
    If IsEmpty(swApp) Or TypeName(swApp) = "Nothing" Then
        E "FATAL: cannot create SldWorks.Application after " & attempt & " attempt(s)"
        E "  err=429 means no free automation connection: another client holds it (a"
        E "  running Harness session always does). Close other clients or restart"
        E "  SolidWorks, then retry - the plugin's own tool calls still work."
        WScript.Quit 1
    End If
    swApp.Visible = True
    E "SolidWorks rev = " & swApp.RevisionNumber

    ' The 2-argument OpenDoc is the only file-open route that binds here;
    ' OpenDoc6 / LoadFile4 fail because their ByRef out-params do not.
    Err.Clear
    Set part = swApp.OpenDoc(PARTDST, 1)
    If Err.Number <> 0 Or IsEmpty(part) Then
        E "FATAL: OpenDoc(part) failed: err=" & Err.Number & " " & Err.Description
        WScript.Quit 1
    End If
    E "opened part    = " & part.GetTitle

    Err.Clear
    Set asm = swApp.NewDocument(ASMTPL, 0, 0, 0)
    WScript.Sleep 3000
    Set asm = swApp.ActiveDoc
    If Err.Number <> 0 Or IsEmpty(asm) Then
        E "FATAL: NewDocument(asmdot) failed: err=" & Err.Number
        WScript.Quit 1
    End If
    E "new assembly   = " & asm.GetTitle & "  (GetType=" & asm.GetType & ")"

    For i = 1 To N_COMPONENTS
        Err.Clear
        Set comp = asm.AddComponent5(PARTDST, 0, "", False, "", (i - 1) * 0.04, 0, 0)
        If Err.Number <> 0 Then
            E "AddComponent5 #" & i & " FAILED err=" & Err.Number & " " & Err.Description
        ElseIf IsEmpty(comp) Then
            E "AddComponent5 #" & i & " returned Empty"
        ElseIf TypeName(comp) = "Nothing" Then
            E "AddComponent5 #" & i & " returned Nothing"
        Else
            E "component " & i & "     = " & comp.Name2
        End If
    Next

    Err.Clear
    asm.EditRebuild3
    WScript.Sleep 600

    refCount = 0
    Set f = asm.FirstFeature
    Do While Not IsEmpty(f) And TypeName(f) <> "Nothing"
        tn = f.GetTypeName2
        If tn = "Reference" Then refCount = refCount + 1
        Set f = f.GetNextFeature
    Loop
    E "Reference feats= " & refCount & " (expected " & N_COMPONENTS & ")"

    Err.Clear
    E "SaveAs3 .SLDASM -> " & asm.SaveAs3(OUTDIR & "assembly_gate.SLDASM", 0, 1)
    asm.ShowNamedView2 "", 7
    asm.ViewZoomtofit2
    WScript.Sleep 900
    If asm.SaveBMP(OUTDIR & "assembly_gate_iso.bmp", 1400, 900) Then E "render iso OK"
    E "DONE"
End Sub

Main
WScript.Quit 0
