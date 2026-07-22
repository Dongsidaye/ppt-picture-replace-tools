' test_load.vbs - load one PPAM path and execute its ping macro.
Option Explicit

If WScript.Arguments.Count <> 1 Then
    WScript.Echo "usage: cscript test_load.vbs <ppam>"
    WScript.Quit 2
End If

Dim path: path = WScript.Arguments(0)
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(path) Then
    WScript.Echo "missing: " & path
    WScript.Quit 2
End If

Dim pp: Set pp = CreateObject("PowerPoint.Application")
pp.Visible = -1
On Error Resume Next

Dim addin: Set addin = pp.AddIns.Add(path)
If Err.Number <> 0 Then Fail "AddIns.Add"

Err.Clear
addin.Loaded = True
If Err.Number <> 0 Then Fail "Loaded=True"

Err.Clear
pp.Run "modPictureReplace.PictureReplace_Ping"
If Err.Number <> 0 Then Fail "Run ping"

WScript.Echo "PASS loaded=" & addin.Loaded & " name=" & addin.Name
pp.Quit
WScript.Quit 0

Sub Fail(stepName)
    WScript.Echo "FAIL " & stepName & " 0x" & Hex(Err.Number) & " " & Err.Description
    On Error Resume Next
    pp.Quit
    WScript.Quit 1
End Sub
