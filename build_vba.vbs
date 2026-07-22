' build_vba.vbs - inject VBA module via PowerPoint COM and save PPAM skeleton.
' Pure ASCII: all Chinese text lives in the .bas file (read as UTF-8 via ADODB).
Option Explicit

Const ROOT = "C:\Users\Administrator\Documents\powerpoint\pic_replace_addin"
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim basPath: basPath = ROOT & "\modPictureReplace.bas"
Dim ppamPath: ppamPath = ROOT & "\dist\PictureReplaceTools.pptm"

Dim vba: vba = StripAttr(ReadUtf8(basPath))

If Not fso.FolderExists(ROOT & "\dist") Then fso.CreateFolder ROOT & "\dist"
If fso.FileExists(ppamPath) Then fso.DeleteFile ppamPath

Dim pp: Set pp = CreateObject("PowerPoint.Application")
pp.Visible = -1
WScript.Sleep 3000
On Error Resume Next

Dim pres
Dim attempt
For attempt = 1 To 5
    Err.Clear
    Set pres = pp.Presentations.Add()
    If Err.Number = 0 Then Exit For
    WScript.Sleep 2000
Next
CheckErr "Presentations.Add"

Dim comp
For attempt = 1 To 5
    Err.Clear
    Set comp = pres.VBProject.VBComponents.Add(1)  ' vbext_ct_StdModule
    If Err.Number = 0 Then Exit For
    WScript.Sleep 2000
Next
CheckErr "VBComponents.Add"
comp.Name = "modPictureReplace"
comp.CodeModule.AddFromString vba
CheckErr "AddFromString"
WScript.Echo "module lines: " & comp.CodeModule.CountOfLines
If comp.CodeModule.CountOfLines < 50 Then
    WScript.Echo "BUILD FAILED: module too small"
    pp.Quit
    WScript.Quit 1
End If

pres.SaveAs ppamPath, 25   ' ppSaveAsOpenXMLPresentationMacroEnabled
CheckErr "SaveAs"
pres.Close
pp.Quit
WScript.Echo "PPAM saved: " & ppamPath
WScript.Quit 0

Sub CheckErr(step)
    If Err.Number <> 0 Then
        WScript.Echo "BUILD FAILED at " & step & ": 0x" & Hex(Err.Number) & " " & Err.Description
        On Error Resume Next
        pp.Quit
        WScript.Quit 1
    End If
End Sub

Function ReadUtf8(path)
    Dim stm: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    ReadUtf8 = stm.ReadText(-1)
    stm.Close
End Function

Function StripAttr(src)
    Dim lines: lines = Split(Replace(src, vbCrLf, vbLf), vbLf)
    Dim out: out = ""
    Dim i, ln
    For i = 0 To UBound(lines)
        ln = lines(i)
        If Len(ln) > 0 Then
            If Right(ln, 1) = vbCr Then ln = Left(ln, Len(ln) - 1)
        End If
        If Left(ln, 9) <> "Attribute" Then
            out = out & ln & vbCrLf
        End If
    Next
    StripAttr = out
End Function
