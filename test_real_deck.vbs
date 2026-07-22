' Regression test against a user-provided real PPTX.
Option Explicit

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript test_real_deck.vbs source.pptx"
    WScript.Quit 2
End If

Dim sourcePath: sourcePath = WScript.Arguments(0)
Dim failures: failures = 0
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim ROOT: ROOT = fso.GetParentFolderName(WScript.ScriptFullName)
Dim ASSETS: ASSETS = ROOT & "\test_assets"
Dim BLUE: BLUE = ASSETS & "\new_blue.png"
Dim WORKING_COPY: WORKING_COPY = ASSETS & "\real_case_working.pptx"
If fso.FileExists(WORKING_COPY) Then fso.DeleteFile WORKING_COPY
fso.CopyFile sourcePath, WORKING_COPY, True

Function ReadUtf8(path)
    Dim stm: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
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
        If Left(ln, 9) <> "Attribute" Then out = out & ln & vbCrLf
    Next
    StripAttr = out
End Function

Function GetGeometry(shp)
    GetGeometry = Array(shp.Left, shp.Top, shp.Width, shp.Height, shp.Rotation, _
                        shp.HorizontalFlip, shp.VerticalFlip)
End Function

Sub Check(name, actual, expected, tolerance)
    If Abs(CSng(actual) - CSng(expected)) <= tolerance Then
        WScript.Echo "PASS " & name
    Else
        WScript.Echo "FAIL " & name & " actual=" & actual & " expected=" & expected
        failures = failures + 1
    End If
End Sub

Dim pp: Set pp = CreateObject("PowerPoint.Application")
pp.Visible = -1
On Error Resume Next

' Keep this regression independent from a PPAM that PowerPoint may have
' registered during a previous load test.
Dim installedAddin
For Each installedAddin In pp.AddIns
    If InStr(1, installedAddin.FullName, "PictureReplaceTools.ppam", 1) > 0 Then
        installedAddin.Loaded = False
    End If
Next

Dim pres: Set pres = pp.Presentations.Open(WORKING_COPY, 0, 0, -1)
If Err.Number <> 0 Then Fatal "open working copy"

Dim before1: before1 = GetGeometry(pres.Slides(1).Shapes(1))
Dim before2: before2 = GetGeometry(pres.Slides(2).Shapes(1))
Dim before3: before3 = GetGeometry(pres.Slides(3).Shapes(1))
Dim before4: before4 = GetGeometry(pres.Slides(4).Shapes(1))

Dim mainCode: mainCode = StripAttr(ReadUtf8(ROOT & "\modPictureReplace.bas"))
Dim comp: Set comp = pres.VBProject.VBComponents.Add(1)
If Err.Number <> 0 Then Fatal "add main module"
comp.Name = "modPictureReplace"
comp.CodeModule.AddFromString mainCode
If Err.Number <> 0 Then Fatal "inject main module"

Dim testCode
testCode = "Option Explicit" & vbCrLf & _
    "Public Function TestBulkReal() As Long" & vbCrLf & _
    "    Dim s As Shape" & vbCrLf & _
    "    Set s = Application.ActivePresentation.Slides(1).Shapes(1)" & vbCrLf & _
    "    TestBulkReal = ReplaceAllMatchingPictures(s, """ & BLUE & """)" & vbCrLf & _
    "End Function"

Dim testComp: Set testComp = pres.VBProject.VBComponents.Add(1)
testComp.Name = "modRealTest"
testComp.CodeModule.AddFromString testCode
If Err.Number <> 0 Then Fatal "inject real test module"

Err.Clear
Dim replaced: replaced = pp.Run("TestBulkReal")
If Err.Number <> 0 Then Fatal "run bulk replacement"
If CLng(replaced) = 3 Then
    WScript.Echo "PASS real deck matched exactly 3 pictures"
Else
    WScript.Echo "FAIL real deck replacement count=" & replaced & " expected=3"
    failures = failures + 1
End If

Dim after1: after1 = GetGeometry(pres.Slides(1).Shapes(1))
Dim after2: after2 = GetGeometry(pres.Slides(2).Shapes(1))
Dim after3: after3 = GetGeometry(pres.Slides(3).Shapes(1))
Dim after4: after4 = GetGeometry(pres.Slides(4).Shapes(1))
Dim statesBefore: statesBefore = Array(before1, before2, before3)
Dim statesAfter: statesAfter = Array(after1, after2, after3)
Dim slideIndex, fieldIndex
For slideIndex = 0 To 2
    For fieldIndex = 0 To 6
        Check "slide " & (slideIndex + 1) & " geometry field " & fieldIndex, _
              statesAfter(slideIndex)(fieldIndex), statesBefore(slideIndex)(fieldIndex), 0.55
    Next
Next
For fieldIndex = 0 To 6
    Check "slide 4 untouched field " & fieldIndex, after4(fieldIndex), before4(fieldIndex), 0.01
Next

pres.Slides(1).Shapes(1).Export ASSETS & "\real_pic1_after.png", 2
pres.Slides(2).Shapes(1).Export ASSETS & "\real_pic2_after.png", 2
pres.Slides(3).Shapes(1).Export ASSETS & "\real_pic3_after.png", 2
pres.Slides(4).Shapes(1).Export ASSETS & "\real_pic4_after.png", 2

pres.Close
pp.Quit
If failures = 0 Then
    WScript.Echo "REAL DECK TEST PASSED"
    WScript.Quit 0
Else
    WScript.Echo "REAL DECK FAILURES=" & failures
    WScript.Quit 1
End If

Sub Fatal(message)
    WScript.Echo "FATAL " & message & ": 0x" & Hex(Err.Number) & " " & Err.Description
    On Error Resume Next
    pp.Quit
    WScript.Quit 2
End Sub
