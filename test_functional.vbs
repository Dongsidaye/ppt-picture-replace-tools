' test_functional.vbs - end-to-end test of ReplacePictureKeepCrop + PPAM load.
' Pure ASCII output. Requires test_assets images built by gen step.
Option Explicit

Const PING = "picreplace_addin_ping.txt"

Dim failures: failures = 0
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim ROOT: ROOT = fso.GetParentFolderName(WScript.ScriptFullName)
Dim ASSETS: ASSETS = ROOT & "\test_assets"
Dim RED: RED = ASSETS & "\old_red.png"
Dim BLUE: BLUE = ASSETS & "\new_blue.png"
Dim GREEN: GREEN = ASSETS & "\new_green_wide.png"
Dim BULK_OLD: BULK_OLD = ASSETS & "\bulk_old_yellow.png"
Dim BULK_NEW: BULK_NEW = ASSETS & "\bulk_new_magenta.png"
Dim BULK_DECOY: BULK_DECOY = ASSETS & "\bulk_decoy_cyan.png"
Dim PPAM: PPAM = ROOT & "\dist\PictureReplaceTools.ppam"
Dim PPTM: PPTM = ASSETS & "\functest.pptm"

Sub Check(name, actual, expected, tol)
    Dim ok: ok = (Abs(CSng(actual) - CSng(expected)) <= tol)
    If ok Then
        WScript.Echo "PASS " & name & " (" & actual & ")"
    Else
        WScript.Echo "FAIL " & name & " actual=" & actual & " expected=" & expected
        failures = failures + 1
    End If
End Sub

Function GetState(shp)
    Dim c: Set c = shp.PictureFormat.Crop
    GetState = Array(shp.Left, shp.Top, shp.Width, shp.Height, shp.Rotation, _
                     c.PictureWidth, c.PictureHeight, c.PictureOffsetX, c.PictureOffsetY, _
                     c.ShapeWidth, c.ShapeHeight, shp.Name, shp.ZOrderPosition, _
                     shp.HorizontalFlip, shp.VerticalFlip)
End Function

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

If fso.FileExists(PPTM) Then fso.DeleteFile PPTM

Dim pp: Set pp = CreateObject("PowerPoint.Application")
pp.Visible = -1
WScript.Sleep 3000
On Error Resume Next

Dim pres: Set pres = pp.Presentations.Add()
If Err.Number <> 0 Then Fatal "Presentations.Add"

Dim s1: Set s1 = pres.Slides.Add(1, 12)
Dim s2: Set s2 = pres.Slides.Add(2, 12)
Dim s3: Set s3 = pres.Slides.Add(3, 12)
Dim s4: Set s4 = pres.Slides.Add(4, 12)
Dim s5: Set s5 = pres.Slides.Add(5, 12)
Dim s6: Set s6 = pres.Slides.Add(6, 12)
Dim s7: Set s7 = pres.Slides.Add(7, 12)
Dim s8: Set s8 = pres.Slides.Add(8, 12)
Dim s9: Set s9 = pres.Slides.Add(9, 12)
Dim s10: Set s10 = pres.Slides.Add(10, 12)
Dim s11: Set s11 = pres.Slides.Add(11, 12)

' Scenario A: same-size replace, crops + rotation
Dim a: Set a = s1.Shapes.AddPicture(RED, 0, -1, 100, 100)
a.Name = "TestPicA"
a.LockAspectRatio = 0
a.Width = 240: a.Height = 150
a.PictureFormat.CropLeft = 25
a.PictureFormat.CropRight = 35
a.PictureFormat.CropTop = 15
a.PictureFormat.CropBottom = 45
a.Rotation = 20

' Scenario B: wide aspect replace
Dim b: Set b = s2.Shapes.AddPicture(RED, 0, -1, 120, 80)
b.Name = "TestPicB"
b.LockAspectRatio = 0
b.Width = 200: b.Height = 120
b.PictureFormat.CropLeft = 0
b.PictureFormat.CropRight = 40
b.PictureFormat.CropTop = 10
b.PictureFormat.CropBottom = 10

' Scenario C: invalid path must leave the original untouched
Dim c: Set c = s3.Shapes.AddPicture(RED, 0, -1, 90, 70)
c.Name = "TestPicC"
c.PictureFormat.CropLeft = 12
c.PictureFormat.CropTop = 8

' Scenario D: flips, z-order and repeated replacement
Dim back: Set back = s4.Shapes.AddShape(1, 20, 20, 300, 220)
Dim d: Set d = s4.Shapes.AddPicture(RED, 0, -1, 110, 90)
d.Name = "TestPicD"
d.PictureFormat.CropRight = 20
d.Flip 0
d.Flip 1
Dim front: Set front = s4.Shapes.AddShape(1, 140, 120, 100, 80)

' Scenario E: three differently cropped instances of the same source image.
Dim e: Set e = s5.Shapes.AddPicture(BULK_OLD, 0, -1, 60, 50)
e.Name = "BulkRef": e.LockAspectRatio = 0: e.Width = 200: e.Height = 130
e.PictureFormat.CropLeft = 35: e.PictureFormat.CropTop = 18: e.Rotation = 7

Dim f: Set f = s6.Shapes.AddPicture(BULK_OLD, 0, -1, 150, 90)
f.Name = "BulkSecond": f.LockAspectRatio = 0: f.Width = 200: f.Height = 130
f.PictureFormat.CropRight = 52: f.PictureFormat.CropBottom = 24

Dim g: Set g = s7.Shapes.AddPicture(BULK_OLD, 0, -1, 70, 150)
g.Name = "BulkThird": g.LockAspectRatio = 0: g.Width = 200: g.Height = 130
g.PictureFormat.CropLeft = 12: g.PictureFormat.CropRight = 44
g.PictureFormat.CropTop = 31: g.Flip 0

' Same dimensions and crop as the target, but different pixels: must not match.
Dim h: Set h = s7.Shapes.AddPicture(BULK_DECOY, 0, -1, 390, 40)
h.Name = "BulkDecoy": h.LockAspectRatio = 0: h.Width = 160: h.Height = 100
h.PictureFormat.CropLeft = 12: h.PictureFormat.CropRight = 44

' Scenario F: single replacement from a picture copied to the clipboard.
Dim clipboardTarget: Set clipboardTarget = s8.Shapes.AddPicture(RED, 0, -1, 90, 75)
clipboardTarget.Name = "ClipboardTarget": clipboardTarget.LockAspectRatio = 0
clipboardTarget.Width = 230: clipboardTarget.Height = 145
clipboardTarget.PictureFormat.CropLeft = 28: clipboardTarget.PictureFormat.CropTop = 17
clipboardTarget.PictureFormat.CropBottom = 21: clipboardTarget.Rotation = -11
Dim clipboardSource: Set clipboardSource = s8.Shapes.AddPicture(BLUE, 0, -1, 600, 20)
clipboardSource.Name = "ClipboardSource"

' Scenario G: batch replacement from one clipboard picture.
Dim cb1: Set cb1 = s9.Shapes.AddPicture(BULK_OLD, 0, -1, 65, 55)
cb1.Name = "ClipboardBulkRef": cb1.LockAspectRatio = 0: cb1.Width = 210: cb1.Height = 135
cb1.PictureFormat.CropLeft = 22: cb1.PictureFormat.CropTop = 14
Dim cb2: Set cb2 = s10.Shapes.AddPicture(BULK_OLD, 0, -1, 150, 100)
cb2.Name = "ClipboardBulkSecond": cb2.LockAspectRatio = 0: cb2.Width = 210: cb2.Height = 135
cb2.PictureFormat.CropRight = 39: cb2.PictureFormat.CropBottom = 19
Dim cb3: Set cb3 = s11.Shapes.AddPicture(BULK_OLD, 0, -1, 85, 145)
cb3.Name = "ClipboardBulkThird": cb3.LockAspectRatio = 0: cb3.Width = 210: cb3.Height = 135
cb3.PictureFormat.CropLeft = 11: cb3.PictureFormat.CropRight = 31: cb3.Flip 1
Dim cbDecoy: Set cbDecoy = s11.Shapes.AddPicture(BULK_DECOY, 0, -1, 420, 45)
cbDecoy.Name = "ClipboardBulkDecoy"

Dim bA: bA = GetState(a)
Dim bB: bB = GetState(b)
Dim bC: bC = GetState(c)
Dim bD: bD = GetState(d)
Dim bE: bE = GetState(e)
Dim bF: bF = GetState(f)
Dim bG: bG = GetState(g)
Dim bH: bH = GetState(h)
Dim bClipboardTarget: bClipboardTarget = GetState(clipboardTarget)
Dim bCb1: bCb1 = GetState(cb1)
Dim bCb2: bCb2 = GetState(cb2)
Dim bCb3: bCb3 = GetState(cb3)
Dim bCbDecoy: bCbDecoy = GetState(cbDecoy)

' inject modules
Dim vbaMain: vbaMain = StripAttr(ReadUtf8(ROOT & "\modPictureReplace.bas"))
Dim testMod
testMod = "Option Explicit" & vbCrLf & _
    "Public Sub TestSameSize()" & vbCrLf & _
    "    Dim s As Shape" & vbCrLf & _
    "    Set s = Application.ActivePresentation.Slides(1).Shapes(""TestPicA"")" & vbCrLf & _
    "    If Not ReplacePictureKeepCrop(s, """ & BLUE & """) Then Err.Raise 9999, , ""A fail""" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Public Sub TestWide()" & vbCrLf & _
    "    Dim s As Shape" & vbCrLf & _
    "    Set s = Application.ActivePresentation.Slides(2).Shapes(""TestPicB"")" & vbCrLf & _
    "    If Not ReplacePictureKeepCrop(s, """ & GREEN & """) Then Err.Raise 9998, , ""B fail""" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Public Sub TestInvalid()" & vbCrLf & _
    "    Dim s As Shape" & vbCrLf & _
    "    Set s = Application.ActivePresentation.Slides(3).Shapes(""TestPicC"")" & vbCrLf & _
    "    If ReplacePictureKeepCrop(s, ""Z:\this\file\does-not-exist.png"") Then Err.Raise 9997, , ""C unexpectedly succeeded""" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Public Sub TestRepeatedFlip()" & vbCrLf & _
    "    Dim s As Shape" & vbCrLf & _
    "    Set s = Application.ActivePresentation.Slides(4).Shapes(""TestPicD"")" & vbCrLf & _
    "    If Not ReplacePictureKeepCrop(s, """ & BLUE & """) Then Err.Raise 9996, , ""D first fail""" & vbCrLf & _
    "    Set s = Application.ActivePresentation.Slides(4).Shapes(""TestPicD"")" & vbCrLf & _
    "    If Not ReplacePictureKeepCrop(s, """ & GREEN & """) Then Err.Raise 9995, , ""D second fail""" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Public Sub TestBulk()" & vbCrLf & _
    "    Dim s As Shape, n As Long" & vbCrLf & _
    "    Set s = Application.ActivePresentation.Slides(5).Shapes(""BulkRef"")" & vbCrLf & _
    "    n = ReplaceAllMatchingPictures(s, """ & BULK_NEW & """)" & vbCrLf & _
    "    If n <> 3 Then Err.Raise 9994, , ""Bulk count="" & CStr(n)" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Public Sub TestClipboardSingle()" & vbCrLf & _
    "    Dim s As Shape" & vbCrLf & _
    "    Application.ActivePresentation.Slides(8).Shapes(""ClipboardSource"").Copy" & vbCrLf & _
    "    DoEvents" & vbCrLf & _
    "    Set s = Application.ActivePresentation.Slides(8).Shapes(""ClipboardTarget"")" & vbCrLf & _
    "    If Not ReplacePictureFromClipboardKeepCrop(s) Then Err.Raise 9993, , ""Clipboard single fail""" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Public Sub TestClipboardBulk()" & vbCrLf & _
    "    Dim s As Shape, n As Long" & vbCrLf & _
    "    Application.ActivePresentation.Slides(8).Shapes(""ClipboardSource"").Copy" & vbCrLf & _
    "    DoEvents" & vbCrLf & _
    "    Set s = Application.ActivePresentation.Slides(9).Shapes(""ClipboardBulkRef"")" & vbCrLf & _
    "    n = ReplaceAllMatchingPicturesFromClipboard(s)" & vbCrLf & _
    "    If n <> 3 Then Err.Raise 9992, , ""Clipboard bulk count="" & CStr(n)" & vbCrLf & _
    "End Sub"

Dim comp: Set comp = pres.VBProject.VBComponents.Add(1)
If Err.Number <> 0 Then Fatal "VBComponents.Add"
comp.Name = "modPictureReplace"
comp.CodeModule.AddFromString vbaMain
If Err.Number <> 0 Then Fatal "AddFromString main"

Dim comp2: Set comp2 = pres.VBProject.VBComponents.Add(1)
comp2.Name = "modTest"
comp2.CodeModule.AddFromString testMod
If Err.Number <> 0 Then Fatal "AddFromString test"

pres.SaveAs PPTM, 25
If Err.Number <> 0 Then Fatal "SaveAs pptm"

' run replacements
Err.Clear
pp.Run "TestSameSize"
If Err.Number <> 0 Then
    WScript.Echo "FAIL run TestSameSize: 0x" & Hex(Err.Number) & " " & Err.Description
    failures = failures + 1
End If
Err.Clear
pp.Run "TestWide"
If Err.Number <> 0 Then
    WScript.Echo "FAIL run TestWide: 0x" & Hex(Err.Number) & " " & Err.Description
    failures = failures + 1
End If
Err.Clear
pp.Run "TestInvalid"
If Err.Number <> 0 Then
    WScript.Echo "FAIL run TestInvalid: 0x" & Hex(Err.Number) & " " & Err.Description
    failures = failures + 1
End If
Err.Clear
pp.Run "TestRepeatedFlip"
If Err.Number <> 0 Then
    WScript.Echo "FAIL run TestRepeatedFlip: 0x" & Hex(Err.Number) & " " & Err.Description
    failures = failures + 1
End If
Err.Clear
pp.Run "TestBulk"
If Err.Number <> 0 Then
    WScript.Echo "FAIL run TestBulk: 0x" & Hex(Err.Number) & " " & Err.Description
    failures = failures + 1
End If
Err.Clear
pp.Run "TestClipboardSingle"
If Err.Number <> 0 Then
    WScript.Echo "FAIL run TestClipboardSingle: 0x" & Hex(Err.Number) & " " & Err.Description
    failures = failures + 1
End If
Err.Clear
pp.Run "TestClipboardBulk"
If Err.Number <> 0 Then
    WScript.Echo "FAIL run TestClipboardBulk: 0x" & Hex(Err.Number) & " " & Err.Description
    failures = failures + 1
End If

Dim aA: aA = GetState(pres.Slides(1).Shapes(1))
Dim aB: aB = GetState(pres.Slides(2).Shapes(1))
Dim aC: aC = GetState(pres.Slides(3).Shapes("TestPicC"))
Dim aD: aD = GetState(pres.Slides(4).Shapes("TestPicD"))
Dim aE: aE = GetState(pres.Slides(5).Shapes("BulkRef"))
Dim aF: aF = GetState(pres.Slides(6).Shapes("BulkSecond"))
Dim aG: aG = GetState(pres.Slides(7).Shapes("BulkThird"))
Dim aH: aH = GetState(pres.Slides(7).Shapes("BulkDecoy"))
Dim aClipboardTarget: aClipboardTarget = GetState(pres.Slides(8).Shapes("ClipboardTarget"))
Dim aCb1: aCb1 = GetState(pres.Slides(9).Shapes("ClipboardBulkRef"))
Dim aCb2: aCb2 = GetState(pres.Slides(10).Shapes("ClipboardBulkSecond"))
Dim aCb3: aCb3 = GetState(pres.Slides(11).Shapes("ClipboardBulkThird"))
Dim aCbDecoy: aCbDecoy = GetState(pres.Slides(11).Shapes("ClipboardBulkDecoy"))

WScript.Echo "--- Scenario A (same-size) ---"
Check "A Left", aA(0), bA(0), 0.5
Check "A Top", aA(1), bA(1), 0.5
Check "A Width", aA(2), bA(2), 0.5
Check "A Height", aA(3), bA(3), 0.5
Check "A Rotation", aA(4), bA(4), 0.01
Check "A crop picture width", aA(5), bA(5), 0.05
Check "A crop picture height", aA(6), bA(6), 0.05
Check "A crop offset X", aA(7), bA(7), 0.05
Check "A crop offset Y", aA(8), bA(8), 0.05
Check "A crop shape width", aA(9), bA(9), 0.05
Check "A crop shape height", aA(10), bA(10), 0.05
Check "A Z", aA(12), bA(12), 0.1
If pres.Slides(1).Shapes(1).Name = "TestPicA" Then
    WScript.Echo "PASS A Name"
Else
    WScript.Echo "FAIL A Name: " & pres.Slides(1).Shapes(1).Name
    failures = failures + 1
End If

WScript.Echo "--- Scenario B (wide aspect) ---"
Check "B Left", aB(0), bB(0), 0.5
Check "B Width", aB(2), bB(2), 0.5
Check "B Height", aB(3), bB(3), 0.5
Check "B crop shape width", aB(9), bB(9), 0.05
Check "B crop shape height", aB(10), bB(10), 0.05
Check "B new image aspect", aB(5) / aB(6), 800 / 300, 0.01
Check "B normalized focus X", aB(7) / aB(5), bB(7) / bB(5), 0.01
Check "B normalized focus Y", aB(8) / aB(6), bB(8) / bB(6), 0.01
If aB(5) >= aB(9) And aB(6) >= aB(10) Then
    WScript.Echo "PASS B cover fills crop frame"
Else
    WScript.Echo "FAIL B cover does not fill crop frame"
    failures = failures + 1
End If

WScript.Echo "--- Scenario C (invalid path rollback) ---"
Check "C Left", aC(0), bC(0), 0.01
Check "C Top", aC(1), bC(1), 0.01
Check "C Width", aC(2), bC(2), 0.01
Check "C Height", aC(3), bC(3), 0.01
If pres.Slides(3).Shapes.Count = 1 Then
    WScript.Echo "PASS C original is the only remaining shape"
Else
    WScript.Echo "FAIL C leaked a replacement shape"
    failures = failures + 1
End If

WScript.Echo "--- Scenario D (flip, z-order, repeated replace) ---"
Check "D Left", aD(0), bD(0), 0.5
Check "D Top", aD(1), bD(1), 0.5
Check "D Width", aD(2), bD(2), 0.5
Check "D Height", aD(3), bD(3), 0.5
Check "D Z", aD(12), bD(12), 0.1
Check "D horizontal flip", aD(13), bD(13), 0.1
Check "D vertical flip", aD(14), bD(14), 0.1

WScript.Echo "--- Scenario E (bulk same-source, independent crops) ---"
Check "E Left", aE(0), bE(0), 0.5
Check "E Top", aE(1), bE(1), 0.5
Check "E Rotation", aE(4), bE(4), 0.01
Check "E crop picture width", aE(5), bE(5), 0.05
Check "E crop picture height", aE(6), bE(6), 0.05
Check "E crop offset X", aE(7), bE(7), 0.05
Check "E crop offset Y", aE(8), bE(8), 0.05
Check "F Left", aF(0), bF(0), 0.5
Check "F Top", aF(1), bF(1), 0.5
Check "F crop picture width", aF(5), bF(5), 0.05
Check "F crop picture height", aF(6), bF(6), 0.05
Check "F crop offset X", aF(7), bF(7), 0.05
Check "F crop offset Y", aF(8), bF(8), 0.05
Check "G Left", aG(0), bG(0), 0.5
Check "G Top", aG(1), bG(1), 0.5
Check "G crop picture width", aG(5), bG(5), 0.05
Check "G crop picture height", aG(6), bG(6), 0.05
Check "G crop offset X", aG(7), bG(7), 0.05
Check "G crop offset Y", aG(8), bG(8), 0.05
Check "G horizontal flip", aG(13), bG(13), 0.1
Check "decoy picture width untouched", aH(5), bH(5), 0.01
Check "decoy crop offset X untouched", aH(7), bH(7), 0.01

WScript.Echo "--- Scenario F (single clipboard replacement) ---"
Check "clipboard single Left", aClipboardTarget(0), bClipboardTarget(0), 0.5
Check "clipboard single Top", aClipboardTarget(1), bClipboardTarget(1), 0.5
Check "clipboard single Width", aClipboardTarget(2), bClipboardTarget(2), 0.5
Check "clipboard single Height", aClipboardTarget(3), bClipboardTarget(3), 0.5
Check "clipboard single Rotation", aClipboardTarget(4), bClipboardTarget(4), 0.01
Check "clipboard single crop width", aClipboardTarget(5), bClipboardTarget(5), 0.05
Check "clipboard single crop height", aClipboardTarget(6), bClipboardTarget(6), 0.05
Check "clipboard single crop offset X", aClipboardTarget(7), bClipboardTarget(7), 0.05
Check "clipboard single crop offset Y", aClipboardTarget(8), bClipboardTarget(8), 0.05

WScript.Echo "--- Scenario G (bulk clipboard replacement) ---"
Check "clipboard bulk 1 crop width", aCb1(5), bCb1(5), 0.05
Check "clipboard bulk 1 crop offset X", aCb1(7), bCb1(7), 0.05
Check "clipboard bulk 2 crop width", aCb2(5), bCb2(5), 0.05
Check "clipboard bulk 2 crop offset Y", aCb2(8), bCb2(8), 0.05
Check "clipboard bulk 3 crop width", aCb3(5), bCb3(5), 0.05
Check "clipboard bulk 3 vertical flip", aCb3(14), bCb3(14), 0.1
Check "clipboard batch decoy untouched", aCbDecoy(5), bCbDecoy(5), 0.01

pres.Slides(1).Export ASSETS & "\slide1_after.png", "PNG", 960, 540
pres.Slides(2).Export ASSETS & "\slide2_after.png", "PNG", 960, 540
pres.Slides(4).Export ASSETS & "\slide4_after.png", "PNG", 960, 540
pres.Slides(5).Export ASSETS & "\slide5_after.png", "PNG", 960, 540
pres.Slides(6).Export ASSETS & "\slide6_after.png", "PNG", 960, 540
pres.Slides(7).Export ASSETS & "\slide7_after.png", "PNG", 960, 540
pres.Slides(8).Shapes("ClipboardTarget").Export ASSETS & "\clipboard_single_after.png", 2
pres.Slides(9).Shapes("ClipboardBulkRef").Export ASSETS & "\clipboard_bulk1_after.png", 2
pres.Slides(10).Shapes("ClipboardBulkSecond").Export ASSETS & "\clipboard_bulk2_after.png", 2
pres.Slides(11).Shapes("ClipboardBulkThird").Export ASSETS & "\clipboard_bulk3_after.png", 2
pres.Slides(11).Shapes("ClipboardBulkDecoy").Export ASSETS & "\clipboard_decoy_after.png", 2

pres.Close

' ---- PPAM load test ----
WScript.Echo "--- PPAM load test ---"
Dim pingPath: pingPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\" & PING
If fso.FileExists(pingPath) Then fso.DeleteFile pingPath
Err.Clear
Dim addin: Set addin = pp.AddIns.Add(PPAM)
If Err.Number <> 0 Then
    WScript.Echo "FAIL AddIns.Add: 0x" & Hex(Err.Number) & " " & Err.Description
    failures = failures + 1
Else
    addin.Loaded = True
    If Err.Number <> 0 Then
        WScript.Echo "FAIL AddIn.Loaded=True: 0x" & Hex(Err.Number) & " " & Err.Description
        failures = failures + 1
    Else
        WScript.Echo "PASS AddIn loaded (Loaded=" & addin.Loaded & ")"
    End If
    Err.Clear
    pp.Run "modPictureReplace.PictureReplace_Ping"
    If Err.Number <> 0 Then
        WScript.Echo "FAIL run ping: 0x" & Hex(Err.Number) & " " & Err.Description
        failures = failures + 1
    Else
        WScript.Sleep 500
        If fso.FileExists(pingPath) Then
            WScript.Echo "PASS macro in PPAM executed (ping file created)"
        Else
            WScript.Echo "FAIL ping file not created"
            failures = failures + 1
        End If
    End If
End If

pp.Quit
If failures = 0 Then
    WScript.Echo "ALL TESTS PASSED"
    WScript.Quit 0
Else
    WScript.Echo "FAILURES: " & failures
    WScript.Quit 1
End If

Sub Fatal(msg)
    WScript.Echo "FATAL " & msg & ": 0x" & Hex(Err.Number) & " " & Err.Description
    On Error Resume Next
    pp.Quit
    WScript.Quit 2
End Sub
