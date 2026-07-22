Attribute VB_Name = "modPictureReplace"
'=============================================================
' 图片原位替换工具 (Replace Picture In Place)
'
' 功能：用新图片替换选中的图片，保持原有位置、大小、旋转、
'       翻转、层级以及裁剪比例与图片格式效果。
'
' 实现策略：嵌入新图片并读取 PictureFormat.Crop 的图片尺寸、
'           裁剪框尺寸与焦点偏移。相近宽高比精确复用几何；
'           差异较大时用 cover 铺满并保持归一化焦点位置。
'=============================================================
Option Explicit

' 容忍 PowerPoint 对同一媒体的亚像素重采样差异；同时限制局部最大差异，
' 避免把只有整体轮廓相似、但局部内容不同的图片误判为同源。
Private Const PREVIEW_GRID_SIZE As Long = 32
Private Const MAX_MEAN_CHANNEL_DIFF As Double = 0.08
Private Const MAX_PIXEL_RGB_SUM As Long = 12
Private Const WIA_FORMAT_BMP As String = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"

'---------- Ribbon 回调 ----------
Public Sub RibbonReplacePicture(control As Object)
    ReplaceSelectedPicture
End Sub

Public Sub RibbonReplaceAllMatchingPictures(control As Object)
    ReplaceAllMatchingPicturesUI
End Sub

'---------- 主入口：替换当前选中的图片 ----------
Public Sub ReplaceSelectedPicture()
    Dim shp As Shape
    Set shp = TryGetSelectedPicture()
    If shp Is Nothing Then Exit Sub   ' 具体原因已弹窗提示

    Dim imgPath As String
    imgPath = PickImageFile()
    If Len(imgPath) = 0 Then Exit Sub

    If Not ReplacePictureKeepCrop(shp, imgPath) Then
        MsgBox "替换失败：无法读取所选图片或新图片文件。" & vbCrLf & _
               "请确认选中了单张图片，且文件格式受支持。", _
               vbCritical, "原位替换图片"
    End If
End Sub

'---------- 批量入口：替换当前演示文稿中所有同源图片 ----------
Public Sub ReplaceAllMatchingPicturesUI()
    Dim referenceShp As Shape
    Set referenceShp = TryGetSelectedPicture()
    If referenceShp Is Nothing Then Exit Sub

    Dim imgPath As String
    imgPath = PickImageFile()
    If Len(imgPath) = 0 Then Exit Sub

    Dim matched As Long, failed As Long, skipped As Long
    Dim succeeded As Long
    succeeded = ReplaceAllMatchingPicturesCore(referenceShp, imgPath, matched, failed, skipped)

    If matched = 0 Then
        MsgBox "无法识别基准图片，或演示文稿中没有找到匹配图片。", _
               vbExclamation, "批量替换同图"
    ElseIf failed = 0 Then
        MsgBox "批量替换完成：共替换 " & succeeded & " 张图片。" & _
               IIf(skipped > 0, vbCrLf & "跳过组合对象 " & skipped & " 个。", ""), _
               vbInformation, "批量替换同图"
    Else
        MsgBox "批量替换完成：成功 " & succeeded & " 张，失败 " & failed & " 张。" & _
               IIf(skipped > 0, vbCrLf & "跳过组合对象 " & skipped & " 个。", ""), _
               vbExclamation, "批量替换同图"
    End If
End Sub

' 供自动化测试或其他宏调用；返回成功替换数量。
Public Function ReplaceAllMatchingPictures(ByVal referenceShp As Shape, _
                                           ByVal imgPath As String) As Long
    Dim matched As Long, failed As Long, skipped As Long
    ReplaceAllMatchingPictures = ReplaceAllMatchingPicturesCore( _
        referenceShp, imgPath, matched, failed, skipped)
End Function

Private Function ReplaceAllMatchingPicturesCore(ByVal referenceShp As Shape, _
                                                ByVal imgPath As String, _
                                                ByRef matched As Long, _
                                                ByRef failed As Long, _
                                                ByRef skipped As Long) As Long
    On Error GoTo Fail

    Dim referencePng As String, referenceHPng As String
    Dim referenceVPng As String, referenceHVPng As String
    Dim candidatePng As String
    referencePng = NewTempPngPath("reference")
    referenceHPng = NewTempPngPath("reference_h")
    referenceVPng = NewTempPngPath("reference_v")
    referenceHVPng = NewTempPngPath("reference_hv")
    candidatePng = NewTempPngPath("candidate")
    If Not ExportUncroppedPreview(referenceShp, referencePng) Then GoTo Fail
    If Not ExportUncroppedPreview(referenceShp, referenceHPng, True, False) Then GoTo Fail
    If Not ExportUncroppedPreview(referenceShp, referenceVPng, False, True) Then GoTo Fail
    If Not ExportUncroppedPreview(referenceShp, referenceHVPng, True, True) Then GoTo Fail

    Dim referencePixels(0 To 3) As Variant
    referencePixels(0) = LoadCanonicalPreviewBytes(referencePng)
    referencePixels(1) = LoadCanonicalPreviewBytes(referenceHPng)
    referencePixels(2) = LoadCanonicalPreviewBytes(referenceVPng)
    referencePixels(3) = LoadCanonicalPreviewBytes(referenceHVPng)

    Dim matches As Collection
    Set matches = New Collection

    Dim sld As Slide, shp As Shape
    Dim shapeIndex As Long, originalShapeCount As Long
    For Each sld In Application.ActivePresentation.Slides
        originalShapeCount = sld.Shapes.Count
        For shapeIndex = 1 To originalShapeCount
            Set shp = sld.Shapes(shapeIndex)
            If shp.Type = msoGroup Then
                skipped = skipped + 1
            ElseIf IsComparablePicture(shp) Then
                If ExportUncroppedPreview(shp, candidatePng) Then
                    Dim isMatch As Boolean
                    isMatch = BinaryFilesEqual(referencePng, candidatePng) Or _
                              BinaryFilesEqual(referenceHPng, candidatePng) Or _
                              BinaryFilesEqual(referenceVPng, candidatePng) Or _
                              BinaryFilesEqual(referenceHVPng, candidatePng)

                    If Not isMatch Then
                        Dim candidatePixels As Variant
                        candidatePixels = LoadCanonicalPreviewBytes(candidatePng)
                        If Not IsEmpty(candidatePixels) Then
                            isMatch = PreviewPixelsNearlyEqual(referencePixels(0), candidatePixels) Or _
                                      PreviewPixelsNearlyEqual(referencePixels(1), candidatePixels) Or _
                                      PreviewPixelsNearlyEqual(referencePixels(2), candidatePixels) Or _
                                      PreviewPixelsNearlyEqual(referencePixels(3), candidatePixels)
                        End If
                    End If

                    If isMatch Then
                        matches.Add shp
                    End If
                End If
            End If
            SafeDeleteFile candidatePng
        Next shapeIndex
    Next sld

    SafeDeleteFile referencePng
    SafeDeleteFile referenceHPng
    SafeDeleteFile referenceVPng
    SafeDeleteFile referenceHVPng
    SafeDeleteFile candidatePng
    matched = matches.Count

    Dim item As Variant, target As Shape
    For Each item In matches
        Set target = item
        If ReplacePictureKeepCrop(target, imgPath) Then
            ReplaceAllMatchingPicturesCore = ReplaceAllMatchingPicturesCore + 1
        Else
            failed = failed + 1
        End If
    Next item
    Exit Function

Fail:
    SafeDeleteFile referencePng
    SafeDeleteFile referenceHPng
    SafeDeleteFile referenceVPng
    SafeDeleteFile referenceHVPng
    SafeDeleteFile candidatePng
End Function

' 使用 Windows 自带 WIA 将预览统一缩小为 32x32、32 位 BMP。
' 返回 BMP 的字节数组；WIA 不可用时返回 Empty，调用者安全退回严格比较。
Private Function LoadCanonicalPreviewBytes(ByVal pngPath As String) As Variant
    On Error GoTo Fail
    Dim imageFile As Object, imageProcess As Object, canonicalImage As Object
    Set imageFile = CreateObject("WIA.ImageFile")
    imageFile.LoadFile pngPath

    Set imageProcess = CreateObject("WIA.ImageProcess")
    imageProcess.Filters.Add imageProcess.FilterInfos("Scale").FilterID
    With imageProcess.Filters(1).Properties
        .Item("MaximumWidth").Value = PREVIEW_GRID_SIZE
        .Item("MaximumHeight").Value = PREVIEW_GRID_SIZE
        .Item("PreserveAspectRatio").Value = False
    End With
    imageProcess.Filters.Add imageProcess.FilterInfos("Convert").FilterID
    imageProcess.Filters(2).Properties("FormatID").Value = WIA_FORMAT_BMP

    Set canonicalImage = imageProcess.Apply(imageFile)
    LoadCanonicalPreviewBytes = canonicalImage.FileData.BinaryData
    Exit Function

Fail:
    LoadCanonicalPreviewBytes = Empty
End Function

Private Function PreviewPixelsNearlyEqual(ByRef bytesA As Variant, _
                                          ByRef bytesB As Variant) As Boolean
    On Error GoTo Different
    If Not IsArray(bytesA) Or Not IsArray(bytesB) Then Exit Function
    Dim baseA As Long, baseB As Long
    baseA = LBound(bytesA)
    baseB = LBound(bytesB)
    If UBound(bytesA) - baseA + 1 < 54 Then Exit Function
    If UBound(bytesB) - baseB + 1 < 54 Then Exit Function

    Dim widthA As Long, heightA As Long, widthB As Long, heightB As Long
    Dim bitsA As Long, bitsB As Long, offsetA As Long, offsetB As Long
    widthA = ReadLittleEndianLong(bytesA, baseA + 18)
    heightA = Abs(ReadLittleEndianLong(bytesA, baseA + 22))
    widthB = ReadLittleEndianLong(bytesB, baseB + 18)
    heightB = Abs(ReadLittleEndianLong(bytesB, baseB + 22))
    bitsA = ReadLittleEndianWord(bytesA, baseA + 28)
    bitsB = ReadLittleEndianWord(bytesB, baseB + 28)
    offsetA = baseA + ReadLittleEndianLong(bytesA, baseA + 10)
    offsetB = baseB + ReadLittleEndianLong(bytesB, baseB + 10)

    If widthA <> PREVIEW_GRID_SIZE Or heightA <> PREVIEW_GRID_SIZE Then Exit Function
    If widthB <> widthA Or heightB <> heightA Then Exit Function
    If bitsA <> 32 Or bitsB <> 32 Then Exit Function

    Dim pixelCount As Long, pixelIndex As Long, channel As Long
    Dim posA As Long, posB As Long, pixelDifference As Long
    Dim totalDifference As Double, maxPixelDifference As Long
    pixelCount = widthA * heightA

    For pixelIndex = 0 To pixelCount - 1
        posA = offsetA + pixelIndex * 4
        posB = offsetB + pixelIndex * 4
        pixelDifference = 0
        For channel = 0 To 2  ' BMP 顺序为 B、G、R；忽略 Alpha
            pixelDifference = pixelDifference + _
                              Abs(CLng(bytesA(posA + channel)) - CLng(bytesB(posB + channel)))
        Next channel
        If pixelDifference > MAX_PIXEL_RGB_SUM Then Exit Function
        If pixelDifference > maxPixelDifference Then maxPixelDifference = pixelDifference
        totalDifference = totalDifference + pixelDifference
    Next pixelIndex

    PreviewPixelsNearlyEqual = _
        (totalDifference / (pixelCount * 3#) <= MAX_MEAN_CHANNEL_DIFF) And _
        (maxPixelDifference <= MAX_PIXEL_RGB_SUM)
    Exit Function

Different:
    PreviewPixelsNearlyEqual = False
End Function

Private Function ReadLittleEndianWord(ByRef bytes As Variant, ByVal position As Long) As Long
    ReadLittleEndianWord = CLng(bytes(position)) + CLng(bytes(position + 1)) * 256&
End Function

Private Function ReadLittleEndianLong(ByRef bytes As Variant, ByVal position As Long) As Long
    ReadLittleEndianLong = CLng(bytes(position)) + _
                           CLng(bytes(position + 1)) * 256# + _
                           CLng(bytes(position + 2)) * 65536# + _
                           CLng(bytes(position + 3)) * 16777216#
End Function

' 导出统一尺寸的完整未裁剪预览；只用于识别，不改变原形状。
Private Function ExportUncroppedPreview(ByVal sourceShp As Shape, _
                                        ByVal filePath As String, _
                                        Optional ByVal extraFlipH As Boolean = False, _
                                        Optional ByVal extraFlipV As Boolean = False) As Boolean
    On Error GoTo Fail
    Dim probe As Shape
    Dim duplicateRange As ShapeRange
    Set duplicateRange = sourceShp.Duplicate
    Set probe = duplicateRange(1)

    ' 先把裁剪框扩展为完整图片尺寸，再最后归零中心偏移。
    ' 必须先改 ShapeWidth/Height、后改 Offset；反过来会因尺寸变化再次偏移。
    Dim fullPictureWidth As Single, fullPictureHeight As Single
    With probe.PictureFormat.Crop
        fullPictureWidth = .PictureWidth
        fullPictureHeight = .PictureHeight
        If fullPictureWidth <= 0 Or fullPictureHeight <= 0 Then GoTo Fail
        .ShapeWidth = fullPictureWidth
        .ShapeHeight = fullPictureHeight
        .PictureOffsetX = 0
        .PictureOffsetY = 0
    End With

    ' 不依赖实例当前的翻转状态；调用者可生成四个方向的基准指纹。
    If extraFlipH Then probe.Flip msoFlipHorizontal
    If extraFlipV Then probe.Flip msoFlipVertical
    probe.Rotation = 0
    probe.LockAspectRatio = msoFalse
    probe.Width = 128
    probe.Height = 128

    ' 去除实例级外观，避免相同原图因边框/阴影/调色不同而漏匹配。
    On Error Resume Next
    probe.Line.Visible = msoFalse
    probe.Shadow.Visible = msoFalse
    probe.Glow.Radius = 0
    probe.SoftEdge.Radius = 0
    probe.Reflection.Type = msoReflectionTypeNone
    On Error GoTo Fail

    SafeDeleteFile filePath
    ' Shape.Export 的缩放宽高是相对整张幻灯片，而不是最终像素尺寸。
    ' 显式按幻灯片/形状比例计算可消除默认导出中的 1px 取整漂移。
    Dim exportScaleWidth As Long, exportScaleHeight As Long
    exportScaleWidth = CLng(Application.ActivePresentation.PageSetup.SlideWidth * _
                            256# / probe.Width)
    exportScaleHeight = CLng(Application.ActivePresentation.PageSetup.SlideHeight * _
                             256# / probe.Height)
    probe.Export filePath, ppShapeFormatPNG, exportScaleWidth, exportScaleHeight, _
                 ppRelativeToSlide
    probe.Delete
    Set probe = Nothing
    ExportUncroppedPreview = (Len(Dir$(filePath)) > 0)
    Exit Function

Fail:
    On Error Resume Next
    If Not probe Is Nothing Then probe.Delete
    SafeDeleteFile filePath
    ExportUncroppedPreview = False
End Function

Private Function BinaryFilesEqual(ByVal pathA As String, ByVal pathB As String) As Boolean
    On Error GoTo Different
    If Len(Dir$(pathA)) = 0 Or Len(Dir$(pathB)) = 0 Then Exit Function
    If FileLen(pathA) <> FileLen(pathB) Then Exit Function

    Dim fileA As Integer, fileB As Integer
    Dim dataA As String, dataB As String
    fileA = FreeFile
    Open pathA For Binary Access Read As #fileA
    fileB = FreeFile
    Open pathB For Binary Access Read As #fileB
    dataA = Space$(LOF(fileA))
    dataB = Space$(LOF(fileB))
    If Len(dataA) > 0 Then Get #fileA, , dataA
    If Len(dataB) > 0 Then Get #fileB, , dataB
    Close #fileA: fileA = 0
    Close #fileB: fileB = 0
    BinaryFilesEqual = (dataA = dataB)
    Exit Function

Different:
    On Error Resume Next
    If fileA <> 0 Then Close #fileA
    If fileB <> 0 Then Close #fileB
End Function

Private Function NewTempPngPath(ByVal label As String) As String
    Randomize
    NewTempPngPath = Environ$("TEMP") & "\PicReplace_" & label & "_" & _
                     CStr(CLng(Timer * 1000)) & "_" & CStr(Int(Rnd() * 1000000)) & ".png"
End Function

Private Sub SafeDeleteFile(ByVal filePath As String)
    On Error Resume Next
    If Len(filePath) > 0 Then
        If Len(Dir$(filePath)) > 0 Then Kill filePath
    End If
    On Error GoTo 0
End Sub

Private Function IsComparablePicture(ByVal shp As Shape) As Boolean
    On Error GoTo Nope
    Select Case shp.Type
        Case msoPicture, msoLinkedPicture
            IsComparablePicture = True
        Case msoPlaceholder
            IsComparablePicture = (shp.PlaceholderFormat.ContainedType = msoPicture)
    End Select
Nope:
End Function

'---------- 核心：原位替换并保留裁剪 ----------
Public Function ReplacePictureKeepCrop(ByVal oldShp As Shape, ByVal imgPath As String) As Boolean
    On Error GoTo Fail

    ' 1. 记录原图几何、Crop 对象与标识信息
    Dim L As Single, T As Single, rot As Single
    Dim flipH As Long, flipV As Long, lockAR As Long
    Dim oldName As String, altText As String
    Dim zPos As Long
    Dim cropShapeW As Single, cropShapeH As Single
    Dim cropPicW As Single, cropPicH As Single
    Dim cropOffX As Single, cropOffY As Single
    Dim ins As Shape

    L = oldShp.Left: T = oldShp.Top
    rot = oldShp.Rotation
    flipH = oldShp.HorizontalFlip
    flipV = oldShp.VerticalFlip
    lockAR = oldShp.LockAspectRatio
    oldName = oldShp.Name
    zPos = oldShp.ZOrderPosition
    On Error Resume Next
    altText = oldShp.AlternativeText
    On Error GoTo Fail

    With oldShp.PictureFormat.Crop
        cropShapeW = .ShapeWidth: cropShapeH = .ShapeHeight
        cropPicW = .PictureWidth: cropPicH = .PictureHeight
        cropOffX = .PictureOffsetX: cropOffY = .PictureOffsetY
    End With
    If cropShapeW <= 0 Or cropShapeH <= 0 Or cropPicW <= 0 Or cropPicH <= 0 Then GoTo Fail

    ' 2. 插入真正的新图片。Fill.UserPicture 只改变 Picture 形状的填充，
    '    导出时底层旧 blip 仍可能可见，因此不能作为图片替换手段。
    Dim sld As Object
    Set sld = oldShp.Parent
    Set ins = sld.Shapes.AddPicture(FileName:=imgPath, LinkToFile:=msoFalse, _
                                    SaveWithDocument:=msoTrue, Left:=L, Top:=T)
    Dim nW As Single, nH As Single
    nW = ins.Width: nH = ins.Height
    If nW <= 0 Or nH <= 0 Then GoTo Fail

    ' 3. 相近宽高比：精确保留原 Crop 几何。
    '    差异较大：保持新图比例，用 cover 铺满原框并保留原焦点偏移。
    If Not ApplyPreservedCrop(ins, cropShapeW, cropShapeH, cropPicW, cropPicH, _
                              cropOffX, cropOffY, nW, nH) Then GoTo Fail

    If flipH = msoTrue Then ins.Flip msoFlipHorizontal
    If flipV = msoTrue Then ins.Flip msoFlipVertical
    ins.Rotation = rot
    ins.Left = L
    ins.Top = T
    ins.LockAspectRatio = lockAR

    CopyCosmetics oldShp, ins

    ' 移动到原图正上方，随后删除原图，层级自然落回原位
    Dim guard As Long
    Do While ins.ZOrderPosition > zPos + 1 And guard < 1000
        ins.ZOrder msoSendBackward
        guard = guard + 1
    Loop

    FinalizeSwap oldShp, ins, oldName, altText
    ReplacePictureKeepCrop = True
    Exit Function

Fail:
    ' 事务式失败：清理新建对象，原图片必须留在原位。
    On Error Resume Next
    If Not ins Is Nothing Then ins.Delete
    On Error GoTo 0
    ReplacePictureKeepCrop = False
End Function

'---------- 将原裁剪几何映射到新图片 ----------
Private Function ApplyPreservedCrop(ByVal shp As Shape, _
                                    ByVal shapeW As Single, ByVal shapeH As Single, _
                                    ByVal oldPicW As Single, ByVal oldPicH As Single, _
                                    ByVal oldOffX As Single, ByVal oldOffY As Single, _
                                    ByVal naturalW As Single, ByVal naturalH As Single) As Boolean
    On Error GoTo Bad

    Dim oldAspect As Double, newAspect As Double, aspectChange As Double
    oldAspect = oldPicW / oldPicH
    newAspect = naturalW / naturalH
    aspectChange = newAspect / oldAspect

    Dim newPicW As Single, newPicH As Single
    Dim newOffX As Single, newOffY As Single

    If aspectChange >= 0.8 And aspectChange <= 1.25 Then
        ' 同尺寸或类似尺寸：像素映射和焦点位置都与原图一致。
        newPicW = oldPicW: newPicH = oldPicH
        newOffX = oldOffX: newOffY = oldOffY
    Else
        ' 宽高比明显不同：保持新图比例，以 cover 方式铺满原裁剪框。
        Dim frameAspect As Double
        frameAspect = shapeW / shapeH
        If newAspect >= frameAspect Then
            newPicH = shapeH
            newPicW = newPicH * newAspect
        Else
            newPicW = shapeW
            newPicH = newPicW / newAspect
        End If

        ' 保留原来的放大程度，避免宽高比变化时突然缩回“适应框”。
        Dim oldBaseW As Double, oldBaseH As Double, zoom As Double
        If oldAspect >= frameAspect Then
            oldBaseH = shapeH: oldBaseW = oldBaseH * oldAspect
        Else
            oldBaseW = shapeW: oldBaseH = oldBaseW / oldAspect
        End If
        zoom = oldPicW / oldBaseW
        If oldPicH / oldBaseH > zoom Then zoom = oldPicH / oldBaseH
        If zoom < 1 Then zoom = 1
        If zoom > 100 Then zoom = 100
        newPicW = newPicW * zoom
        newPicH = newPicH * zoom

        ' 偏移按图片尺寸归一化，等价于保留原焦点位置。
        newOffX = (oldOffX / oldPicW) * newPicW
        newOffY = (oldOffY / oldPicH) * newPicH
        newOffX = ClampValue(newOffX, -(newPicW - shapeW) / 2, (newPicW - shapeW) / 2)
        newOffY = ClampValue(newOffY, -(newPicH - shapeH) / 2, (newPicH - shapeH) / 2)
    End If

    shp.LockAspectRatio = msoFalse
    With shp.PictureFormat.Crop
        .ShapeWidth = shapeW
        .ShapeHeight = shapeH
        .PictureWidth = newPicW
        .PictureHeight = newPicH
        .PictureOffsetX = newOffX
        .PictureOffsetY = newOffY
    End With
    ApplyPreservedCrop = True
    Exit Function
Bad:
    ApplyPreservedCrop = False
End Function

Private Function ClampValue(ByVal value As Single, ByVal minimum As Single, _
                            ByVal maximum As Single) As Single
    If maximum < minimum Then
        ClampValue = 0
    ElseIf value < minimum Then
        ClampValue = minimum
    ElseIf value > maximum Then
        ClampValue = maximum
    Else
        ClampValue = value
    End If
End Function

'---------- 收尾：删除原图、恢复名称与选中状态 ----------
Private Sub FinalizeSwap(ByVal oldShp As Shape, ByVal newShp As Shape, _
                         ByVal oldName As String, ByVal altText As String)
    oldShp.Delete
    On Error Resume Next
    newShp.Name = oldName
    newShp.AlternativeText = altText
    newShp.Select
    On Error GoTo 0
End Sub

'---------- 尽力复制外观效果 ----------
Private Sub CopyCosmetics(ByVal src As Shape, ByVal dst As Shape)
    On Error Resume Next
    ' 边框线条
    If src.Line.Visible = msoTrue Then
        dst.Line.Visible = msoTrue
        dst.Line.ForeColor.RGB = src.Line.ForeColor.RGB
        dst.Line.Weight = src.Line.Weight
        dst.Line.DashStyle = src.Line.DashStyle
        dst.Line.Style = src.Line.Style
        dst.Line.Transparency = src.Line.Transparency
    End If
    ' 阴影
    If src.Shadow.Visible = msoTrue Then
        dst.Shadow.Visible = msoTrue
        dst.Shadow.ForeColor.RGB = src.Shadow.ForeColor.RGB
        dst.Shadow.Transparency = src.Shadow.Transparency
        dst.Shadow.Blur = src.Shadow.Blur
        dst.Shadow.OffsetX = src.Shadow.OffsetX
        dst.Shadow.OffsetY = src.Shadow.OffsetY
        dst.Shadow.RotateWithShape = src.Shadow.RotateWithShape
    End If
    ' 发光
    If src.Glow.Radius > 0 Then
        dst.Glow.Radius = src.Glow.Radius
        dst.Glow.Color.RGB = src.Glow.Color.RGB
        dst.Glow.Transparency = src.Glow.Transparency
    End If
    ' 柔化边缘
    dst.SoftEdge.Radius = src.SoftEdge.Radius
    ' 映像
    If src.Reflection.Type <> msoReflectionTypeNone Then
        dst.Reflection.Type = src.Reflection.Type
        dst.Reflection.Transparency = src.Reflection.Transparency
        dst.Reflection.Size = src.Reflection.Size
        dst.Reflection.Blur = src.Reflection.Blur
        dst.Reflection.Offset = src.Reflection.Offset
    End If
    ' 图片颜色调整
    dst.PictureFormat.Brightness = src.PictureFormat.Brightness
    dst.PictureFormat.Contrast = src.PictureFormat.Contrast
    dst.PictureFormat.ColorType = src.PictureFormat.ColorType
    If src.PictureFormat.TransparentBackground = msoTrue Then
        dst.PictureFormat.TransparencyColor = src.PictureFormat.TransparencyColor
        dst.PictureFormat.TransparentBackground = msoTrue
    End If
    On Error GoTo 0
End Sub

'---------- 获取并校验当前选中的图片 ----------
Private Function TryGetSelectedPicture() As Shape
    On Error GoTo BadState
    Dim sel As Selection
    Set sel = Application.ActiveWindow.Selection
    If sel.Type <> ppSelectionShapes Then
        Notify "请先在幻灯片中选中一张图片。"
        Exit Function
    End If
    If sel.ShapeRange.Count <> 1 Then
        Notify "一次只能替换一张图片，请只选中一张。"
        Exit Function
    End If

    Dim shp As Shape
    Set shp = sel.ShapeRange(1)

    Dim inGroup As Boolean
    Dim pg As Shape
    On Error Resume Next
    Set pg = shp.ParentGroup
    inGroup = Not pg Is Nothing
    On Error GoTo BadState
    If inGroup Then
        Notify "所选图片在组合内。请先取消组合（Ctrl+Shift+G），替换后再重新组合。"
        Exit Function
    End If

    Select Case shp.Type
        Case msoPicture, msoLinkedPicture
            ' 支持
        Case msoPlaceholder
            Dim ct As Long
            On Error Resume Next
            ct = shp.PlaceholderFormat.ContainedType
            On Error GoTo BadState
            If ct <> msoPicture Then
                Notify "该占位符中的内容不是图片。"
                Exit Function
            End If
        Case Else
            Notify "所选对象不是图片（支持普通图片、链接图片和图片占位符）。"
            Exit Function
    End Select

    Set TryGetSelectedPicture = shp
    Exit Function
BadState:
    Notify "无法读取当前选择，请切换到普通视图后重试。"
End Function

'---------- 文件选择对话框 ----------
Private Function PickImageFile() As String
    On Error GoTo NoDlg
    Dim fd As Object
    Set fd = Application.FileDialog(3)   ' msoFileDialogFilePicker
    With fd
        .Title = "选择替换图片"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "图片文件", "*.png;*.jpg;*.jpeg;*.jfif;*.bmp;*.gif;*.tif;*.tiff;*.webp;*.emf;*.wmf"
        .Filters.Add "所有文件", "*.*"
        If .Show = -1 Then
            PickImageFile = .SelectedItems(1)
        End If
    End With
    Exit Function
NoDlg:
    PickImageFile = ""
End Function

Private Sub Notify(ByVal msg As String)
    MsgBox msg, vbExclamation, "原位替换图片"
End Sub

'---------- 自检：供加载测试调用 ----------
Public Sub PictureReplace_Ping()
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open Environ$("TEMP") & "\picreplace_addin_ping.txt" For Output As #f
    Print #f, "ok"
    Close #f
End Sub
