Option Explicit

Const RED = "C:\Users\Administrator\Documents\powerpoint\pic_replace_addin\test_assets\old_red.png"
Const BLUE = "C:\Users\Administrator\Documents\powerpoint\pic_replace_addin\test_assets\new_blue.png"
Const GREEN = "C:\Users\Administrator\Documents\powerpoint\pic_replace_addin\test_assets\new_green_wide.png"

Dim pp: Set pp = CreateObject("PowerPoint.Application")
pp.Visible = -1
Dim pres: Set pres = pp.Presentations.Add()
Dim sld: Set sld = pres.Slides.Add(1, 12)
Dim shp: Set shp = sld.Shapes.AddPicture(RED, 0, -1, 100, 100)
shp.LockAspectRatio = 0
shp.Width = 240
shp.Height = 150
shp.PictureFormat.CropLeft = 25
shp.PictureFormat.CropRight = 35
shp.PictureFormat.CropTop = 15
shp.PictureFormat.CropBottom = 45

Dump "OLD", shp

Dim same: Set same = shp.Duplicate()(1)
same.Fill.UserPicture BLUE
Dump "SAME AFTER FILL", same
RestoreCropGeometry shp, same
Dump "SAME RESTORED", same

Dim wide: Set wide = shp.Duplicate()(1)
wide.Fill.UserPicture GREEN
Dump "WIDE AFTER FILL", wide
RestoreCropGeometry shp, wide
Dump "WIDE RESTORED EXACT", wide

pres.Close
pp.Quit

Sub RestoreCropGeometry(src, dst)
    With dst.PictureFormat.Crop
        .ShapeWidth = src.PictureFormat.Crop.ShapeWidth
        .ShapeHeight = src.PictureFormat.Crop.ShapeHeight
        .PictureWidth = src.PictureFormat.Crop.PictureWidth
        .PictureHeight = src.PictureFormat.Crop.PictureHeight
        .PictureOffsetX = src.PictureFormat.Crop.PictureOffsetX
        .PictureOffsetY = src.PictureFormat.Crop.PictureOffsetY
    End With
    dst.Left = src.Left
    dst.Top = src.Top
End Sub

Sub Dump(label, s)
    WScript.Echo "--- " & label & " ---"
    WScript.Echo "shape L=" & s.Left & " T=" & s.Top & " W=" & s.Width & " H=" & s.Height
    WScript.Echo "sides L=" & s.PictureFormat.CropLeft & " R=" & s.PictureFormat.CropRight & " T=" & s.PictureFormat.CropTop & " B=" & s.PictureFormat.CropBottom
    With s.PictureFormat.Crop
        WScript.Echo "crop shape W=" & .ShapeWidth & " H=" & .ShapeHeight & " L=" & .ShapeLeft & " T=" & .ShapeTop
        WScript.Echo "crop pic W=" & .PictureWidth & " H=" & .PictureHeight & " offX=" & .PictureOffsetX & " offY=" & .PictureOffsetY
    End With
End Sub
