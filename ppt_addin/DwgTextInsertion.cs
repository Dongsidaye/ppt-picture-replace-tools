using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Dongsidaye.PowerPointTools
{
    internal sealed class DwgTextItem
    {
        public string SourceId { get; set; }
        public string Text { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double CenterX { get; set; }
        public double CenterY { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double InkWidth { get; set; }
        public double InkHeight { get; set; }
        public double FontSize { get; set; }
        public string FontFamily { get; set; }
        public double Rotation { get; set; }
        public string Color { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
    }

    internal static class DwgTextInsertion
    {
        private const double MinimumFontSize = 1.0;
        private const double MaximumFontSize = 4000.0;

        internal static PowerPoint.Shape Add(PowerPoint.Slide slide, DwgTextItem item,
            float svgLeft, float svgTop, double viewBoxLeft, double viewBoxTop,
            double scale, string selectedFont)
        {
            if (slide == null) throw new ArgumentNullException("slide");
            if (item == null) throw new ArgumentNullException("item");
            if (String.IsNullOrWhiteSpace(item.Text)) return null;
            if (item.Text.IndexOfAny(new[] { '\r', '\n' }) >= 0)
                throw new InvalidOperationException("DWG 多行文字需要先按原行位置拆分。");
            if (item.Text.Length > 32767)
                throw new InvalidOperationException("单行 DWG 文字过长，请拆分后导入。");

            RequireFinite(scale, "文字缩放比例");
            RequireFinite(item.FontSize, "文字字号");
            RequireFinite(item.CenterX, "文字位置");
            RequireFinite(item.CenterY, "文字位置");
            RequireFinite(item.Rotation, "文字角度");
            RequireFinite(viewBoxLeft, "图纸范围");
            RequireFinite(viewBoxTop, "图纸范围");
            if (scale <= 0 || item.FontSize <= 0)
                throw new InvalidOperationException("DWG 文字尺寸无效。");

            double pointSize = item.FontSize * scale;
            // Rounding at the Office limit should not reject a nominal 1pt line.
            if (pointSize >= 0.9999 && pointSize < MinimumFontSize) pointSize = MinimumFontSize;
            if (!IsFinite(pointSize) || pointSize < MinimumFontSize || pointSize > MaximumFontSize)
                throw new InvalidOperationException("DWG 文字缩放后超出 PowerPoint 的 1–4000 磅字号范围，请放大图纸或保留 SVG 文字。");

            double targetWidth = PositiveDimension(item.InkWidth, item.Width) * scale;
            double targetHeight = PositiveDimension(item.InkHeight, item.Height) * scale;
            if (!IsFinite(targetWidth) || !IsFinite(targetHeight) || targetWidth <= 0 || targetHeight <= 0)
                throw new InvalidOperationException("DWG 文字边界无效。");
            double centerX = svgLeft + (item.CenterX - viewBoxLeft) * scale;
            double centerY = svgTop + (item.CenterY - viewBoxTop) * scale;
            RequireSingle(centerX, "文字位置");
            RequireSingle(centerY, "文字位置");

            int first = 0;
            int last = item.Text.Length - 1;
            while (first <= last && Char.IsWhiteSpace(item.Text[first])) first++;
            while (last >= first && Char.IsWhiteSpace(item.Text[last])) last--;
            string paintedText = item.Text.Substring(first, last - first + 1);
            string fontName = String.IsNullOrWhiteSpace(selectedFont) ? "Microsoft YaHei" : selectedFont.Trim();
            FontStyle style = (item.Bold ? FontStyle.Bold : FontStyle.Regular) | (item.Italic ? FontStyle.Italic : FontStyle.Regular);
            InkMetrics sourceMetrics = MeasureInk(paintedText, fontName, style, pointSize);
            double fittedSize = pointSize * targetHeight / sourceMetrics.Bounds.Height;
            if (!IsFinite(fittedSize) || fittedSize < MinimumFontSize || fittedSize > MaximumFontSize)
                throw new InvalidOperationException("所选字体匹配原文字高度后超出 PowerPoint 的字号范围，需保留 SVG 文字。");
            InkMetrics fittedMetrics = MeasureInk(paintedText, fontName, style, fittedSize);
            double horizontalScale = targetWidth / fittedMetrics.Bounds.Width;
            bool useWordArt = Math.Abs(horizontalScale - 1.0) > 0.03;

            PowerPoint.Shape shape = null;
            try
            {
                float initialWidth = ToSingle(Math.Max(targetWidth * 1.2, pointSize * 2), "文字宽度");
                float initialHeight = ToSingle(Math.Max(targetHeight * 1.5, pointSize * 2), "文字高度");
                if (useWordArt)
                    shape = slide.Shapes.AddTextEffect(Office.MsoPresetTextEffect.msoTextEffect1, item.Text,
                        fontName, (float)fittedSize, item.Bold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse,
                        item.Italic ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse, 0, 0);
                else
                    shape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        0, 0, initialWidth, initialHeight);
                PowerPoint.TextFrame2 frame = shape.TextFrame2;
                frame.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;
                frame.WordWrap = Office.MsoTriState.msoFalse;
                frame.MarginLeft = 0;
                frame.MarginRight = 0;
                frame.MarginTop = 0;
                frame.MarginBottom = 0;
                frame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;
                if (useWordArt) ConfigurePlainWordArt(shape, frame);

                Office.TextRange2 range = frame.TextRange;
                range.Text = item.Text;
                range.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignLeft;
                range.ParagraphFormat.SpaceBefore = 0;
                range.ParagraphFormat.SpaceAfter = 0;
                range.ParagraphFormat.LineRuleWithin = Office.MsoTriState.msoTrue;
                range.ParagraphFormat.SpaceWithin = 1;

                Office.Font2 font = range.Font;
                font.Name = fontName;
                font.NameAscii = fontName;
                font.NameFarEast = fontName;
                font.NameComplexScript = fontName;
                font.Size = (float)pointSize;
                font.Bold = item.Bold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                font.Italic = item.Italic ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                font.Spacing = 0;
                font.Fill.Visible = Office.MsoTriState.msoTrue;
                font.Fill.Solid();
                font.Fill.ForeColor.RGB = ParseColor(item.Color);
                font.Line.Visible = Office.MsoTriState.msoFalse;
                font.Shadow.Visible = Office.MsoTriState.msoFalse;
                font.Glow.Radius = 0;
                font.Reflection.Type = Office.MsoReflectionType.msoReflectionTypeNone;
                font.SoftEdgeFormat = Office.MsoSoftEdgeType.msoSoftEdgeTypeNone;

                Office.TextRange2 paintedRange = range.Characters[first + 1, last - first + 1];
                InkMetrics metrics = FitHeight(range, paintedText, fontName, style, targetHeight, pointSize);
                if (useWordArt)
                    FitWordArtAndPlace(shape, targetWidth, targetHeight, centerX, centerY, item.Rotation);
                else
                {
                    double nativeWidthAtZeroSpacing = paintedRange.BoundWidth;
                    FitWidth(range, paintedRange, metrics.Bounds.Width, nativeWidthAtZeroSpacing, targetWidth, paintedText);
                    FitFrameAndPlace(shape, range, paintedRange, metrics, nativeWidthAtZeroSpacing,
                        targetWidth, targetHeight, centerX, centerY, item.Rotation);
                }
                // Resizing a text warp can restore the preset's theme outline.
                // Apply an explicit transparent line after all sizing changes.
                font.Line.ForeColor.RGB = ParseColor(item.Color);
                font.Line.Transparency = 1;
                font.Line.Visible = Office.MsoTriState.msoFalse;
                return shape;
            }
            catch (Exception error)
            {
                if (shape != null)
                {
                    try { shape.Delete(); }
                    catch (Exception cleanupError)
                    {
                        // The importer can include this owned shape in its batch
                        // rollback if Office refused the local cleanup attempt.
                        error.Data["DwgCreatedShape"] = shape;
                        error.Data["DwgTextCleanupError"] = cleanupError.Message;
                    }
                }
                throw;
            }
        }

        private static void ConfigurePlainWordArt(PowerPoint.Shape shape, PowerPoint.TextFrame2 frame)
        {
            // PlainText is the affine text warp. Merely resizing an ordinary
            // text frame leaves the glyph width unchanged in PowerPoint.
            shape.TextEffect.PresetShape = Office.MsoPresetTextEffectShape.msoTextEffectShapePlainText;
            shape.TextEffect.NormalizedHeight = Office.MsoTriState.msoFalse;
            shape.Fill.Visible = Office.MsoTriState.msoFalse;
            shape.Line.Visible = Office.MsoTriState.msoFalse;
            shape.Shadow.Visible = Office.MsoTriState.msoFalse;
            shape.Glow.Radius = 0;
            shape.Reflection.Type = Office.MsoReflectionType.msoReflectionTypeNone;
            shape.SoftEdge.Radius = 0;
            ClearTextDepth(shape.ThreeD);
            ClearTextDepth(frame.ThreeD);
        }

        private static void ClearTextDepth(PowerPoint.ThreeDFormat depth)
        {
            depth.Visible = Office.MsoTriState.msoFalse;
            depth.Depth = 0;
            depth.BevelTopType = Office.MsoBevelType.msoBevelNone;
            depth.BevelBottomType = Office.MsoBevelType.msoBevelNone;
            depth.BevelTopDepth = 0;
            depth.BevelBottomDepth = 0;
            depth.BevelTopInset = 0;
            depth.BevelBottomInset = 0;
            depth.Visible = Office.MsoTriState.msoFalse;
        }

        private static void FitWordArtAndPlace(PowerPoint.Shape shape, double targetWidth, double targetHeight,
            double centerX, double centerY, double rotation)
        {
            shape.LockAspectRatio = Office.MsoTriState.msoFalse;
            shape.Width = ToSingle(targetWidth, "文字宽度");
            shape.Height = ToSingle(targetHeight, "文字高度");
            ValidateFittedBounds(shape.Width, shape.Height, targetWidth, targetHeight);
            shape.Left = ToSingle(centerX - targetWidth / 2.0, "文字位置");
            shape.Top = ToSingle(centerY - targetHeight / 2.0, "文字位置");
            shape.Rotation = (float)(((rotation % 360) + 360) % 360);
        }

        private sealed class InkMetrics
        {
            internal RectangleF Bounds;
            internal double LineHeight;
        }

        private static InkMetrics MeasureInk(string text, string fontName, FontStyle style, double pointSize)
        {
            using (var family = new FontFamily(fontName))
            using (var format = (StringFormat)StringFormat.GenericTypographic.Clone())
            using (var path = new GraphicsPath())
            {
                format.FormatFlags |= StringFormatFlags.NoWrap | StringFormatFlags.MeasureTrailingSpaces;
                // AddString's emSize is in path coordinates. Passing points here
                // gives real glyph outlines in the same unit used by Office.
                path.AddString(text, family, (int)style, (float)pointSize, PointF.Empty, format);
                RectangleF bounds = path.GetBounds();
                if (path.PointCount == 0 || bounds.Width <= 0 || bounds.Height <= 0)
                    throw new InvalidOperationException("所选字体无法生成 DWG 文字的墨迹边界。");
                double lineHeight = pointSize * family.GetLineSpacing(style) / family.GetEmHeight(style);
                return new InkMetrics { Bounds = bounds, LineHeight = lineHeight };
            }
        }

        private static InkMetrics FitHeight(Office.TextRange2 range, string text, string fontName,
            FontStyle style, double targetHeight, double originalSize)
        {
            InkMetrics initial = MeasureInk(text, fontName, style, originalSize);
            double size = originalSize * targetHeight / initial.Bounds.Height;
            if (!IsFinite(size) || size < MinimumFontSize || size > MaximumFontSize)
                throw new InvalidOperationException("所选字体匹配原文字高度后超出 PowerPoint 的字号范围，需保留 SVG 文字。");
            range.Font.Size = (float)size;
            return MeasureInk(text, fontName, style, range.Font.Size);
        }

        private static void FitWidth(Office.TextRange2 range, Office.TextRange2 paintedRange,
            double inkWidthAtZeroSpacing, double nativeWidthAtZeroSpacing, double targetWidth, string text)
        {
            int characters = StringInfo.ParseCombiningCharacters(text).Length;
            if (characters < 2) return;
            double pointSize = range.Font.Size;
            double minimum = -pointSize * 0.6;
            double maximum = pointSize * 2;
            for (int iteration = 0; iteration < 3; iteration++)
            {
                double measured = InkWidth(paintedRange, inkWidthAtZeroSpacing, nativeWidthAtZeroSpacing);
                if (!IsFinite(measured) || measured <= 0) throw new InvalidOperationException("PowerPoint 无法测量 DWG 文字宽度。");
                double difference = targetWidth - measured;
                if (Math.Abs(difference) <= Math.Max(0.05, targetWidth * 0.005)) break;
                double spacing = range.Font.Spacing + difference / (characters - 1);
                double bounded = Math.Max(minimum, Math.Min(maximum, spacing));
                range.Font.Spacing = (float)bounded;
                if (Math.Abs(spacing - bounded) > 0.001) break;
            }
        }

        private static void FitFrameAndPlace(PowerPoint.Shape shape, Office.TextRange2 range, Office.TextRange2 paintedRange,
            InkMetrics metrics, double nativeWidthAtZeroSpacing, double targetWidth, double targetHeight,
            double centerX, double centerY, double rotation)
        {
            double measuredWidth = range.BoundWidth;
            double measuredHeight = range.BoundHeight;
            if (!IsFinite(measuredWidth) || !IsFinite(measuredHeight) || measuredWidth <= 0 || measuredHeight <= 0)
                throw new InvalidOperationException("PowerPoint 返回的文字边界无效。");
            shape.Width = ToSingle(Math.Max(targetWidth, measuredWidth) + 0.1, "文字宽度");
            shape.Height = ToSingle(Math.Max(targetHeight, measuredHeight) + 0.1, "文字高度");
            double inkWidth = InkWidth(paintedRange, metrics.Bounds.Width, nativeWidthAtZeroSpacing);
            ValidateFittedBounds(inkWidth, metrics.Bounds.Height, targetWidth, targetHeight);

            // BoundHeight is a line box, not painted glyph height. Center the
            // font's typographic line box in it and then apply glyph bearings.
            double inkLeft = paintedRange.BoundLeft + metrics.Bounds.Left;
            double inkTop = paintedRange.BoundTop + (paintedRange.BoundHeight - metrics.LineHeight) / 2.0 + metrics.Bounds.Top;
            double localCenterX = inkLeft + inkWidth / 2.0 - shape.Left - shape.Width / 2.0;
            double localCenterY = inkTop + metrics.Bounds.Height / 2.0 - shape.Top - shape.Height / 2.0;
            double normalized = ((rotation % 360) + 360) % 360;
            double radians = normalized * Math.PI / 180.0;
            double rotatedX = localCenterX * Math.Cos(radians) - localCenterY * Math.Sin(radians);
            double rotatedY = localCenterX * Math.Sin(radians) + localCenterY * Math.Cos(radians);
            shape.Rotation = (float)normalized;
            shape.Left = ToSingle(centerX - shape.Width / 2.0 - rotatedX, "文字位置");
            shape.Top = ToSingle(centerY - shape.Height / 2.0 - rotatedY, "文字位置");
        }

        private static double InkWidth(Office.TextRange2 paintedRange, double inkWidthAtZeroSpacing,
            double nativeWidthAtZeroSpacing)
        {
            // Keep the actual glyph width from GDI+, using Office only for the
            // width change caused by its character-spacing implementation.
            return inkWidthAtZeroSpacing + paintedRange.BoundWidth - nativeWidthAtZeroSpacing;
        }

        private static void ValidateFittedBounds(double width, double height, double targetWidth, double targetHeight)
        {
            if (!IsFinite(width) || !IsFinite(height) || width <= 0 || height <= 0)
                throw new InvalidOperationException("PowerPoint 返回的文字墨迹边界无效。");
            // A single glyph cannot be compressed using character spacing.
            // Reject large residuals for both single glyphs and clamped runs so
            // the caller can restore their original SVG text by SourceId.
            if (Math.Abs(width - targetWidth) > Math.Max(0.5, targetWidth * 0.05))
                throw new InvalidOperationException("替换字体后的文字宽度差异较大，需保留 SVG 文字。");
            if (Math.Abs(height - targetHeight) > Math.Max(0.5, targetHeight * 0.15))
                throw new InvalidOperationException("替换字体后的文字高度差异较大，需保留 SVG 文字。");
        }

        private static int ParseColor(string value)
        {
            if (String.IsNullOrWhiteSpace(value)) return 0;
            string text = value.Trim();
            uint rgb;
            if (text.Length != 7 || text[0] != '#' || !UInt32.TryParse(text.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out rgb))
                throw new InvalidOperationException("DWG 文字颜色无效。");
            return (int)(((rgb >> 16) & 255) | (rgb & 65280) | ((rgb & 255) << 16));
        }

        private static double PositiveDimension(double preferred, double fallback)
        {
            if (IsFinite(preferred) && preferred > 0) return preferred;
            return fallback;
        }

        private static bool IsFinite(double value) { return !Double.IsNaN(value) && !Double.IsInfinity(value); }

        private static void RequireFinite(double value, string label)
        {
            if (!IsFinite(value)) throw new InvalidOperationException(label + "无效。");
        }

        private static void RequireSingle(double value, string label)
        {
            if (!IsFinite(value) || Math.Abs(value) > Single.MaxValue) throw new InvalidOperationException(label + "超出 PowerPoint 范围。");
        }

        private static float ToSingle(double value, string label)
        {
            RequireSingle(value, label);
            return (float)value;
        }
    }
}
