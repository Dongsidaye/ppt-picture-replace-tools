using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using Dongsidaye.PowerPointTools;

namespace Dongsidaye.WpsCad
{
    internal sealed class InkMetrics
    {
        public double x { get; set; }
        public double y { get; set; }
        public double width { get; set; }
        public double height { get; set; }
        public double lineHeight { get; set; }
    }

    internal static class FontFitting
    {
        internal static TextPlan Plan(DwgTextItem item, bool ocr, PreparedState prepared)
        {
            var plan = new TextPlan { item = item, ocr = ocr, unsupported = "", metrics = null };
            try
            {
                if (item == null || String.IsNullOrWhiteSpace(item.Text) || item.Text.Length > 32767 || item.Text.IndexOfAny(new[] { '\r', '\n' }) >= 0)
                    throw new InvalidOperationException("文字必须是有效的单行内容。");
                RequireFinite(item.FontSize, item.CenterX, item.CenterY, item.Rotation, item.Width, item.Height, item.InkWidth, item.InkHeight);
                double pointSize = item.FontSize * prepared.scale;
                if (pointSize >= 0.9999 && pointSize < 1) pointSize = 1;
                if (!Program.Finite(pointSize) || pointSize < 1 || pointSize > 4000) throw new InvalidOperationException("缩放后文字超出 1–4000 磅范围。");
                plan.targetWidth = (item.InkWidth > 0 ? item.InkWidth : item.Width) * prepared.scale;
                plan.targetHeight = (item.InkHeight > 0 ? item.InkHeight : item.Height) * prepared.scale;
                plan.centerX = prepared.left + (item.CenterX - prepared.viewBox[0]) * prepared.scale;
                plan.centerY = prepared.top + (item.CenterY - prepared.viewBox[1]) * prepared.scale;
                RequireFinite(plan.targetWidth, plan.targetHeight, plan.centerX, plan.centerY);
                if (plan.targetWidth <= 0 || plan.targetHeight <= 0 || Math.Abs(plan.centerX) > 1e9 || Math.Abs(plan.centerY) > 1e9)
                    throw new InvalidOperationException("文字位置或边界无效。");
                string painted = item.Text.Trim();
                FontStyle style = (item.Bold ? FontStyle.Bold : FontStyle.Regular) | (item.Italic ? FontStyle.Italic : FontStyle.Regular);
                InkMetrics initial = Measure(painted, prepared.font, style, pointSize);
                double fitted = pointSize * plan.targetHeight / initial.height;
                if (!Program.Finite(fitted) || fitted < 1 || fitted > 4000) throw new InvalidOperationException("字体匹配原文字高度后超出字号范围。");
                plan.fontSize = Math.Round(fitted, 2, MidpointRounding.AwayFromZero);
                if (plan.fontSize < 1 || plan.fontSize > 4000) throw new InvalidOperationException("字体量化后超出字号范围。");
                plan.metrics = Measure(painted, prepared.font, style, plan.fontSize);
            }
            catch (Exception error)
            {
                plan.unsupported = error.Message;
            }
            return plan;
        }

        private static InkMetrics Measure(string text, string fontName, FontStyle style, double pointSize)
        {
            using (var family = new FontFamily(fontName))
            using (var format = (StringFormat)StringFormat.GenericTypographic.Clone())
            using (var path = new GraphicsPath())
            {
                format.FormatFlags |= StringFormatFlags.NoWrap | StringFormatFlags.MeasureTrailingSpaces;
                path.AddString(text, family, (int)style, (float)pointSize, PointF.Empty, format);
                RectangleF bounds = path.GetBounds();
                if (path.PointCount == 0 || bounds.Width <= 0 || bounds.Height <= 0) throw new InvalidOperationException("所选字体无法生成文字墨迹边界。");
                return new InkMetrics { x = bounds.X, y = bounds.Y, width = bounds.Width, height = bounds.Height,
                    lineHeight = pointSize * family.GetLineSpacing(style) / family.GetEmHeight(style) };
            }
        }

        private static void RequireFinite(params double[] values)
        {
            foreach (double value in values) if (!Program.Finite(value)) throw new InvalidOperationException("文字几何数据不是有限数值。");
        }
    }
}
