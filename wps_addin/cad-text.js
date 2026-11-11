/* Native editable text for DWG import. Geometry and font metrics come from the CAD worker. */
(function (global) {
  "use strict";

  function number(value, label) {
    if (value === null || value === undefined || value === "" || typeof value === "boolean") {
      throw new Error(label + "无效。");
    }
    var result = Number(value);
    if (!isFinite(result)) throw new Error(label + "无效。");
    return result;
  }

  function positive(value, label) {
    var result = number(value, label);
    if (result <= 0) throw new Error(label + "无效。");
    return result;
  }

  function optional(work) {
    try { work(); } catch (_) {}
  }

  function setNumber(owner, name, value, tolerance, label) {
    owner[name] = value;
    if (Math.abs(number(owner[name], label) - value) > tolerance) {
      throw new Error("WPS 无法应用" + label + "。");
    }
  }

  function setFrameNumber(shape, frame, name, value, tolerance, label) {
    optional(function () { frame[name] = value; });
    try {
      if (Math.abs(number(frame[name], label) - value) <= tolerance) return;
    } catch (_) {}
    // WPS exposes a few TextFrame2 setters but implements them only through
    // the original TextFrame interface. Confirm the shared state afterwards.
    shape.TextFrame[name] = value;
    if (Math.abs(number(frame[name], label) - value) > tolerance) {
      throw new Error("WPS 无法应用" + label + "。");
    }
  }

  function flag(value) {
    return value === true || value === -1 || value === 1;
  }

  function fontAlias(value) {
    var name = String(value || "").replace(/^\s+|\s+$/g, "").toLowerCase();
    if (name === "微软雅黑") return "microsoft yahei";
    if (name === "宋体") return "simsun";
    return name;
  }

  function setFontName(font, property, name) {
    font[property] = name;
    if (fontAlias(font[property]) !== fontAlias(name)) {
      throw new Error("WPS 无法应用所选文字字体。");
    }
  }

  function rgb(value) {
    var text = String(value || "#000000");
    if (!/^#[0-9a-f]{6}$/i.test(text)) throw new Error("DWG 文字颜色无效。");
    var color = parseInt(text.substr(1), 16);
    return ((color >>> 16) & 255) | (color & 65280) | ((color & 255) << 16);
  }

  function paint(font, color) {
    font.Fill.Visible = -1;
    font.Fill.Solid();
    font.Fill.ForeColor.RGB = color;
    if ((number(font.Fill.ForeColor.RGB, "文字颜色") & 16777215) !== color) {
      throw new Error("WPS 无法应用文字颜色。");
    }
  }

  function clearLine(line, color) {
    optional(function () { line.ForeColor.RGB = color; });
    optional(function () { line.Weight = 0; });
    optional(function () { line.Transparency = 1; });
    optional(function () { line.Visible = 0; });
  }

  function clearDepth(depth) {
    optional(function () { depth.Depth = 0; });
    optional(function () { depth.BevelTopType = 1; });
    optional(function () { depth.BevelBottomType = 1; });
    optional(function () { depth.BevelTopDepth = 0; });
    optional(function () { depth.BevelBottomDepth = 0; });
    optional(function () { depth.BevelTopInset = 0; });
    optional(function () { depth.BevelBottomInset = 0; });
    optional(function () { depth.Visible = 0; });
  }

  function clearShapeStyle(shape, frame, color) {
    optional(function () { shape.Fill.Visible = 0; });
    optional(function () { clearLine(shape.Line, color); });
    optional(function () { shape.Shadow.Visible = 0; });
    optional(function () { shape.Glow.Radius = 0; });
    optional(function () { shape.Reflection.Type = 0; });
    optional(function () { shape.SoftEdge.Radius = 0; });
    optional(function () { clearDepth(shape.ThreeD); });
    optional(function () { clearDepth(frame.ThreeD); });
  }

  function clearFontStyle(font, color) {
    optional(function () { clearLine(font.Line, color); });
    optional(function () { font.Shadow.Visible = 0; });
    optional(function () { font.Glow.Radius = 0; });
    optional(function () { font.Reflection.Type = 0; });
    optional(function () { font.SoftEdgeFormat = 0; });
  }

  function plainWarp(shape, frame) {
    setNumber(shape.TextEffect, "PresetShape", 1, 0, "文字变形类型");
    setNumber(shape.TextEffect, "NormalizedHeight", 0, 0, "文字高度模式");
    // WPS does not synchronize WarpFormat when only PresetShape is assigned.
    // 36 is Office msoWarpFormat37 and DrawingML textPlain.
    setNumber(frame, "WarpFormat", 36, 0, "文字横向变形");
    setNumber(frame, "AutoSize", 0, 0, "文字自动调整模式");
  }

  function textElements(text) {
    var count = 0;
    var joined = false;
    for (var index = 0; index < text.length; index += 1) {
      var code = text.charCodeAt(index);
      if (code >= 56320 && code <= 57343 && index > 0) {
        var previous = text.charCodeAt(index - 1);
        if (previous >= 55296 && previous <= 56319) continue;
      }
      if (code === 8205) { joined = true; continue; }
      if ((code >= 768 && code <= 879) || (code >= 6832 && code <= 6911) ||
          (code >= 7616 && code <= 7679) || (code >= 8400 && code <= 8447) ||
          (code >= 65024 && code <= 65039) || (code >= 65056 && code <= 65071)) continue;
      if (joined) { joined = false; continue; }
      count += 1;
    }
    return count;
  }

  function paintedPart(range, text, start, length) {
    try { return range.Characters(start + 1, length); }
    catch (error) {
      if (start === 0 && length === text.length) return range;
      throw new Error("WPS 无法测量带首尾空白的文字。");
    }
  }

  function verifyFit(width, height, targetWidth, targetHeight) {
    positive(width, "文字墨迹宽度");
    positive(height, "文字墨迹高度");
    if (Math.abs(width - targetWidth) > Math.max(0.5, targetWidth * 0.05)) {
      throw new Error("替换字体后的文字宽度差异较大，需保留 SVG 文字。");
    }
    if (Math.abs(height - targetHeight) > Math.max(0.5, targetHeight * 0.15)) {
      throw new Error("替换字体后的文字高度差异较大，需保留 SVG 文字。");
    }
  }

  function rotation(shape, angle) {
    var normalized = ((angle % 360) + 360) % 360;
    shape.Rotation = normalized;
    var actual = ((number(shape.Rotation, "文字角度") % 360) + 360) % 360;
    var difference = Math.abs(actual - normalized);
    if (Math.min(difference, 360 - difference) > 0.11) throw new Error("WPS 无法应用文字角度。");
    return normalized;
  }

  function add(slide, plan, selectedFont) {
    if (!slide || !slide.Shapes || !plan || !plan.item) throw new Error("缺少 DWG 文字插入信息。");
    if (plan.unsupported) throw new Error(String(plan.unsupported));
    var item = plan.item;
    var text = String(item.Text === null || item.Text === undefined ? "" : item.Text);
    if (!/\S/.test(text)) return null;
    if (/[\r\n]/.test(text)) throw new Error("DWG 多行文字需要按原行位置拆分。");
    if (text.length > 32767) throw new Error("单行 DWG 文字过长，请拆分后导入。");
    var pointSize = positive(plan.fontSize, "文字字号");
    if (pointSize < 1 || pointSize > 4000) throw new Error("文字字号超出 WPS 的 1–4000 磅范围。");
    var sourceMetrics = plan.metrics;
    if (!sourceMetrics) throw new Error("缺少文字字体度量。");
    var metrics = {
      x: number(sourceMetrics.x, "文字左边距"),
      y: number(sourceMetrics.y, "文字上边距"),
      width: positive(sourceMetrics.width, "文字墨迹宽度"),
      height: positive(sourceMetrics.height, "文字墨迹高度"),
      lineHeight: positive(sourceMetrics.lineHeight, "文字行高")
    };
    var targetWidth = positive(plan.targetWidth, "文字目标宽度");
    var targetHeight = positive(plan.targetHeight, "文字目标高度");
    var centerX = number(plan.centerX, "文字位置");
    var centerY = number(plan.centerY, "文字位置");
    var angle = number(item.Rotation, "文字角度");
    var fontName = String(selectedFont || "Microsoft YaHei").replace(/^\s+|\s+$/g, "");
    if (!fontName) throw new Error("请选择文字字体。");
    var color = rgb(item.Color);
    var useWarp = Math.abs(targetWidth / metrics.width - 1) > 0.03;
    var shape = null;
    try {
      if (useWarp) {
        shape = slide.Shapes.AddTextEffect(0, text, fontName, pointSize,
          flag(item.Bold) ? -1 : 0, flag(item.Italic) ? -1 : 0, 0, 0);
      } else {
        shape = slide.Shapes.AddTextbox(1, 0, 0,
          Math.max(targetWidth * 1.2, pointSize * 2), Math.max(targetHeight * 1.5, pointSize * 2));
      }
      var frame = shape.TextFrame2;
      var range = frame.TextRange;
      // WPS initializes some frame properties lazily when text is assigned.
      range.Text = text;
      if (String(range.Text) !== text) throw new Error("WPS 未保留完整文字内容。");
      setNumber(frame, "AutoSize", 0, 0, "文字自动调整模式");
      setFrameNumber(shape, frame, "WordWrap", 0, 0, "文字换行模式");
      setFrameNumber(shape, frame, "MarginLeft", 0, 0.01, "文字左边距");
      setFrameNumber(shape, frame, "MarginRight", 0, 0.01, "文字右边距");
      setFrameNumber(shape, frame, "MarginTop", 0, 0.01, "文字上边距");
      setFrameNumber(shape, frame, "MarginBottom", 0, 0.01, "文字下边距");
      setNumber(frame, "VerticalAnchor", 1, 0, "文字垂直对齐");
      if (useWarp) plainWarp(shape, frame);
      clearShapeStyle(shape, frame, color);

      setNumber(range.ParagraphFormat, "Alignment", 1, 0, "文字水平对齐");
      setNumber(range.ParagraphFormat, "SpaceBefore", 0, 0.01, "文字段前间距");
      setNumber(range.ParagraphFormat, "SpaceAfter", 0, 0.01, "文字段后间距");
      optional(function () { range.ParagraphFormat.LineRuleWithin = -1; });
      optional(function () { range.ParagraphFormat.SpaceWithin = 1; });
      var font = range.Font;
      setFontName(font, "Name", fontName);
      setFontName(font, "NameFarEast", fontName);
      optional(function () { font.NameAscii = fontName; });
      optional(function () { font.NameComplexScript = fontName; });
      setNumber(font, "Size", pointSize, 0.11, "文字字号");
      font.Bold = flag(item.Bold) ? -1 : 0;
      font.Italic = flag(item.Italic) ? -1 : 0;
      if (flag(font.Bold) !== flag(item.Bold) || flag(font.Italic) !== flag(item.Italic)) {
        throw new Error("WPS 无法应用文字粗斜体。");
      }
      setNumber(font, "Spacing", 0, 0.01, "文字字距");
      paint(font, color);
      clearFontStyle(font, color);

      // WPS rounds native font sizes to a tenth of a point on some builds.
      var metricScale = number(font.Size, "文字字号") / pointSize;
      metrics.x *= metricScale;
      metrics.y *= metricScale;
      metrics.width *= metricScale;
      metrics.height *= metricScale;
      metrics.lineHeight *= metricScale;

      if (useWarp) {
        plainWarp(shape, frame);
        setNumber(shape, "LockAspectRatio", 0, 0, "文字宽高锁定");
        setNumber(shape, "Width", targetWidth, 0.11, "文字宽度");
        setNumber(shape, "Height", targetHeight, 0.11, "文字高度");
        verifyFit(number(shape.Width, "文字宽度"), number(shape.Height, "文字高度"), targetWidth, targetHeight);
        setNumber(shape, "Left", centerX - targetWidth / 2, 0.11, "文字横向位置");
        setNumber(shape, "Top", centerY - targetHeight / 2, 0.11, "文字纵向位置");
        rotation(shape, angle);
      } else {
        var first = text.search(/\S/);
        var trimmed = text.replace(/^\s+|\s+$/g, "");
        var painted = paintedPart(range, text, first, trimmed.length);
        var zeroSpacingWidth = positive(painted.BoundWidth, "文字测量宽度");
        var gaps = textElements(trimmed) - 1;
        if (gaps > 0) {
          for (var iteration = 0; iteration < 3; iteration += 1) {
            var inkWidth = metrics.width + positive(painted.BoundWidth, "文字测量宽度") - zeroSpacingWidth;
            var difference = targetWidth - inkWidth;
            if (Math.abs(difference) <= Math.max(0.05, targetWidth * 0.005)) break;
            var next = number(font.Spacing, "文字字距") + difference / gaps;
            var bounded = Math.max(-font.Size * 0.6, Math.min(font.Size * 2, next));
            setNumber(font, "Spacing", bounded, 0.11, "文字字距");
            if (Math.abs(next - bounded) > 0.001) break;
          }
        }
        shape.Width = Math.max(targetWidth, positive(range.BoundWidth, "文字测量宽度")) + 0.1;
        shape.Height = Math.max(targetHeight, positive(range.BoundHeight, "文字测量高度")) + 0.1;
        var fittedWidth = metrics.width + positive(painted.BoundWidth, "文字测量宽度") - zeroSpacingWidth;
        verifyFit(fittedWidth, metrics.height, targetWidth, targetHeight);
        var inkLeft = number(painted.BoundLeft, "文字测量位置") + metrics.x;
        var inkTop = number(painted.BoundTop, "文字测量位置") +
          (positive(painted.BoundHeight, "文字测量高度") - metrics.lineHeight) / 2 + metrics.y;
        var localX = inkLeft + fittedWidth / 2 - shape.Left - shape.Width / 2;
        var localY = inkTop + metrics.height / 2 - shape.Top - shape.Height / 2;
        var degrees = rotation(shape, angle);
        var radians = degrees * Math.PI / 180;
        var rotatedX = localX * Math.cos(radians) - localY * Math.sin(radians);
        var rotatedY = localX * Math.sin(radians) + localY * Math.cos(radians);
        setNumber(shape, "Left", centerX - shape.Width / 2 - rotatedX, 0.11, "文字横向位置");
        setNumber(shape, "Top", centerY - shape.Height / 2 - rotatedY, 0.11, "文字纵向位置");
      }

      // Geometry changes can restore a preset's outline. Clear it last and set
      // its color too, so a host ignoring transparency cannot restore blue ink.
      paint(font, color);
      clearFontStyle(font, color);
      optional(function () { clearLine(shape.Line, color); });
      if (String(range.Text) !== text) throw new Error("WPS 插入后文字内容发生变化。");
      if (useWarp && number(frame.WarpFormat, "文字横向变形") !== 36) {
        throw new Error("WPS 未保留文字横向变形。");
      }
      return shape;
    } catch (error) {
      var failure = error;
      if (!failure || (typeof failure !== "object" && typeof failure !== "function")) failure = new Error(String(error));
      if (shape) {
        try { shape.Delete(); }
        catch (cleanupError) {
          failure.createdShape = shape;
          failure.cleanupError = String(cleanupError && (cleanupError.message || cleanupError.description) || cleanupError);
        }
      }
      throw failure;
    }
  }

  global.WpsCadText = { add: add };
}(window));
