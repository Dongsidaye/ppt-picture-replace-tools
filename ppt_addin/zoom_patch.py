"""PowerPoint-only Smart Zoom typography patch.

``prepare_shared.py`` imports :func:`prepare_zoom` after it has created the
PowerPoint Smart Zoom implementation. The WPS source remains untouched.
"""
from __future__ import annotations


def _replace_once(text: str, old: str, new: str, label: str) -> str:
    count = text.count(old)
    if count != 1:
        raise ValueError(f"{label}: expected one target, found {count}")
    return text.replace(old, new, 1)


def prepare_zoom(text: str) -> str:
    """Keep text layout geometrically proportional during Smart Zoom."""
    # Shared snapshots are checked in after generation. Allow test tools to run
    # this transform against that already-prepared output without double-patch
    # failures; real generation always starts from the WPS-derived clean text.
    if "SMART_ZOOM_HOST_MIN_TEXT_SIZE_PT" in text:
        return text
    text = _replace_once(text, "  const SMART_ZOOM_MIN_TEXT_SIZE_PT = 8;\n",
        "  const SMART_ZOOM_MIN_TEXT_SIZE_PT = 8;\n"
        "  // A PowerPoint host limit, not a readability preference.\n"
        "  const SMART_ZOOM_HOST_MIN_TEXT_SIZE_PT = 1;\n",
        "Smart Zoom host font limit")
    text = _replace_once(text, "      protectTextReadability: true,\n",
        "      // Default to geometric fidelity; a font floor changes soft wrapping.\n"
        "      protectTextReadability: false,\n",
        "Smart Zoom default typography policy")
    text = _replace_once(text,
        '    { key: "lineRuleAfter", option: "scaleText", paths: [["TextFrame", "TextRange", "ParagraphFormat", "LineRuleAfter"]] },\n',
        "    // LineRuleAfter is an enum, not a physical distance.\n",
        "Smart Zoom line-rule enum")
    text = _replace_once(text,
        '    { key: "textSize", option: "scaleText", paths: [["TextFrame2", "TextRange", "Font", "Size"], ["TextFrame", "TextRange", "Font", "Size"]] },\n',
        '    { key: "textSize", option: "scaleText", paths: [["TextFrame2", "TextRange", "Font", "Size"], ["TextFrame", "TextRange", "Font", "Size"]] },\n'
        '    { key: "textSpacing", option: "scaleText", paths: [["TextFrame2", "TextRange", "Font", "Spacing"], ["TextFrame", "TextRange", "Font", "Spacing"]] },\n',
        "Smart Zoom character spacing")
    text = _replace_once(text, '''      const result = spec.key === "textSize"
        ? smartZoomReadTextSize(shape, spec.paths)
        : smartZoomReadFirstNumber(shape, spec.paths);
      if (!result.ok) continue;
      // A mixed/empty WPS font size is commonly exposed as 0 or -2.  It is''', '''      const result = spec.key === "textSize"
        ? smartZoomReadTextSize(shape, spec.paths)
        : smartZoomReadFirstNumber(shape, spec.paths);
      if (!result.ok) continue;
      // -2 is msoTriStateMixed for a text range. It is not a tracking value;
      // enumerate runs instead of multiplying a sentinel into every glyph.
      if (spec.key === "textSpacing" && result.value === -2) {
        styles._textSpacingMixed = true;
        continue;
      }
      // A mixed/empty WPS font size is commonly exposed as 0 or -2.  It is''', "Smart Zoom mixed character spacing")
    text = _replace_once(text, '''      result.push({
        shape: run,
        styles: { textSize: { value: size.value, path: ["Font", "Size"], option: "scaleText" } }
      });''', '''      const styles = { textSize: { value: size.value, path: ["Font", "Size"], option: "scaleText" } };
      const spacing = smartZoomReadFirstNumber(run, [["Font", "Spacing"]]);
      if (spacing.ok && spacing.value !== -2) styles.textSpacing = { value: spacing.value, path: ["Font", "Spacing"], option: "scaleText" };
      result.push({ shape: run, styles: styles });''', "Smart Zoom rich-text spacing")
    text = _replace_once(text, '''        const styles = smartZoomSnapshotStyles(cellShape, true);
        const textRuns = styles.textSize ? [] : smartZoomSnapshotTextRuns(cellShape, context);
        if (Object.keys(styles).length || textRuns.length) {
          result.push({ shape: cellShape, styles: styles, textRuns: textRuns });
        }''', '''        const styles = smartZoomSnapshotStyles(cellShape, true);
        const textRuns = styles.textSize ? [] : smartZoomSnapshotTextRuns(cellShape, context);
        const textLayout = smartZoomSnapshotTextLayout(cellShape);
        const paragraphMetrics = smartZoomSnapshotParagraphMetrics(cellShape);
        if (Object.keys(styles).length || textRuns.length || textLayout.length || paragraphMetrics.length) {
          result.push({ shape: cellShape, styles: styles, textRuns: textRuns, textLayout: textLayout, paragraphMetrics: paragraphMetrics, tableText: [], children: [] });
        }''', "Smart Zoom table text layout snapshot")

    helpers = r'''  // AutoSize may resize a shape between geometry and font writes. Freeze it
  // for an active zoom, retain WordWrap, and restore both only on Reset. A
  // committed scale keeps original automatic sizing disabled so it cannot
  // immediately reflow and invalidate the selected geometric ratio.
  function smartZoomSnapshotTextLayout(shape) {
    const result = [];
    const candidates = [
      { owner: "TextFrame", auto: ["TextFrame", "AutoSize"], wrap: ["TextFrame", "WordWrap"] },
      { owner: "TextFrame2", auto: ["TextFrame2", "AutoSize"], wrap: ["TextFrame2", "WordWrap"] }
    ];
    for (let i = 0; i < candidates.length; i += 1) {
      const item = candidates[i];
      const auto = smartZoomReadFirst(shape, [item.auto]);
      const wrap = smartZoomReadFirst(shape, [item.wrap]);
      if (!auto.ok && !wrap.ok) continue;
      result.push({ owner: item.owner, auto: auto.ok ? auto.value : undefined, wrap: wrap.ok ? wrap.value : undefined });
    }
    return result;
  }

  function smartZoomPrepareTextLayout(node) {
    const layouts = node && node.textLayout || [];
    for (let i = 0; i < layouts.length; i += 1) {
      const item = layouts[i];
      if (item.auto !== undefined) smartZoomWritePath(node.shape, [item.owner, "AutoSize"], 0);
      if (item.wrap !== undefined) smartZoomWritePath(node.shape, [item.owner, "WordWrap"], item.wrap);
    }
    (node && node.tableText || []).forEach(function (entry) { smartZoomPrepareTextLayout(entry); });
    (node && node.children || []).forEach(smartZoomPrepareTextLayout);
  }

  function smartZoomRestoreTextLayout(node) {
    const layouts = node && node.textLayout || [];
    for (let i = 0; i < layouts.length; i += 1) {
      const item = layouts[i];
      if (item.auto !== undefined) smartZoomWritePath(node.shape, [item.owner, "AutoSize"], item.auto);
      if (item.wrap !== undefined) smartZoomWritePath(node.shape, [item.owner, "WordWrap"], item.wrap);
    }
    (node && node.tableText || []).forEach(function (entry) { smartZoomRestoreTextLayout(entry); });
    (node && node.children || []).forEach(smartZoomRestoreTextLayout);
  }

  function smartZoomParagraphAt(textRange, index) {
    let paragraphs = null;
    try { paragraphs = textRange && textRange.Paragraphs; } catch (_) {}
    if (!paragraphs) return null;
    if (typeof paragraphs === "function") {
      try { const value = paragraphs.call(textRange, index, 1); if (value) return value; } catch (_) {}
      try { const value = paragraphs.call(textRange, index); if (value) return value; } catch (_) {}
    }
    if (hasMethod(paragraphs, "Item")) { try { return paragraphs.Item(index); } catch (_) {} }
    try { return paragraphs[index] || null; } catch (_) { return null; }
  }

  function smartZoomSnapshotParagraphMetrics(shape) {
    const ranges = smartZoomTextRanges(shape);
    const result = [];
    const seen = Object.create(null);
    for (let r = 0; r < ranges.length; r += 1) {
      for (let i = 1; i <= SMART_ZOOM_MAX_OBJECTS; i += 1) {
        const paragraph = smartZoomParagraphAt(ranges[r], i);
        if (!paragraph) break;
        const key = smartZoomTextRunKey(paragraph) || ("paragraph:" + r + ":" + i);
        if (seen[key]) break;
        seen[key] = true;
        const metrics = {};
        ["LeftMargin", "RightMargin", "FirstLineIndent"].forEach(function (name) {
          const value = smartZoomReadNumber(paragraph, ["ParagraphFormat", name]);
          if (value.ok) metrics[name] = value.value;
        });
        // PowerPoint uses MsoTriState here: 0 means a points value, -1 a
        // lines multiplier. Only physical point values are scalable.
        const beforeRule = smartZoomReadNumber(paragraph, ["ParagraphFormat", "LineRuleBefore"]);
        const afterRule = smartZoomReadNumber(paragraph, ["ParagraphFormat", "LineRuleAfter"]);
        const lineRule = smartZoomReadNumber(paragraph, ["ParagraphFormat", "LineRuleWithin"]);
        const before = smartZoomReadNumber(paragraph, ["ParagraphFormat", "SpaceBefore"]);
        const after = smartZoomReadNumber(paragraph, ["ParagraphFormat", "SpaceAfter"]);
        const within = smartZoomReadNumber(paragraph, ["ParagraphFormat", "SpaceWithin"]);
        if (beforeRule.ok && beforeRule.value === 0 && before.ok) metrics.SpaceBefore = before.value;
        if (afterRule.ok && afterRule.value === 0 && after.ok) metrics.SpaceAfter = after.value;
        if (lineRule.ok && lineRule.value === 0 && within.ok) metrics.SpaceWithin = within.value;
        if (Object.keys(metrics).length) result.push({ shape: paragraph, metrics: metrics });
      }
      if (result.length) break;
    }
    return result;
  }

  function smartZoomApplyParagraphMetrics(entries, factor) {
    if (!smartZoomSession || smartZoomSession.options.scaleText === false) return;
    const list = entries || [];
    for (let i = 0; i < list.length; i += 1) {
      const entry = list[i];
      Object.keys(entry.metrics || {}).forEach(function (name) {
        smartZoomWritePath(entry.shape, ["ParagraphFormat", name], entry.metrics[name] * factor);
      });
    }
  }

  function smartZoomMinimumTextSize(node) {
    let minimum = 0;
    function include(styles) {
      const value = smartZoomTextSizeFromStyles(styles);
      if (value > 0 && (!minimum || value < minimum)) minimum = value;
    }
    function walk(current) {
      if (!current) return;
      include(current.styles);
      (current.textRuns || []).forEach(function (run) { include(run && run.styles); });
      (current.tableText || []).forEach(function (cell) {
        include(cell && cell.styles);
        (cell && cell.textRuns || []).forEach(function (run) { include(run && run.styles); });
      });
      (current.children || []).forEach(walk);
    }
    walk(node);
    return minimum;
  }

  function smartZoomValidateHostTextFloor(nodes, factor, options) {
    if (!isFinite(factor) || factor >= 1 || !options || options.scaleText === false) return;
    let minimum = 0;
    for (let i = 0; i < nodes.length; i += 1) {
      const size = smartZoomMinimumTextSize(nodes[i]);
      if (size > 0 && (!minimum || size < minimum)) minimum = size;
    }
    if (!(minimum > 0) || minimum * factor >= SMART_ZOOM_HOST_MIN_TEXT_SIZE_PT) return;
    const minPercent = Math.ceil((SMART_ZOOM_HOST_MIN_TEXT_SIZE_PT / minimum) * 10000) / 100;
    throw new Error("等比缩放会使文字小于 PowerPoint 1pt 下限；最小允许比例为 " + minPercent + "% 。");
  }

'''
    text = _replace_once(text, "  function smartZoomSnapshotNode(shape, context, depth) {\n", helpers + "  function smartZoomSnapshotNode(shape, context, depth) {\n", "Smart Zoom layout helpers")
    text = _replace_once(text, '''      styles: styles,
      textRuns: styles.textSize ? [] : smartZoomSnapshotTextRuns(shape, snapshotContext),
      tableText: smartZoomSnapshotTableText(shape, snapshotContext),
      children: []''', '''      styles: styles,
      textRuns: (styles.textSize && !styles._textSpacingMixed) ? [] : smartZoomSnapshotTextRuns(shape, snapshotContext),
      textLayout: smartZoomSnapshotTextLayout(shape),
      paragraphMetrics: smartZoomSnapshotParagraphMetrics(shape),
      tableText: smartZoomSnapshotTableText(shape, snapshotContext),
      children: []''', "Smart Zoom node text layout snapshot")
    text = _replace_once(text, '''  function smartZoomApplyStyles(node, factor) {
    smartZoomApplyStyleSet(node.shape, node.styles, factor);
    smartZoomApplyTextRuns(node.textRuns, factor);
    const tableText = node.tableText || [];
    const tableFactor = smartZoomTableTextFactor(tableText, factor);
    for (let i = 0; i < tableText.length; i += 1) {
      smartZoomApplyStyleSet(tableText[i].shape, tableText[i].styles, tableFactor, SMART_ZOOM_MIN_TABLE_TEXT_SIZE_PT);
      smartZoomApplyTextRuns(tableText[i].textRuns, tableFactor, SMART_ZOOM_MIN_TABLE_TEXT_SIZE_PT);
    }''', '''  function smartZoomApplyStyles(node, factor) {
    smartZoomApplyStyleSet(node.shape, node.styles, factor);
    smartZoomApplyParagraphMetrics(node.paragraphMetrics, factor);
    smartZoomApplyTextRuns(node.textRuns, factor);
    const tableText = node.tableText || [];
    const tableFactor = smartZoomTableTextFactor(tableText, factor);
    for (let i = 0; i < tableText.length; i += 1) {
      smartZoomApplyStyleSet(tableText[i].shape, tableText[i].styles, tableFactor, SMART_ZOOM_MIN_TABLE_TEXT_SIZE_PT);
      smartZoomApplyTextRuns(tableText[i].textRuns, tableFactor, SMART_ZOOM_MIN_TABLE_TEXT_SIZE_PT);
      smartZoomApplyParagraphMetrics(tableText[i].paragraphMetrics, tableFactor);
    }''', "Smart Zoom paragraph application")
    text = _replace_once(text, '''    styleKeys.forEach(function (key) {
      const style = styles[key];''', '''    styleKeys.forEach(function (key) {
      if (key.charAt(0) === "_") return;
      const style = styles[key];''', "Smart Zoom skip style metadata")
    text = _replace_once(text, '''    if (geometryDirty || stylesDirty) {
      for (let i = 0; i < smartZoomSession.nodes.length; i += 1) {
        if (geometryDirty) smartZoomApplyGeometry(smartZoomSession.nodes[i], factor);
        if (stylesDirty) smartZoomApplyStyles(smartZoomSession.nodes[i], factor);
      }
    }''', '''    // Validate the raw requested sizes before any host write. The optional
    // readability floor may alter wrapping, but the 1pt host floor never may.
    if (geometryDirty || stylesDirty) smartZoomValidateHostTextFloor(smartZoomSession.nodes, factor, smartZoomSession.options);
    if (geometryDirty || stylesDirty) {
      for (let i = 0; i < smartZoomSession.nodes.length; i += 1) {
        const node = smartZoomSession.nodes[i];
        smartZoomPrepareTextLayout(node);
        if (geometryDirty) smartZoomApplyGeometry(node, factor);
        if (stylesDirty) smartZoomApplyStyles(node, factor);
      }
    }''', "Smart Zoom host-floor preflight and auto-size freeze")
    text = _replace_once(text, '''  function smartZoomReset(options) {
    const config = options && typeof options === "object" ? options : {};
    config._force = true;
    return smartZoomApply(100, config);
  }''', '''  function smartZoomReset(options) {
    const config = options && typeof options === "object" ? options : {};
    config._force = true;
    const info = smartZoomApply(100, config);
    if (!info || !info.stale) {
      for (let i = 0; smartZoomSession && i < smartZoomSession.nodes.length; i += 1) {
        smartZoomRestoreTextLayout(smartZoomSession.nodes[i]);
      }
    }
    return info;
  }''', "Smart Zoom reset text layout")
    return text
