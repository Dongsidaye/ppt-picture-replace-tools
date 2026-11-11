/* Optional OCR choice and explicit review of recovered CAD labels. */
(function (global) {
  "use strict";
  const dialog = document.getElementById("cadOcrChoice");
  const results = document.getElementById("cadOcrResults");
  const items = document.getElementById("cadOcrItems");
  const preview = document.getElementById("cadImportPreview");
  if (!dialog || !results || !items) return;
  let finishChoice = null;
  let rows = [];
  let page = null;
  let originalSvg = "";
  let previewUrl = "";

  function settle(value) {
    const complete = finishChoice;
    finishChoice = null;
    if (dialog.open) dialog.close();
    if (complete) complete(value);
  }

  document.getElementById("cadOcrEnable").addEventListener("click", function () { settle(true); });
  document.getElementById("cadOcrSkip").addEventListener("click", function () { settle(false); });
  document.getElementById("cadOcrAbort").addEventListener("click", function () { settle(null); });
  dialog.addEventListener("cancel", function (event) { event.preventDefault(); settle(null); });

  function selections() {
    return rows.filter(row => row.check.checked).map(row => ({ index: row.index, text: row.text.value }));
  }

  function highlight() {
    if (!originalSvg || !page || !preview || !global.DOMParser) return;
    const svg = new DOMParser().parseFromString(originalSvg, "image/svg+xml");
    const ns = "http://www.w3.org/2000/svg";
    selections().forEach(selection => {
      const line = page.lines[selection.index];
      const polygon = svg.createElementNS(ns, "polygon");
      const box = page.viewBox;
      const points = line.points.map(point => (box[0] + point.x * box[2] / page.width) + "," + (box[1] + point.y * box[3] / page.height));
      polygon.setAttribute("points", points.join(" "));
      polygon.setAttribute("fill", "none");
      polygon.setAttribute("stroke", "#2563eb");
      polygon.setAttribute("stroke-width", "2");
      polygon.setAttribute("vector-effect", "non-scaling-stroke");
      svg.documentElement.appendChild(polygon);
    });
    if (previewUrl) URL.revokeObjectURL(previewUrl);
    previewUrl = URL.createObjectURL(new Blob([new XMLSerializer().serializeToString(svg)], { type: "image/svg+xml" }));
    preview.src = previewUrl;
  }

  function clear() {
    settle(null);
    rows = [];
    page = null;
    originalSvg = "";
    items.replaceChildren();
    results.hidden = true;
    if (previewUrl) URL.revokeObjectURL(previewUrl);
    previewUrl = "";
  }

  function show(ocr, base64) {
    clear();
    page = ocr;
    if (!page || !Array.isArray(page.lines) || !page.lines.length) return;
    originalSvg = new TextDecoder().decode(Uint8Array.from(atob(base64), character => character.charCodeAt(0)));
    page.lines.forEach((line, index) => {
      const row = document.createElement("label");
      row.className = "cad-ocr-item";
      const check = document.createElement("input");
      check.type = "checkbox";
      check.checked = false;
      check.setAttribute("aria-label", "恢复第 " + (index + 1) + " 项识别文字");
      const text = document.createElement("input");
      text.type = "text";
      text.maxLength = 2000;
      text.value = String(line.text || "");
      text.setAttribute("aria-label", "核对识别文字");
      const confidence = document.createElement("small");
      confidence.textContent = Math.round(Number(line.confidence || 0) * 100) + "%";
      row.append(check, text, confidence);
      check.addEventListener("change", highlight);
      rows.push({ index: index, check: check, text: text });
      items.appendChild(row);
    });
    results.hidden = false;
  }

  global.PptCadOcr = {
    choose: function () {
      settle(null);
      return new Promise(resolve => { finishChoice = resolve; dialog.showModal(); });
    },
    show: show,
    clear: clear,
    setDisabled: function (disabled) { rows.forEach(row => { row.check.disabled = disabled; row.text.disabled = disabled; }); },
    selections: selections
  };
}(window));
