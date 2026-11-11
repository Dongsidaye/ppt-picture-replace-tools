"""PowerPoint-only DWG import controls and task-pane content."""
import xml.etree.ElementTree as ET


def _once(text, before, after, label):
    if text.count(before) != 1:
        raise ValueError(label + ": source drift")
    return text.replace(before, after, 1)


def prepare_cad_html(text):
    text = _once(text, '<button data-panel="photoshop">PS</button>',
                 '<button data-panel="photoshop">PS</button>\n        <button data-panel="cad">CAD</button>',
                 "CAD tool navigation")
    text = _once(text, 'photoshop: "PS助手" };', 'photoshop: "PS助手", cad: "导入 DWG" };',
                 "CAD tool title")
    panel = '''        <section id="tool-panel-cad" class="tool-panel">
          <div class="tool-card cad-import-card">
            <h3>导入 DWG</h3>
            <p class="tool-help">转换模型空间，检查 SVG 预览后插入当前幻灯片。图形采用透明背景并等比居中。</p>
            <div class="tool-row">
              <button id="cadImportChoose" class="primary">选择 DWG 文件…</button>
              <button id="cadImportCancel" hidden>取消导入</button>
            </div>
            <div id="cadImportFile" class="cad-import-file" hidden></div>
            <div id="cadImportStatus" class="tool-status" role="status" aria-live="polite">请选择 DWG 文件。</div>
            <div id="cadImportWarning" class="cad-import-warning" role="note" hidden></div>
            <div id="cadImportPreviewWrap" class="cad-import-preview" hidden>
              <img id="cadImportPreview" alt="DWG 转换后的 SVG 预览" />
            </div>
            <div class="tool-row cad-import-insert-row">
              <button id="cadImportInsert" class="primary" disabled>插入当前幻灯片</button>
            </div>
          </div>
        </section>
'''
    text = _once(text, '        <section id="tool-panel-style" class="tool-panel active">',
                 panel + '        <section id="tool-panel-style" class="tool-panel active">', "CAD import panel")
    styles = '''  <style>
    .cad-import-card .tool-help { margin: 0 0 14px; line-height: 1.65; }
    .cad-import-card [hidden] { display: none !important; }
    .cad-import-file { margin: 12px 0 4px; color: #3e526b; font-size: 12px; overflow-wrap: anywhere; }
    .cad-import-warning { margin: 12px 0; padding: 10px 12px; border: 1px solid #f0d4a3; border-radius: 8px; background: #fff8e9; color: #805416; font-size: 12px; line-height: 1.6; white-space: pre-wrap; overflow-wrap: anywhere; }
    .cad-import-preview { display: flex; align-items: center; justify-content: center; min-height: 190px; max-height: 400px; margin-top: 14px; padding: 14px; overflow: hidden; border: 1px solid #dbe3ed; border-radius: 10px; background-color: #fff; background-image: linear-gradient(45deg,#edf1f5 25%,transparent 25%),linear-gradient(-45deg,#edf1f5 25%,transparent 25%),linear-gradient(45deg,transparent 75%,#edf1f5 75%),linear-gradient(-45deg,transparent 75%,#edf1f5 75%); background-size: 18px 18px; background-position: 0 0,0 9px,9px -9px,-9px 0; }
    .cad-import-preview img { display: block; width: 100%; max-height: 360px; object-fit: contain; }
    .cad-import-insert-row { margin-top: 14px; }
    #cadImportInsert { flex: 1; }
  </style>
'''
    text = _once(text, '</head>', styles + '</head>', "CAD import styles")
    return _once(text, '</body>', '  <script src="dwg-import.js"></script>\n</body>', "CAD import script")


def prepare_cad_ribbon(root, icons, actions):
    namespace = root.tag.split('}')[0] + '}'
    tab = root.find('.//' + namespace + 'tab')
    if tab is None:
        raise ValueError("CAD ribbon: no toolbox tab")
    if any(element.get('id') == 'CadImportDwgButton' for element in root.iter()):
        raise ValueError("CAD ribbon: duplicate import button")
    group = ET.Element(namespace + 'group', {'id': 'CadImportGroup', 'label': 'CAD'})
    ET.SubElement(group, namespace + 'button', {
        'id': 'CadImportDwgButton', 'label': '导入DWG', 'size': 'large',
        'onAction': 'OnAction', 'getImage': 'GetImage',
        'screentip': '将 DWG 转为 SVG 插入',
        'supertip': '选择 DWG 文件，转换模型空间并检查 SVG 预览，然后等比插入当前幻灯片。',
    })
    export_group = next((child for child in tab if child.get('id') == 'DesignExportGroup'), None)
    if export_group is None:
        tab.append(group)
    else:
        tab.insert(list(tab).index(export_group), group)
    icons['CadImportDwgButton'] = 'icon_dwg.png'
    actions['CadImportDwgButton'] = 'OpenCadImportPane'
    return root
