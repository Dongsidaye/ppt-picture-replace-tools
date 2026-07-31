# Changelog

## 1.1.2

- 新增 Windows 一键安装器 `PictureReplaceTools-WPS-1.1.2.exe`。
- 安装器写入当前用户 `%APPDATA%\\kingsoft\\wps\\jsaddons`，合并 `publish.xml` 并自动备份原配置。
- 增加可审计的 `install-wps.ps1` / `uninstall-wps.ps1` 安装与卸载脚本。

## 1.1.0

- 新增 PowerPoint 剪贴板原位替换与剪贴板批量替换。
- 新增独立 WPS 演示 JavaScript 加载项源码（`wps_addin/`），遵循 `ribbon.xml + main.js` 与 `wpsjs publish` 部署规范。
- WPS 版支持文件选择、剪贴板 PNG/位图/JPG/GIF 粘贴、裁剪几何恢复和多页同源批量替换。
- 增加 50 MB 文件上限、临时文件清理和失败安全提示。
- 增加 WPS 运行时兼容性诊断，并在批量操作中显示匹配/成功/失败明细。
- 已在本机检测到 WPS WPP 12.1.0.28043 与 COM 宿主；JS 加载项静态构建通过，剪贴板/FileSystem 端到端替换仍需关闭现有用户文档后验收。

## 1.0.0

- 新增“原位替换图片”：保留位置、显示尺寸、旋转、翻转、层级与裁剪焦点。
- 新增“批量替换同图”：扫描当前演示文稿所有幻灯片，按同源图片批量替换。
- 支持不同页面对同一原图设置不同裁剪效果。
- 支持水平/垂直翻转实例识别。
- 使用显式导出缩放和 WIA 像素网格，修复 PowerPoint 导出取整造成的同源图漏检。
- 完成真实 PPTX 回归：3 个共同引用同一媒体部件的实例全部命中，异图保持不变。
