# Picture Replace Tools — WPS WPP 版

这是面向 WPS 演示（WPP）的独立 JavaScript 加载项版本。它不依赖 PowerPoint 的 `.ppam` 或 VBA，遵循 WPS 官方 `ribbon.xml + main.js` Web 加载项结构。`index.html` 是一个最小入口，用于兼容当前 `wpsjs debug` 对根页面的检查；它只引入 `main.js`，不承载业务界面。

## 功能

- 文件原位替换：选择一张图片文件，替换当前选中的图片。
- 剪贴板原位替换：用截图、复制的图片或可粘贴为图片的对象替换当前图片。
- 文件批量替换：选中一个原图实例，跨全部幻灯片识别同源图片并批量替换。
- 剪贴板批量替换：以剪贴板图片作为新图，批量替换全部同源实例。
- 每个实例独立保留裁剪框、裁剪焦点、位置、尺寸、旋转、翻转、层级和名称。

## 开发与调试

需要 Node.js、WPS Windows（启用 JS 加载项）和 `wpsjs` CLI：

```powershell
npm install
npm run debug
```

`wpsjs debug` 会启动本地服务并连接 WPS。Windows 上如果当前 `wpsjs` CLI 版本无法自动写入 WPS 的 `oem.ini`，请使用 `npm run build` 生成静态包，再按 WPS 管理员提供的 `publish.xml` 地址部署；不要把仓库中的调试配置直接复制到生产环境。

## 发布

官方推荐使用 `publish.xml` 流程：

```powershell
npm run publish
```

将生成的 `wps-addon-build` 文件部署到服务器，并把 `wps-addon-publish/publish.html` 地址提供给用户。不要把 `.ppam` 改名或当作 WPS JS 加载项安装包；WPS 新版本也不应把修改 `oem.ini/jsplugins.xml` 作为公开产品的主安装方式。

## 一键安装包

仓库 Release 同时提供 `PictureReplaceTools-WPS-*.exe`。双击后安装器会把 WPP 加载项复制到当前 Windows 用户的 `%APPDATA%\\kingsoft\\wps\\jsaddons`，合并写入本地 `publish.xml`，并自动备份原有配置；安装完成后重启 WPS 即可。

如果需要自行构建安装器：

```powershell
npm.cmd install
npm.cmd run installer
```

安装器没有数字签名，Windows SmartScreen 可能显示未知发布者；确认文件来自本项目 Release 后选择继续。安装目录同时包含 `uninstall-wps.ps1`，可用以下命令卸载：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\\uninstall-wps.ps1
```

## 已知边界

- v1.1.6 起，参考图导出不再依赖 WPS JSAPI 未文档化的 `Shape.SaveAsPicture` / `PictureFormat.Crop` 子对象，改为“临时演示文稿 + `Slide.Export`”方案；已在 WPS WPP 12.1 真机验证：真实案例中同一原图的 3 个不同裁剪实例可被统一识别并分别保持各自裁剪，异图不受影响。
- 剪贴板读取使用 `PasteSpecial`，若剪贴板为空或没有 PNG/位图/JPG/GIF 格式，会安全失败且不会删除目标图。
- 批量识别使用临时无裁剪预览和 32×32 像素特征；极端压缩、透明背景、滤镜或不同 WPS 渲染器可能需要人工复核。
- 文件和剪贴板内容仅在本机 WPS 临时目录中短暂处理，插件不上传图片。

## 官方依据

- [WPS 加载项开发说明](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-addin-development/wps-addin-development-instructions)
- [自定义功能区概述](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/customize-ribbon/overview)
- [WPP Shapes 对象](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wpp/Shapes/obj)
- [WPP PictureFormat 对象](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wpp/PictureFormat/obj)
- [WPP View.PasteSpecial](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wpp/View/member/PasteSpecial)
- [WPS FileSystem 对象](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/FileSystem/obj)

