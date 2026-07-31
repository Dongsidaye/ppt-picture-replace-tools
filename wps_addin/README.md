# Picture Replace Tools — WPS WPP 版

这是面向 WPS 演示（WPP）的独立 JavaScript 加载项版本。它不依赖 PowerPoint 的 `.ppam` 或 VBA，遵循 WPS 官方 `ribbon.xml + main.js` Web 加载项结构。

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

`wpsjs debug` 会启动本地服务并连接 WPS。由于 WPS 版本、个人版/企业版及策略配置会影响 JSAPI，必须在目标 WPS 实机上验证剪贴板格式和 FileSystem 权限。

## 发布

官方推荐使用 `publish.xml` 流程：

```powershell
npm run publish
```

将生成的 `wps-addon-build` 文件部署到服务器，并把 `wps-addon-publish/publish.html` 地址提供给用户。不要把 `.ppam` 改名或当作 WPS JS 加载项安装包；WPS 新版本也不应把修改 `oem.ini/jsplugins.xml` 作为公开产品的主安装方式。

## 已知边界

- 本目录需要 WPS WPP 实机验收；当前开发机没有 WPS Windows，仓库只能完成静态构建检查。
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
