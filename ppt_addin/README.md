# 东四大爷的工具箱 · PowerPoint

PowerPoint 桌面版采用原生 COM 加载项、分类 Ribbon 和 WebView2 任务窗格，复用 WPS 2.1.15 的工具引擎。它独立于旧版仅提供图片替换的 PPAM。

功能包括图片清单与链接更新、文件/剪贴板单图和同源批替、智能缩放、分类对象管理、对象筛选、样式刷、批量文字、排版、清理、页面导出、颜色和 Photoshop 往返编辑。逐项映射及宿主边界见 [PARITY.md](PARITY.md)。

## 安装

运行 `PictureReplaceTools-PPT-3.0.3.exe`。通常安装到当前用户目录并注册当前用户 COM 加载项；如果电脑原本已关闭 UAC 且以管理员身份运行，Windows 会忽略用户级 COM 注册，安装器改用 Program Files 和机器级注册。安装不会修改 UAC 或 Office 安全设置，不关闭未保存的文稿，不修改 WPS 安装。首次安装会尝试连接当前 PowerPoint；已有旧版本加载时，下次正常启动 PowerPoint 生效。

环境：Windows、桌面 Microsoft PowerPoint 2016 或更高版本、.NET Framework 4.8、Microsoft Edge WebView2 Runtime。32/64 位加载器均随包提供；实际验证的平台记录在测试报告中。

卸载运行安装目录的 `uninstall.ps1`，只移除本工具箱注册，保留文稿与可恢复的安装文件。旧 PPAM 可以在 PowerPoint 加载项管理中停用，避免重复图片工具入口。

## 构建与验证

```powershell
node .\tests\adapter.test.js
node .\tests\port-regressions.test.js
node .\tests\thumbnail-regression.test.js
node .\tests\zoom-regression.test.js
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Package
```

构建需要已安装的 Office Interop、Python 和 Node.js。打包首次运行前在仓库的 `wps_addin` 执行 `npm ci`，以取得现有打包工具链中的 7-Zip 和 SFX 模块；也可不加 `-Package` 先构建目录安装包。

默认构建使用已提交的 `shared` 快照，避免 WPS 的独立更新改变 PowerPoint 构建内容。更新共享引擎时，先运行 `python .\prepare_shared.py`，或给构建加 `-RefreshShared`；生成器从同仓库 `wps_addin` 读取资源，精确校验适配补丁并记录来源 SHA-256，不覆盖 WPS 文件。WebView2 SDK 固定版本与 SHA-256。安装前校验包内全部文件。

真实宿主测试通过 `COMAddIns.Item("Dongsidaye.PowerPointTools").Object.RunSelfTest(reportPath)` 调用内置测试脚本，只创建/关闭自己的测试文稿，不退出 PowerPoint。`Status()` 返回原生任务窗格、WebView2 与加载错误状态。日志：`%LOCALAPPDATA%\Dongsidaye\PowerPointTools\host.log`。

更新使用独立的 `ppt-v*` Release 和 `PictureReplaceTools-PPT-*` 安装包。远端网页无法在带 Office 桥接权限的面板内导航；外链交给系统浏览器。

PowerPoint 的文稿提取采用自有文件副本和 [`Slides.InsertFromFile`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slides.insertfromfile)，避免全局剪贴板变化干扰提取。图片字节通过 Base64 跨越 WebView2/COM 边界，避免二进制字符串中的空字节被截断。

designed by Dongsidaye
