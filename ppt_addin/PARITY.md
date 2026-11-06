# PowerPoint 3.0.0 功能迁移对照

本版本以仓库当前 WPS **2.1.15** 为功能来源，使用 C# COM 宿主和 WebView2 前端。`prepare_shared.py` 从 `wps_addin` 生成 `shared`，不会修改 WPS 源文件；生成的界面保留既有命令、面板、分组与图标。JavaScript 的 `WpsPictureReplace` / `CurrentWPSAddIn` 名称作为桥接兼容接口保留，不代表运行在 WPS。

本文是代码覆盖与边界记录。资源生成、语法检查或构建通过，不代表 PowerPoint 实际宿主中的每项功能已经验收。

## 功能覆盖

| 分类 | 迁移功能 | 来源入口 |
|---|---|---|
| 图片替换 | 文件/剪贴板单图替换、同源批替；保留裁剪、位置、尺寸、旋转、翻转、层级与可读取的图片效果；进度和取消 | `replacePictureKeepCrop`、`replaceAllMatching` |
| 图片清单 | 普通页/母版/版式图片、缩略图、同源与裁剪窗口分组、选中实例替换、搜索/排序/筛选、正文定位、名称修改 | `collectDeckImages`、`locateSlideShape` |
| 图片链接 | 来源路径记录、已修改/丢失检查、更新链接、解除链接 | `parseLink`、`updateLinkedInstances` |
| 清单加载 | 结构先显示、缩略图分批补齐、内存/磁盘缓存、缓存失效、文稿事件后台预热 | `collectDeckImagesCached`、`preloadDeckImages` |
| 智能缩放 | 1%–300%、五个锚点、目标宽度、初始快照、重新拾取/确认结束/恢复初始；文字与效果开关、表格统一字号比例 | `smartZoomBegin`、`smartZoomApply` |
| 筛选 | 全选/反选、相同类型/字号/宽高/颜色、线条/文字/组合、原生动画窗格 | `objectFilterRun`、`OpenAnimationPane` |
| 对象管理 | 类型颜色分组、定位、单个及同类型批量锁定/解锁和显示/隐藏、选择本类、框选辅助 | `layerList`、`layerSetLockedMany`、`layerSetVisibleMany` |
| 样式刷 | 拾取并选择性应用填充、轮廓、文字、几何、阴影、映像、发光、柔边、3D | `designStyleCapture`、`designStyleApply` |
| 文字 | 当前页/全文查找替换、大小写与整字匹配、文字交换、提取文字和备注/复制 | `designTextFindReplace`、`designTextExtract` |
| 排版 | 对齐、页面居中、等距、交换位置、统一大小/宽高/比例/角度、矩阵、环形 | `designAlignRun` |
| 清理 | 真空白页、备注、动画、隐藏对象、完全画外对象；删除前面板确认 | `designCleanup` |
| 导出 | PNG/JPG、当前/所选/全文、DPI、盖印、提取所选页到新文稿 | `designExportSlides`、`designExtractSlides` |
| 颜色 | 填充/轮廓/文字 RGB 容差替换、HSL 调整 | `designColorReplace`、`designColorAdjust` |
| Photoshop | 选择/检测应用、导出编辑、临时文件位置、原位载回、撤销载回 | `designPhotoshopOpen`、`designPhotoshopUndo` |
| 插件外壳 | Office 2009 Ribbon 分类/菜单/右键入口、独立图标、GitHub、作者、版本、诊断、PPT 独立更新 | `ribbon.xml`、`checkForUpdates` |

## PowerPoint 副本专用修正

1. 提取幻灯片先固定源文稿、所选页码和页面尺寸，通过 `SaveCopyAs` 保存唯一自有临时副本，再逐页 `InsertFromFile` 导入新文稿并核对返回数和页数增量，不使用全局剪贴板。无宏副本使用 PPTX（24），含宏或无法判断宏状态使用 PPTM（25）；只用于页面导入，不访问 VBA 项目。未保存源也可提取，不改源 `Saved`/路径。失败只关闭本次创建的目标，始终删除自有快照；关闭失败保留原生回收登记。
2. 空白页清理只有在 `Shapes.Count == 0`、没有备注文本、没有主序列/交互动画，且所有检查成功时才允许删除。隐藏内容不再被当作空白；读取失败保留页面。
3. 首次打开清单前也允许文稿事件预热，保留 3 秒去抖和原来的分片扫描。`#idle` 路由隐藏所有面板，不在初始化时触发前台清单扫描。
4. 文字替换先计算带大小写和 Unicode 整字边界的字面匹配，再倒序修改 `Characters`，不会反复替换新插入的文字。出现部分失败时不整段回写覆盖已完成结果。仅在宿主完全没有 Characters 时使用带提示的整段回写。
5. 颜色工具在 HTML RGB 与 COM OLE BGR 的读写边界交换红蓝通道；样式刷等 COM 到 COM 属性复制不转换。
6. 临时文稿必须通过 `Presentations.Add(false)` 独立创建。失败直接中止，不把用户活动文稿当临时文稿，不设置用户文稿的 `Saved`。确认自有临时文稿后通过 adapter 的 `__pptTrackScratch(presentation, true)` 注册到原生退出回收列表；清理临时对象后只将自有文稿标记 `Saved=-1`，避免池化文稿造成退出保存提示。成功关闭才解除注册；关闭失败仍保留原生回收记录。
7. `OnAddInLoad` 保留选择事件和后台预热，但不运行 WPS 的任何诊断标志文件入口。更新仅读取 `ppt_addin/package.json`（产品名 `picture-replace-tools-ppt`），release tag 为 `ppt-v3.0.0`，安装资产为 `PictureReplaceTools-PPT-3.0.0.exe`。安装/重启委托宿主 `PptHostUpdate()`，不运行 WPS 安装器。
8. 文件替换核心不删除调用方传入的图片。仅由浏览器上传入口创建并登记的临时文件，可由对应界面在成功、失败或取消后释放；位于临时目录的用户文件也不会被当作自有文件删除。
9. Photoshop 从 `PictureFormat` 读取裁剪；有裁剪时导出完整图源，载回时再恢复原裁剪，避免重复裁剪。作业绑定文稿窗口，切换文稿后的载回与撤销拒绝执行；裁剪信息读取失败中止。
10. 剪贴板单图替换遇到原图快照不完整时保留原图，并清除本次新插入的图片。原图删除前的其他失败同样回滚新图；删除已开始后不冒险再删新图，避免同时丢失两者。生成缩略图或指纹时，源图片 `Copy` 失败立即终止，不粘贴剪贴板中的旧内容。

## 既有覆盖边界与宿主验收重点

- 图片清单目前扫描普通页顶层图片和一个 `SlideMaster` 的母版/版式图片，不递归组合内图片，也不枚举全部 `Designs`。这是源引擎边界，不应宣称任意复杂文稿全量覆盖。
- 对象管理仅列当前编辑区域的顶层最多 500 个对象；“本类全部”仅作用于清单已列项目。锁定需要宿主真实支持 `Shape.Locked` 或选择事件守卫，宿主能力不足时必须保留界面中的降级说明。
- 框选辅助会临时移动锁定对象到画外；切换/关闭面板、关闭文稿和退出应用时的位置恢复属于宿主生命周期验收项。
- 缩放最多快照 500 个节点、组深度 64；普通文本保护下限 8pt，表格用统一系数保护最小 4pt。宿主不暴露效果可见状态时跳过，避免意外开启效果。
- 样式刷是已列属性复制，未覆盖所有渐变/纹理/主题格式。颜色工具针对选中顶层对象的填充/轮廓/文字颜色，不递归组和表格单元格。
- “删除动画”当前仅删 `MainSequence`，不覆盖交互触发序列；空白页保护会检查两者，避免误删包含交互内容的页面。
- 排版沿用源行为：左/顶以首对象为基准，右/底和居中以页面为基准；统一角度以首对象为基准。部分界面说明仍称末对象，需要后续统一语义。
- 缓存验证检查结构、名称、几何、可见状态和裁剪；同尺寸原位更换图片、相同布局名的多母版等情况仍需用户刷新。后台生成缩略图会使用复制/粘贴，应验证剪贴板与编辑响应。
- PS 作业、缩放快照、样式来源属于运行期状态；文稿切换/关闭后必须防止使用失效 COM 引用。不能把静态 API 映射视为这些工作流已经通过。

## 生成资源接口

- `shared/ribbon.xml` 使用 Office 2009 namespace，保留所有菜单与按钮；`onAction=OnAction`、`getImage=GetImage`、根 `onLoad=OnAddInLoad`。WPS 的 `size="small"` 转为 Office 合法的 `size="normal"`，未知尺寸值拒绝生成。
- `shared/action-map.json`：控件 ID → 原 JavaScript 回调名称。调用需将原控件 ID 作为参数传入，供设计工具识别子命令。
- `shared/icon-map.json`：控件 ID → 图标资源文件名。
- HTML 加载顺序为 `host-adapter.js` → `main.js` → 原内联 UI 脚本 → `host-bootstrap.js`。宿主适配器文件由宿主项目维护，生成器不覆盖。
- `shared/source-manifest.json` 记录全部输入文件 SHA-256。生成器对重要修正函数先校验完整源码块的 SHA-256，再替换；源代码漂移会停止生成，要求人工复核。

## 共享资源验证记录

- 生成器连续运行结果一致；WPS 输入文件与记录的 SHA-256 一致，故意改动待修正源码块会被拒绝。
- 生成 `main.js` 与三个 HTML 中所有内联脚本均通过 Node.js 语法检查；脚本加载顺序已核对。
- Office Ribbon 保留原有所有节点，48 个动作控件、24 个带图标控件均有映射，19 个图标资产存在。
- 隔离 JavaScript 验证通过：10 个文字替换场景、6 个空白页保护场景、OLE RGB 往返及混合色跳过、3 个临时文稿生命周期、目标文稿激活后的源幻灯片提取。
- 对抗检查重点：正文含 `$&`/正则符号或替换文含查找词时不误展开/循环；逐段写入中途失败不整段覆盖；临时文稿创建失败不写用户文稿 `Saved`；更新清单产品名必须是 PPT，不能串用 WPS 发布。
- 以上未使用用户文稿，不替代 COM/WebView2 实际宿主交互、安装、视觉和性能验收。
- 页面提取隔离验证覆盖：未保存源、非连续选页及重复索引、有宏/未知宏状态、保存部分失败、导入返回 0/抛错/页数不符、错误返回源文稿、关闭失败、临时路径碰撞。依据：[SaveCopyAs](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.savecopyas)、[InsertFromFile](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slides.insertfromfile)、[保存格式枚举](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype)、[HasVBProject](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.hasvbproject)。

## 来源校验

<!-- SOURCE_HASHES_START -->

| WPS source | SHA-256 |
|---|---|
| `main.js` | `08ac802dd6151770dd583d9767fdd24c959df3a2f7c184c62eceb6347f1181ad` |
| `taskpane.html` | `f37e3d9d6a6ea027dc66dddc6395ddaec402897cd335a7588b591d6f54887ce0` |
| `ribbon.xml` | `fca963661d4952d181c4769bfaab9ab0602924d72ab89298e04037ca9b52263d` |

<!-- SOURCE_HASHES_END -->
