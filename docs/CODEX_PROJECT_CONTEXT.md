# CODEX_PROJECT_CONTEXT

本文件给后续 Codex 会话快速恢复项目上下文使用，不包含客户资料、真实样本文档内容或输出结果。

## 1. 项目定位

这是一个 Windows 本地运行的交付文件保护工具，面向商业计划书、可研报告、融资材料、测算模型等客户预览文件。

核心目标是：在不把文件上传到外部服务的前提下，把 PDF/Word/Excel 处理成更适合客户预览、公开展示或内部流转的低风险版本。

## 2. 当前功能

- PDF/Word 脱敏详情图：支持 PDF、DOCX、DOC 上传，输出图片化脱敏 PDF、分页详情图 ZIP、分页预览图和 JSON 报告。
- 自动脱敏：基于 PyMuPDF 读取可复制文本，结合 `rules_sensitive.json` 和正则规则识别公司、项目、地址、金额、财务指标、手机号、身份证、邮箱、银行卡、项目代码等敏感内容。
- 手动遮挡：前端可在页面预览图上框选区域，保存为人工遮挡，并重新生成当前页预览或导出结果。
- Excel 客户预览版：支持 `.xlsx` / `.xlsm`，`.xls` 需要先另存为 `.xlsx`。
- Excel 水印预览：上传后先读取工作表列表，可选择单个工作表生成水印效果 PNG 预览。
- Excel 图片化防复制版：默认模式，使用 Excel/WPS COM 真实截图工作表显示效果，烙印水印后插入新 `.xlsx`，不保留原始单元格数据。
- Excel 普通锁定版：保留真实单元格结构，公式转值、添加水印、保护工作表和工作簿结构。

## 3. 关键业务场景

- 给客户看商业计划书或可研报告的页面效果，但隐藏公司名、项目名、联系方式、金额和财务指标等敏感内容。
- 把 PDF 页面导出成适合详情页、朋友圈、案例展示或预览沟通的分页图片。
- 给客户发 Excel 测算模型预览版，让客户能看样式和大概内容，同时尽量降低直接复制单元格数据、公式和原表结构的风险。

## 4. 重要技术路线

- Web 服务：`app.py` 启动 Flask，只监听 `127.0.0.1`。
- PDF 文本识别：`desensitizer/processor.py` 用 PyMuPDF 读取文本和坐标。
- PDF/Word 流程：PDF 直接复制为 `work/source.pdf`；DOC/DOCX 通过 LibreOffice headless 转为 PDF。
- 脱敏详情图：从原始 `source.pdf` 渲染页面 PNG，再用 Pillow 做马赛克、模糊、浅灰马赛克或黑块。
- 脱敏 PDF：由处理后的分页图片重新合成为 PDF，避免保留可复制文字层。
- Excel 工作表读取：`openpyxl` 读取 sheet 列表和基础元信息。
- Excel 真实截图：优先 Microsoft Excel COM，失败后尝试 WPS `Ket.Application` COM。
- 图片化 Excel：批量截图时应复用一个 COM 应用实例和一个打开的工作簿，逐 sheet 截图，减少窗口反复弹出。
- 近似绘制：`openpyxl + Pillow` 只能作为用户显式允许的兜底，不可作为默认客户交付路径。

## 5. 不能破坏的功能

- PDF 上传后要能生成分页预览图，并能手动框选补充遮挡。
- 导出脱敏 PDF 必须是图片化 PDF，不保留原始可复制文字层。
- Word 自动转 PDF 只能作为预览辅助，界面和报告必须提示可能发生排版变化。
- Excel 图片化防复制版必须默认使用真实截图，不能默认降级成近似绘制。
- 图片化模式必须跳过公式转数值，因为输出文件只保留截图图片。
- 普通锁定版仍要支持公式转值、水印、工作表保护和工作簿结构保护。
- 下载预览版 Excel 和下载处理报告 JSON 必须分别指向对应文件，不能串。
- 输出到桌面默认只复制 Excel；JSON 报告只有用户勾选时才复制到桌面。
- 重复导出同名文件必须自动生成 `_v2`、`_v3` 等版本，不覆盖旧文件。

## 6. GitHub 安全红线

GitHub 只应保存代码、配置、启动脚本、规则文件和说明文档。

严禁提交：

- `uploads/` 中的上传缓存。
- `output/` 中的处理结果。
- `private_samples/` 中的真实测试文件。
- `.venv/`。
- Word、PDF、Excel、图片、ZIP/RAR、JSON 处理报告、日志等运行产物。

当前 `.gitignore` 已忽略这些运行产物，并通过 README 文件保留空目录说明。

## 7. Excel 图片化防复制版核心原则

- 默认推荐给客户预览使用。
- 输出仍是 `.xlsx`，但每个工作表主要是一张真实截图图片。
- 水印烙印在截图图片上，而不是只作为可删除的普通浮层。
- 不设置打开密码，客户可以直接打开浏览。
- 不保留原始单元格数据和公式，降低复制风险。
- 不能防截图、拍照或 OCR，界面和报告需要保留风险提示。
- 真实截图失败时默认停止，并提示修复 Excel/WPS COM 环境。
- `openpyxl + Pillow` 近似绘制可能丢失图表、图片、形状、分页和复杂排版，不能默认用于正式交付。

## 8. 后续每次修改前必须先阅读

- `README.md`
- `PROJECT_FLOW.md`
- `app.py`
- `desensitizer/processor.py`
- `desensitizer/excel_preview.py`
- `templates/index.html`
- `static/app.js`
- `rules_sensitive.json`
- `.gitignore`
- `docs/CODEX_PROJECT_CONTEXT.md`

## 9. 常见问题和排查方向

- Word 转 PDF 页数异常：优先检查 LibreOffice 转换结果，建议用户用 WPS/Word 手动导出 PDF 再上传。
- PDF 自动脱敏漏识别：确认是否为可复制文字型 PDF；扫描件和图片文字当前默认不启用 OCR。
- 手动遮挡坐标不准：检查前端图片显示尺寸、原图 naturalWidth/naturalHeight、PDF 坐标换算。
- 图片化 Excel 截图失败：确认 Excel/WPS COM 是否可用、文件是否被占用、是否能打开原文件、是否需要切换可见调试模式。
- 每个 sheet 都弹窗口：检查是否走了 `capture_workbook_sheets_with_com()` 的批量截图路径，而不是逐 sheet 创建 COM 应用。
- 近似绘制被使用：检查前端 `excelFallbackMode`、后端 `allow_approximate_fallback` 和报告中的 `screenshot_engine`。
- 下载按钮串文件：检查 `with_excel_urls()`、`outputs.excel`、`outputs.report`、`excel_download_url`、`report_download_url`。
- 桌面输出覆盖：检查 `choose_output_version()` 是否同时扫描程序 output 和桌面同名文件。

## 10. 用户偏好的输出逻辑

- Excel 默认输出到桌面，方便直接取用。
- 默认只把预览版 Excel 输出到桌面。
- JSON 处理报告默认保留在程序 `output` 目录，页面提供下载按钮。
- 用户勾选“同时输出处理报告 JSON 到桌面”时，才复制报告到桌面。
- 同名文件不覆盖旧文件，自动生成 `_v2`、`_v3`。
- 所有客户文件、输出文件和真实样本只留在本机，不上传 GitHub。
