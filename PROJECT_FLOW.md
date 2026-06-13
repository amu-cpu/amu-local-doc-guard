# PROJECT_FLOW

本文件用于代码审查和流程分析，不包含真实客户文件、测试原文或处理结果。

## 当前工具完整处理流程

1. 启动与页面
   - `app.py` 启动 Flask 本地服务，默认监听 `127.0.0.1:5000`，也可通过环境变量 `APP_PORT` 改端口。
   - 首页路由是 `GET /`，渲染 `templates/index.html`。
   - 健康检查路由是 `GET /health`，返回 `{"status": "ok"}`。

2. 上传文件后走哪个接口
   - 前端 `static/app.js` 在点击“开始处理”后，把批量文件作为 `files` 字段提交到 `POST /api/process`。
   - 后端从 `request.files.getlist("files")` 读取所有上传文件。

3. 文件保存到哪里
   - 每个上传文件调用 `desensitizer.processor.process_uploaded_file()`。
   - 后端为每个文件生成一个任务目录：`output/<原文件名安全化>_<随机id>/`。
   - 原始上传文件保存到：`output/<job_id>/input/<原文件名>`。
   - 统一处理用 PDF 保存到：`output/<job_id>/work/source.pdf`。
   - 原文件不会覆盖。

4. DOC/DOCX 如何转 PDF
   - `desensitizer/processor.py` 的 `convert_to_source_pdf()` 负责转换。
   - 如果上传的是 PDF，直接复制为 `work/source.pdf`。
   - 如果上传的是 DOC 或 DOCX，通过 LibreOffice headless 执行 `--convert-to pdf --outdir <work_dir>`。
   - LibreOffice 查找逻辑在 `find_soffice()`，优先读取 `LIBREOFFICE_PATH`，然后优先使用 `F:\Tools\LibreOffice\program\soffice.com` 和 F 盘 LibreOffice 路径，再回退到 PATH 和 C 盘常见安装路径。

5. PDF 如何提取文本
   - `detect_auto_redactions()` 用 PyMuPDF 打开 PDF。
   - `detect_page_redactions()` 逐页处理。
   - `iter_text_lines()` 调用 `page.get_text("rawdict")`，按文本行读取每个字符及其 bbox 坐标。
   - 后续规则根据字符区间合并 bbox，得到需要遮挡的矩形区域。

6. 规则在哪里
   - 商业计划书/可研报告关键词规则在 `rules_sensitive.json`。
   - `subject_keywords`：公司名称、建设单位、融资主体、股东、客户和供应商等主体类关键词。
   - `project_keywords`：项目名称、建设地点、项目地址、项目代码、备案代码等项目类关键词。
   - `financial_keywords`：总投资、融资金额、银行贷款、收入、利润、IRR、NPV、回收期、DSCR 等财务类关键词。
   - `business_keywords`：客户资源、资金用途、估值、股权结构、高管团队等商业敏感类关键词。
   - `desensitizer/processor.py` 保留手机号、身份证、邮箱、银行卡、统一社会信用代码、项目代码、金额和财务指标等正则。
   - 处理时会从封面和前 5 页提取临时敏感词种子，并在全文复用。

7. 如何生成脱敏 PDF
   - `reprocess_job()` 先运行 `detect_auto_redactions()`，再合并人工遮挡记录。
   - `apply_pdf_redactions()` 使用 PyMuPDF 的 `page.add_redact_annot()` 和 `page.apply_redactions()`。
   - PDF 中遮挡填充为黑色，并保存为：`output/<job_id>/脱敏后_<原文件名>.pdf`。

8. 如何生成分页图片
   - `render_pdf_pages()` 从原始 `source.pdf` 每页渲染成 PNG，不再从黑块 PDF 渲染。
   - 图片输出到：`output/<job_id>/pages/脱敏后_<原文件名>_page_001.png`。
   - `apply_image_masks()` 使用 Pillow 对图片区域二次处理，支持黑块、马赛克、模糊和浅灰马赛克。
   - `create_images_zip()` 把全部分页图片打包成：`output/<job_id>/脱敏后_<原文件名>_分页图片.zip`。

9. 手动框选如何保存
   - 前端预览页在图片上用 canvas 框选区域。
   - 点击“添加遮挡”后提交到 `POST /api/manual-redaction`。
   - 后端 `add_manual_redaction()` 根据图片坐标和页面尺寸换算为 PDF 坐标。
   - 人工遮挡记录写入报告的 `redactions` 列表，字段包含 `sensitive_type`、`page`、`rect`、`method`、`manual`、`rule`。
   - 人工遮挡会触发 `reprocess_job()` 重新生成 PDF、分页图片、ZIP 和报告。

10. 输出文件在哪里
    - 所有输出都在 `output/<job_id>/`。
    - 脱敏 PDF：`脱敏后_<原文件名>.pdf`。
    - 分页 PNG：`pages/脱敏后_<原文件名>_page_001.png` 等。
    - 图片 ZIP：`脱敏后_<原文件名>_分页图片.zip`。
    - 处理报告：`脱敏后_<原文件名>_处理报告.json`。

## 当前脱敏规则说明

1. 已识别的敏感类型
   - 手机号：`phone`。
   - 身份证号：`id_card`。
   - 邮箱：`email`。
   - 银行卡号：`bank_card`。
   - 统一社会信用代码：`unified_social_credit_code`。
   - 标签后的内容：`label_value`。
   - 财务字段后的数字或取值：`financial_value`。
   - 人工框选区域：`manual`。

2. 正则或关键词规则
   - 手机号：匹配中国大陆 11 位手机号，形如 `1[3-9]` 开头。
   - 身份证号：匹配 18 位身份证号，包含出生年月日和末位数字或 X。
   - 邮箱：匹配常见 `name@domain.tld` 格式。
   - 银行卡号：匹配 16 到 19 位数字，允许中间有空格或短横线。
   - 统一社会信用代码：匹配 18 位大写字母和数字组合，排除容易混淆的字符。
   - 标签关键词：见 `LABELS`。
   - 财务关键词：见 `FINANCIAL_FIELDS`。
   - 财务数字：匹配金额、万元、亿元、元、百分比、年、月、区间等常见表达。

3. 目前不支持自动识别的内容
   - 不在关键词列表里的公司名称、客户名称、项目简称、品牌名、人名。
   - 没有固定标签的封面标题、页眉页脚、落款、页脚版权信息。
   - 图片里的文字、扫描件里的文字、水印里的文字。
   - 复杂表格中没有关键词提示的敏感单元格。
   - 多行跨段落的标签值，例如标签在一行、具体内容在下一行且没有连续文本关系时可能漏识别。

4. 表格、页眉页脚、图片中文字、扫描件支持情况
   - 表格：如果表格文字能被 PyMuPDF 提取为可复制文本，并且命中现有关键词或正则，会尽量遮挡；复杂表格仍可能漏遮。
   - 页眉页脚：如果能提取为文本并命中规则，可以遮挡；未命中关键词的页眉页脚不会自动遮挡。
   - 图片中文字：第一版不支持自动识别。
   - 扫描件：第一版不启用 OCR，因此扫描件不能自动识别图片文字。

## 当前已知不足

1. 公司名称、项目名称、建设单位和地址可能漏遮
   - 只有命中 `LABELS` 且内容与标签在同一可提取文本行附近时，才会自动遮挡。
   - 封面大标题、独立公司名、项目简称、页眉页脚中的名称可能漏遮。

2. 融资金额和财务指标可能漏遮
   - 只有命中 `FINANCIAL_FIELDS` 后的数字或取值才会遮挡。
   - 没有关键词的金额、表格中拆分严重的数值、图片化图表里的数字可能漏遮。

3. OCR 未启用
   - `desensitizer/ocr.py` 只预留 `PaddleOCRProvider` 接口。
   - 当前报告中 `ocr_enabled` 为 `False`。

4. 当前主要处理文字型 PDF
   - 可复制文字型 PDF、由 DOC/DOCX 转出的文字型 PDF 是第一版重点。
   - 扫描版 PDF 或图片型 PDF 需要后续接入 OCR 后才能自动识别。

5. 手动框选对 PDF 和图片的影响
   - 人工框选会重新生成脱敏 PDF，并在 PDF 中用黑块遮挡对应区域。
   - 人工可选择黑块、马赛克、模糊或浅灰马赛克，导出的图片会用 Pillow 按所选方式处理。
   - 为保证 PDF 中文本不可复制，PDF 里仍统一使用黑块 redaction。

## 安全边界说明

1. 这个诊断包不包含真实客户文件。
2. 这个诊断包不包含 `uploads/`、`output/`、`private_samples/` 里的真实内容。
3. 这个诊断包不包含 `.doc`、`.docx`、`.pdf`、`.png`、`.jpg`、`.jpeg`、`.json`、`.zip`、`.log` 文件。
4. 这个诊断包只用于代码审查和流程分析。
