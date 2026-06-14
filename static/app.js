const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const fileList = document.getElementById("fileList");
const processBtn = document.getElementById("processBtn");
const clearBtn = document.getElementById("clearBtn");
const statusText = document.getElementById("statusText");
const results = document.getElementById("results");
const errors = document.getElementById("errors");
const defaultMaskMethod = document.getElementById("defaultMaskMethod");
const expectedPages = document.getElementById("expectedPages");
const tabButtons = document.querySelectorAll("[data-tab]");
const pdfPanel = document.getElementById("pdfPanel");
const excelPanel = document.getElementById("excelPanel");
const excelDropzone = document.getElementById("excelDropzone");
const excelFileInput = document.getElementById("excelFileInput");
const excelFileList = document.getElementById("excelFileList");
const excelProcessBtn = document.getElementById("excelProcessBtn");
const excelStatusText = document.getElementById("excelStatusText");
const excelErrors = document.getElementById("excelErrors");
const excelResults = document.getElementById("excelResults");
const excelWatermarkText = document.getElementById("excelWatermarkText");
const excelWatermarkFontSize = document.getElementById("excelWatermarkFontSize");
const excelWatermarkColor = document.getElementById("excelWatermarkColor");
const excelWatermarkOpacity = document.getElementById("excelWatermarkOpacity");
const excelWatermarkRotation = document.getElementById("excelWatermarkRotation");
const excelWatermarkSpacing = document.getElementById("excelWatermarkSpacing");
const excelProtectionPassword = document.getElementById("excelProtectionPassword");
const excelSecurityMode = document.getElementById("excelSecurityMode");
const excelConvertFormulas = document.getElementById("excelConvertFormulas");
const excelAddWatermark = document.getElementById("excelAddWatermark");
const excelProtectSheets = document.getElementById("excelProtectSheets");
const excelProtectWorkbook = document.getElementById("excelProtectWorkbook");
const excelIncludeHidden = document.getElementById("excelIncludeHidden");
const excelFallbackMode = document.getElementById("excelFallbackMode");
const excelOutputMode = document.getElementById("excelOutputMode");
const excelOutputReportToDesktop = document.getElementById("excelOutputReportToDesktop");
const excelWindowMode = document.getElementById("excelWindowMode");
const excelFormulaModeHint = document.getElementById("excelFormulaModeHint");
const excelSheetSelect = document.getElementById("excelSheetSelect");
const excelPreviewRange = document.getElementById("excelPreviewRange");
const excelPreviewBtn = document.getElementById("excelPreviewBtn");
const excelRefreshPreviewBtn = document.getElementById("excelRefreshPreviewBtn");
const excelPreviewStatus = document.getElementById("excelPreviewStatus");
const excelPreviewStage = document.getElementById("excelPreviewStage");
const excelPreviewPlaceholder = document.getElementById("excelPreviewPlaceholder");
const excelPreviewImage = document.getElementById("excelPreviewImage");
const excelPreviewFitBtn = document.getElementById("excelPreviewFitBtn");
const excelPreviewOriginalBtn = document.getElementById("excelPreviewOriginalBtn");
const excelPreviewZoomInBtn = document.getElementById("excelPreviewZoomInBtn");
const excelPreviewZoomOutBtn = document.getElementById("excelPreviewZoomOutBtn");

const modal = document.getElementById("previewModal");
const previewTitle = document.getElementById("previewTitle");
const previewImage = document.getElementById("previewImage");
const canvas = document.getElementById("selectionCanvas");
const submitManualBtn = document.getElementById("submitManualBtn");
const undoPageMaskBtn = document.getElementById("undoPageMaskBtn");
const clearPageMasksBtn = document.getElementById("clearPageMasksBtn");
const closeModalBtn = document.getElementById("closeModalBtn");
const maskMethod = document.getElementById("maskMethod");
const manualStatus = document.getElementById("manualStatus");
const ctx = canvas.getContext("2d");

let selectedFiles = [];
let selectedExcelFile = null;
let excelUploadId = "";
let excelUploadMeta = null;
let excelSheets = [];
let excelPreviewScale = 1;
let jobs = [];
let activeJob = null;
let activePage = null;
let dragStart = null;
let selection = null;

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropzone.classList.add("is-over");
});
dropzone.addEventListener("dragleave", () => dropzone.classList.remove("is-over"));
dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropzone.classList.remove("is-over");
  addFiles(event.dataTransfer.files);
});
fileInput.addEventListener("change", () => addFiles(fileInput.files));
clearBtn.addEventListener("click", clearFiles);
processBtn.addEventListener("click", processFiles);
tabButtons.forEach((button) => button.addEventListener("click", () => switchTab(button.dataset.tab)));
excelDropzone.addEventListener("click", () => excelFileInput.click());
excelDropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  excelDropzone.classList.add("is-over");
});
excelDropzone.addEventListener("dragleave", () => excelDropzone.classList.remove("is-over"));
excelDropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  excelDropzone.classList.remove("is-over");
  setExcelFile(event.dataTransfer.files[0]);
});
excelFileInput.addEventListener("change", () => setExcelFile(excelFileInput.files[0]));
excelProcessBtn.addEventListener("click", processExcelFile);
excelPreviewBtn.addEventListener("click", generateExcelWatermarkPreview);
excelRefreshPreviewBtn.addEventListener("click", generateExcelWatermarkPreview);
excelSheetSelect.addEventListener("change", () => {
  if (excelUploadId) excelPreviewStatus.textContent = "工作表已切换，请点击生成水印预览图查看效果。";
});
excelPreviewRange.addEventListener("change", markExcelWatermarkDirty);
excelSecurityMode.addEventListener("change", () => {
  syncExcelSecurityMode();
  markExcelWatermarkDirty();
});
excelFallbackMode.addEventListener("change", markExcelWatermarkDirty);
excelIncludeHidden.addEventListener("change", () => {
  populateExcelSheetSelect();
  markExcelWatermarkDirty();
});
[
  excelWatermarkText,
  excelWatermarkFontSize,
  excelWatermarkColor,
  excelWatermarkOpacity,
  excelWatermarkRotation,
  excelWatermarkSpacing,
  excelAddWatermark,
].forEach((control) => {
  control.addEventListener("input", markExcelWatermarkDirty);
  control.addEventListener("change", markExcelWatermarkDirty);
});
excelPreviewFitBtn.addEventListener("click", setExcelPreviewFit);
excelPreviewOriginalBtn.addEventListener("click", () => setExcelPreviewScale(1));
excelPreviewZoomInBtn.addEventListener("click", () => setExcelPreviewScale(excelPreviewScale + 0.2));
excelPreviewZoomOutBtn.addEventListener("click", () => setExcelPreviewScale(Math.max(0.4, excelPreviewScale - 0.2)));
excelPreviewImage.addEventListener("load", () => {
  if (excelPreviewImage.dataset.mode !== "scale") setExcelPreviewFit();
});
excelPreviewImage.addEventListener("error", () => {
  excelPreviewImage.hidden = true;
  excelPreviewImage.removeAttribute("src");
  excelPreviewPlaceholder.hidden = false;
  excelPreviewStatus.textContent = "预览图片加载失败，请重新生成预览图。";
  setExcelPreviewButtonsEnabled(Boolean(excelUploadId && excelSheetSelect.value));
  showExcelErrors([{ file: "Excel 水印预览", error: "预览图片加载失败，请重新生成预览图。" }]);
});
syncExcelSecurityMode();
closeModalBtn.addEventListener("click", closePreview);
previewImage.addEventListener("load", syncCanvas);
window.addEventListener("resize", syncCanvas);

canvas.addEventListener("mousedown", (event) => {
  if (modal.hidden) return;
  dragStart = pointerPosition(event);
  selection = null;
  setManualStatus("");
  drawSelection();
});

canvas.addEventListener("mousemove", (event) => {
  if (!dragStart) return;
  const point = pointerPosition(event);
  selection = normalizeRect(dragStart.x, dragStart.y, point.x - dragStart.x, point.y - dragStart.y);
  drawSelection();
});

window.addEventListener("mouseup", () => {
  if (!dragStart) return;
  dragStart = null;
  drawSelection();
});

submitManualBtn.addEventListener("click", submitManualRedaction);
undoPageMaskBtn.addEventListener("click", undoPageManualRedaction);
clearPageMasksBtn.addEventListener("click", clearPageManualRedactions);

function addFiles(fileListLike) {
  const incoming = Array.from(fileListLike).filter((file) => /\.(pdf|docx|doc)$/i.test(file.name));
  selectedFiles = selectedFiles.concat(incoming);
  renderFileList();
}

function renderFileList() {
  fileList.innerHTML = "";
  selectedFiles.forEach((file, index) => {
    const item = document.createElement("li");
    item.innerHTML = `<span>${escapeHtml(file.name)}</span><span>${formatSize(file.size)}</span>`;
    item.addEventListener("dblclick", () => {
      selectedFiles.splice(index, 1);
      renderFileList();
    });
    fileList.appendChild(item);
  });
  processBtn.disabled = selectedFiles.length === 0;
  clearBtn.disabled = selectedFiles.length === 0;
}

function clearFiles() {
  selectedFiles = [];
  fileInput.value = "";
  renderFileList();
}

function switchTab(tabName) {
  tabButtons.forEach((button) => button.classList.toggle("is-active", button.dataset.tab === tabName));
  pdfPanel.classList.toggle("is-active", tabName === "pdf");
  excelPanel.classList.toggle("is-active", tabName === "excel");
}

function setExcelFile(file) {
  excelErrors.hidden = true;
  excelErrors.innerHTML = "";
  if (!file) return;
  if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
    showExcelErrors([{ file: file.name, error: "请选择 .xlsx、.xlsm 文件；.xls 请先另存为 .xlsx 后处理。" }]);
    return;
  }
  resetExcelUploadState();
  selectedExcelFile = file;
  excelFileInput.value = "";
  renderExcelFile();
  inspectExcelFile(file);
}

function renderExcelFile() {
  excelFileList.innerHTML = "";
  if (selectedExcelFile) {
    const item = document.createElement("li");
    item.innerHTML = `<span>${escapeHtml(selectedExcelFile.name)}</span><span>${formatSize(selectedExcelFile.size)}</span>`;
    item.addEventListener("dblclick", () => {
      selectedExcelFile = null;
      resetExcelUploadState();
      renderExcelFile();
    });
    excelFileList.appendChild(item);
  }
  updateExcelActionState();
}

function resetExcelUploadState() {
  excelUploadId = "";
  excelUploadMeta = null;
  excelSheets = [];
  excelSheetSelect.innerHTML = `<option value="">请先上传 Excel</option>`;
  excelSheetSelect.disabled = true;
  excelPreviewRange.disabled = true;
  excelPreviewStatus.textContent = "";
  excelPreviewImage.hidden = true;
  excelPreviewImage.removeAttribute("src");
  excelPreviewPlaceholder.hidden = false;
  excelResults.innerHTML = "";
  setExcelPreviewButtonsEnabled(false);
  updateExcelActionState();
}

async function inspectExcelFile(file) {
  excelStatusText.textContent = "正在读取工作表...";
  excelProcessBtn.disabled = true;
  setExcelPreviewButtonsEnabled(false);
  const formData = new FormData();
  formData.append("file", file);
  try {
    const response = await fetch("/excel/inspect", { method: "POST", body: formData });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "读取 Excel 工作表失败");
    excelUploadMeta = payload.upload;
    excelUploadId = excelUploadMeta.upload_id;
    excelSheets = excelUploadMeta.sheets || [];
    populateExcelSheetSelect();
    excelStatusText.textContent = `已读取 ${excelSheets.length} 个工作表`;
    excelPreviewStatus.textContent = "请选择工作表后生成水印预览图。";
  } catch (error) {
    excelStatusText.textContent = "读取失败";
    showExcelErrors([{ file: file.name, error: error.message }]);
    resetExcelUploadState();
  } finally {
    updateExcelActionState();
  }
}

function populateExcelSheetSelect() {
  const previous = excelSheetSelect.value;
  const availableSheets = excelSheets.filter((sheet) => sheet.visible || excelIncludeHidden.checked);
  excelSheetSelect.innerHTML = "";
  if (!availableSheets.length) {
    excelSheetSelect.innerHTML = `<option value="">没有可预览的工作表</option>`;
    excelSheetSelect.disabled = true;
    excelPreviewRange.disabled = true;
    updateExcelActionState();
    return;
  }
  availableSheets.forEach((sheet) => {
    const option = document.createElement("option");
    option.value = sheet.name;
    option.textContent = sheet.visible ? sheet.name : `${sheet.name}（隐藏）`;
    excelSheetSelect.appendChild(option);
  });
  const defaultSheet = availableSheets.find((sheet) => sheet.name === previous)
    || availableSheets.find((sheet) => sheet.name === excelUploadMeta?.default_sheet)
    || availableSheets[0];
  excelSheetSelect.value = defaultSheet.name;
  excelSheetSelect.disabled = false;
  excelPreviewRange.disabled = false;
  updateExcelActionState();
}

function updateExcelActionState() {
  const ready = Boolean(excelUploadId && excelSheetSelect.value);
  excelProcessBtn.disabled = !excelUploadId;
  excelPreviewBtn.disabled = !ready;
  excelRefreshPreviewBtn.disabled = !ready;
}

function setExcelPreviewButtonsEnabled(enabled) {
  excelPreviewBtn.disabled = !enabled;
  excelRefreshPreviewBtn.disabled = !enabled;
  excelPreviewFitBtn.disabled = !enabled;
  excelPreviewOriginalBtn.disabled = !enabled;
  excelPreviewZoomInBtn.disabled = !enabled;
  excelPreviewZoomOutBtn.disabled = !enabled;
}

function buildExcelOptionsFormData() {
  const formData = new FormData();
  formData.append("watermark_text", excelWatermarkText.value || "保密文件");
  formData.append("watermark_font_size", excelWatermarkFontSize.value || "28");
  formData.append("watermark_color", excelWatermarkColor.value || "#b8b8b8");
  formData.append("watermark_opacity", excelWatermarkOpacity.value || "20");
  formData.append("watermark_rotation", excelWatermarkRotation.value || "-30");
  formData.append("watermark_spacing", excelWatermarkSpacing.value || "medium");
  formData.append("protection_password", excelProtectionPassword.value || "123456");
  formData.append("preview_security_mode", excelSecurityMode.value || "image_based");
  formData.append("image_range_type", excelPreviewRange.value || "auto");
  formData.append("allow_approximate_fallback", excelFallbackMode.value === "approximate" ? "1" : "0");
  formData.append("output_mode", excelOutputMode.value || "desktop");
  formData.append("output_report_to_desktop", excelOutputReportToDesktop.checked ? "1" : "0");
  formData.append("screenshot_window_mode", excelWindowMode.value || "quiet");
  appendCheckbox(formData, "convert_formulas", excelConvertFormulas);
  appendCheckbox(formData, "add_watermark", excelAddWatermark);
  appendCheckbox(formData, "protect_sheets", excelProtectSheets);
  appendCheckbox(formData, "protect_workbook_structure", excelProtectWorkbook);
  appendCheckbox(formData, "include_hidden_sheets", excelIncludeHidden);
  return formData;
}

function markExcelWatermarkDirty() {
  if (excelUploadId) {
    excelPreviewStatus.textContent = "参数已修改，请点击生成水印预览图查看效果。";
  }
}

function syncExcelSecurityMode() {
  const imageBased = (excelSecurityMode.value || "image_based") === "image_based";
  excelConvertFormulas.disabled = imageBased;
  excelConvertFormulas.checked = !imageBased;
  excelFallbackMode.disabled = !imageBased;
  if (excelFormulaModeHint) {
    excelFormulaModeHint.textContent = imageBased
      ? "图片化防复制版会直接截图当前显示效果，不需要公式转数值。"
      : "普通锁定版会将公式转为当前值，并尽量保留原数字格式。";
  }
}

async function processExcelFile() {
  if (!excelUploadId) {
    showExcelErrors([{ file: "Excel 文件", error: "请先上传 Excel，并等待工作表读取完成后再导出。" }]);
    return;
  }
  const fileName = selectedExcelFile?.name || excelUploadMeta?.original_name || "Excel 文件";
  excelStatusText.textContent = "正在导出 Excel...";
  excelProcessBtn.textContent = "正在导出...";
  excelProcessBtn.disabled = true;
  excelErrors.hidden = true;
  excelErrors.innerHTML = "";
  excelResults.innerHTML = "";

  const formData = buildExcelOptionsFormData();
  if (excelUploadId) {
    formData.append("upload_id", excelUploadId);
  }
  const slowTimer = window.setTimeout(() => {
    excelStatusText.textContent = (excelWindowMode.value || "quiet") === "quiet"
      ? "正在安静模式处理，WPS/Excel 可能会短暂出现在任务栏，这是正常现象。"
      : "正在调用 WPS/Excel 真实截图，窗口可能短暂出现，请勿手动关闭。";
  }, 3000);

  try {
    const response = await fetch("/excel/export", { method: "POST", body: formData });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "Excel 导出失败");
    await pollExcelExportTask(payload.task_id);
  } catch (error) {
    showExcelErrors([{ file: fileName, error: error.message }]);
    excelStatusText.textContent = "导出失败";
  } finally {
    window.clearTimeout(slowTimer);
    excelProcessBtn.textContent = "直接导出 Excel";
    updateExcelActionState();
  }
}

async function pollExcelExportTask(taskId) {
  while (true) {
    const response = await fetch(`/excel/export-status/${encodeURIComponent(taskId)}`);
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "读取 Excel 导出进度失败");
    if (payload.current && payload.total) {
      excelStatusText.textContent = `${payload.message || "正在导出"}（${payload.current} / ${payload.total} 个工作表）`;
    } else {
      excelStatusText.textContent = payload.message || "正在导出 Excel...";
    }
    if (payload.state === "done") {
      renderExcelResult(payload.job);
      excelStatusText.textContent = "导出完成";
      return;
    }
    if (payload.state === "error") {
      throw new Error(payload.message || "Excel 导出失败");
    }
    await wait(800);
  }
}

function appendCheckbox(formData, name, checkbox) {
  formData.append(name, checkbox.checked ? "1" : "0");
}

async function generateExcelWatermarkPreview() {
  if (!excelUploadId) {
    showExcelErrors([{ file: "Excel 水印预览", error: "当前文件不存在，请重新上传 Excel。" }]);
    return;
  }
  if (!excelSheetSelect.value) {
    showExcelErrors([{ file: "Excel 水印预览", error: "请选择一个工作表。" }]);
    return;
  }

  const formData = buildExcelOptionsFormData();
  formData.append("upload_id", excelUploadId);
  formData.append("sheet_name", excelSheetSelect.value);
  formData.append("range_type", excelPreviewRange.value || "used_range");

  excelErrors.hidden = true;
  excelErrors.innerHTML = "";
  excelPreviewStatus.textContent = "正在生成预览...";
  excelPreviewBtn.textContent = "正在生成预览...";
  excelPreviewBtn.disabled = true;
  excelRefreshPreviewBtn.disabled = true;
  const slowTimer = window.setTimeout(() => {
    excelPreviewStatus.textContent = "工作表较大，正在生成预览，请稍等。";
  }, 3000);

  try {
    const response = await fetch("/excel/watermark-preview", { method: "POST", body: formData });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "生成水印预览图失败");
    const preview = payload.preview;
    excelPreviewPlaceholder.hidden = true;
    excelPreviewImage.hidden = false;
    excelPreviewImage.dataset.mode = "fit";
    excelPreviewImage.src = `${preview.preview_url}?t=${Date.now()}`;
    excelPreviewStatus.textContent = preview.warnings && preview.warnings.length
      ? `预览已更新。${preview.warnings.join(" ")}`
      : "预览已更新。";
    excelPreviewFitBtn.disabled = false;
    excelPreviewOriginalBtn.disabled = false;
    excelPreviewZoomInBtn.disabled = false;
    excelPreviewZoomOutBtn.disabled = false;
  } catch (error) {
    excelPreviewImage.hidden = true;
    excelPreviewImage.removeAttribute("src");
    excelPreviewPlaceholder.hidden = false;
    excelPreviewStatus.textContent = error.message;
    showExcelErrors([{ file: "Excel 水印预览", error: error.message }]);
  } finally {
    window.clearTimeout(slowTimer);
    excelPreviewBtn.textContent = "生成水印预览图";
    updateExcelActionState();
  }
}

function setExcelPreviewFit() {
  if (excelPreviewImage.hidden) return;
  excelPreviewImage.dataset.mode = "fit";
  excelPreviewImage.style.maxWidth = "100%";
  excelPreviewImage.style.width = "100%";
  excelPreviewScale = 1;
}

function setExcelPreviewScale(scale) {
  if (excelPreviewImage.hidden || !excelPreviewImage.naturalWidth) return;
  excelPreviewImage.dataset.mode = "scale";
  excelPreviewScale = Math.max(0.4, Math.min(3, scale));
  excelPreviewImage.style.maxWidth = "none";
  excelPreviewImage.style.width = `${Math.round(excelPreviewImage.naturalWidth * excelPreviewScale)}px`;
  excelPreviewStage.scrollTop = 0;
}

function renderExcelResult(job) {
  const downloads = job.download_urls || {};
  const excelDownloadUrl = job.excel_download_url || downloads.excel || "";
  const reportDownloadUrl = job.report_download_url || downloads.report || "";
  const warnings = job.warnings && job.warnings.length ? `<div class="warning">${job.warnings.map(escapeHtml).join("<br>")}</div>` : "";
  const errorsHtml = job.errors && job.errors.length ? `<div class="errors">${job.errors.map(escapeHtml).join("<br>")}</div>` : "";
  const formulaNote = job.formula_conversion_note ? `<div class="job-meta">${escapeHtml(job.formula_conversion_note)}</div>` : "";
  const riskNotice = job.risk_notice ? `<div class="warning">${escapeHtml(job.risk_notice)}</div>` : "";
  const desktopMessage = job.desktop_output_message ? `<div class="job-meta">${escapeHtml(job.desktop_output_message)}</div>` : "";
  const actualSaveLocation = job.excel_desktop_path || job.excel_output_path || job.actual_save_location || job.app_output_path || "";
  const reportLocation = job.report_desktop_path || job.report_output_path || "程序 output 目录，可点击“下载处理报告 JSON”查看。";
  const formulaCells = job.preview_security_mode === "image_based"
    ? `<div>公式转数值：已跳过（图片化模式无需转值）</div>`
    : `
        <div>检测到公式：${job.formulas_detected_count || 0}</div>
        <div>成功转值公式：${job.formulas_converted_count || 0}</div>
      `;
  excelResults.innerHTML = `
    <section class="job">
      <div class="job-header">
        <div>
          <div class="job-title">${escapeHtml(job.source_file)}</div>
          <div class="job-meta">输出文件：${escapeHtml(job.output_excel)}</div>
          <div class="job-meta">当前版本：${escapeHtml(job.output_version || "v1")}</div>
          ${actualSaveLocation ? `<div class="job-meta">实际保存位置：${escapeHtml(actualSaveLocation)}</div>` : ""}
          <div class="job-meta">报告位置：${escapeHtml(reportLocation)}</div>
        </div>
        <div class="download-row">
          ${excelDownloadUrl ? `<a href="${excelDownloadUrl}" download>下载预览版 Excel</a>` : ""}
          ${reportDownloadUrl ? `<a href="${reportDownloadUrl}" download>下载处理报告 JSON</a>` : ""}
        </div>
      </div>
      <div class="excel-result-grid">
        <div>预览安全模式：${escapeHtml(securityModeLabel(job.preview_security_mode))}</div>
        <div>工作表总数：${job.total_sheets || 0}</div>
        <div>已处理工作表：${job.processed_sheets_count || 0}</div>
        <div>图片化工作表：${job.image_based_sheets_count || 0}</div>
        <div>跳过隐藏工作表：${job.skipped_hidden_sheets_count || 0}</div>
        ${formulaCells}
        <div>添加水印工作表：${job.watermark_sheets_count || 0}</div>
        <div>保护工作表：${job.protected_sheets_count || 0}</div>
        <div>保护工作簿结构：${job.workbook_structure_protected ? "是" : "否"}</div>
        <div>公式处理引擎：${escapeHtml(job.formula_conversion_engine || "-")}</div>
        <div>截图方式：${escapeHtml(screenshotEngineLabel(job.screenshot_engine || job.processing_engine || "-"))}</div>
        <div>处理引擎：${escapeHtml(job.processing_engine || "-")}</div>
        <div>图片插入自检：${job.image_insert_check_passed ? "通过" : "-"}</div>
        <div>插入图片：${job.inserted_images_count || 0} / ${job.expected_sheets_count || 0}</div>
        <div>预览/导出共用截图函数：${job.screenshot_function_shared ? "是" : "否"}</div>
        <div>输出位置：${escapeHtml(outputModeLabel(job.output_mode))}</div>
        <div>截图窗口模式：${escapeHtml(windowModeLabel(job.window_mode))}</div>
        <div>桌面 Excel：${job.desktop_excel_copied ? "已复制" : "未复制"}</div>
        <div>桌面报告：${job.desktop_report_copied ? "已复制" : "未复制"}</div>
        <div>打开密码：未设置</div>
        <div>编辑/保护：${job.edit_protection_enabled ? "已开启" : "未开启"}</div>
        <div>原始单元格数据：${job.real_cell_data_removed ? "已移除" : "仍保留"}</div>
        <div>复制风险：${escapeHtml(copyRiskLabel(job.copy_risk_level))}</div>
      </div>
      ${formulaNote}
      ${desktopMessage}
      ${riskNotice}
      ${warnings}
      ${errorsHtml}
    </section>
  `;
}

function showExcelErrors(items) {
  excelErrors.hidden = false;
  excelErrors.innerHTML = items
    .map((item) => `<div><strong>${escapeHtml(item.file)}：</strong>${escapeHtml(item.error)}</div>`)
    .join("");
}

async function processFiles() {
  if (!selectedFiles.length) return;
  statusText.textContent = "处理中...";
  processBtn.disabled = true;
  errors.hidden = true;
  errors.innerHTML = "";

  const formData = new FormData();
  selectedFiles.forEach((file) => formData.append("files", file));
  formData.append("image_method", defaultMaskMethod.value);
  if (expectedPages.value) {
    formData.append("expected_pages", expectedPages.value);
  }

  try {
    const response = await fetch("/api/process", { method: "POST", body: formData });
    const payload = await response.json();
    if (payload.jobs) {
      jobs = payload.jobs.concat(jobs);
      renderJobs();
    }
    if (payload.errors && payload.errors.length) {
      showErrors(payload.errors);
    }
    statusText.textContent = payload.jobs && payload.jobs.length ? "处理完成" : "处理未完成";
    if (response.ok || response.status === 207) {
      clearFiles();
    }
  } catch (error) {
    showErrors([{ file: "批量任务", error: error.message }]);
    statusText.textContent = "处理失败";
  } finally {
    processBtn.disabled = selectedFiles.length === 0;
  }
}

function renderJobs() {
  results.innerHTML = "";
  jobs.forEach((job) => {
    const section = document.createElement("section");
    section.className = "job";
    section.dataset.jobId = job.job_id;
    const downloads = job.download_urls || {};
    const pageCheck = job.page_count_check || {};
    const warning = pageCheck.page_count_warning || pageCheck.conversion_warning || "";
    section.innerHTML = `
      <div class="job-header">
        <div>
          <div class="job-title">${escapeHtml(job.original_file)}</div>
          <div class="job-meta" data-role="job-meta">${jobMetaHtml(job)}</div>
          <div class="conversion-meta">
            原文件类型：${escapeHtml(pageCheck.source_type || "-")}｜
            转换方式：${conversionLabel(pageCheck.conversion_method)}｜
            转换后页数：${pageCheck.converted_page_count || job.summary.page_count}｜
            页数异常：${pageCheck.page_count_anomaly ? "是" : "否"}
          </div>
        </div>
        <div class="download-row">
          <button type="button" data-export="pdf">导出脱敏 PDF</button>
          <button type="button" data-export="zip">导出全部详情图 ZIP</button>
          ${downloads.report ? `<a href="${downloads.report}" download>处理报告 JSON</a>` : ""}
        </div>
      </div>
      ${warning ? `<div class="${pageCheck.page_count_anomaly ? "conversion-warning" : "warning"}">${escapeHtml(warning)}</div>` : ""}
      <div class="job-progress" data-role="job-progress"></div>
      <div class="job-tools">
        <button type="button" data-action="reprocess">重新处理当前文件</button>
      </div>
      <div class="page-grid"></div>
    `;
    section.querySelector('[data-action="reprocess"]').addEventListener("click", () => runJobAction(job, "reprocess"));
    section.querySelector('[data-export="pdf"]').addEventListener("click", (event) => startExport(job, "pdf", event.currentTarget));
    section.querySelector('[data-export="zip"]').addEventListener("click", (event) => startExport(job, "zip", event.currentTarget));
    const grid = section.querySelector(".page-grid");
    (job.pages || []).forEach((page) => {
      const thumb = document.createElement("div");
      thumb.className = "thumb";
      thumb.dataset.page = String(page.page);
      thumb.innerHTML = `
        <button type="button">
          <img src="${page.image_url}?t=${Date.now()}" alt="第 ${page.page} 页">
          <span>第 ${page.page} 页</span>
        </button>
      `;
      thumb.querySelector("button").addEventListener("click", () => openPreview(job, page));
      grid.appendChild(thumb);
    });
    results.appendChild(section);
  });
}

function jobMetaHtml(job) {
  const summary = job.summary || {};
  return `
    共 ${summary.page_count || 0} 页｜
    自动识别 ${summary.auto_redactions || 0} 处｜
    脱敏 PDF：马赛克图片化｜
    详情图：${methodLabel(summary.detail_image_method)}｜
    人工补充：${summary.manual_redactions || 0} 处
  `;
}

async function runJobAction(job, action) {
  if (action !== "reprocess") return;
  const section = jobSection(job.job_id);
  const button = section?.querySelector('[data-action="reprocess"]');
  setButtonWorking(button, "正在生成...");
  setJobProgress(job.job_id, "正在重新处理当前文件...");
  const slowTimer = window.setTimeout(() => {
    setJobProgress(job.job_id, "文件页数较多，正在生成，请稍等。");
  }, 3000);
  try {
    const response = await fetch("/api/reprocess", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ job_id: job.job_id, image_method: defaultMaskMethod.value }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "操作失败");
    replaceJob(data.job);
    renderJobs();
    statusText.textContent = "处理完成";
  } catch (error) {
    showErrors([{ file: job.original_file, error: error.message }]);
    statusText.textContent = "处理失败";
  } finally {
    window.clearTimeout(slowTimer);
    restoreButton(button);
  }
}

async function startExport(job, kind, button) {
  const label = kind === "pdf" ? "导出脱敏 PDF" : "导出全部详情图 ZIP";
  setButtonWorking(button, "正在生成...");
  setJobProgress(job.job_id, kind === "pdf" ? "正在生成脱敏 PDF" : "正在生成分页详情图");
  const slowTimer = window.setTimeout(() => {
    setJobProgress(job.job_id, "文件页数较多，正在生成，请稍等。");
  }, 3000);
  try {
    const response = await fetch("/api/export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ job_id: job.job_id, kind }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "导出失败");
    await pollExportTask(data.task_id, job, button, label, slowTimer);
  } catch (error) {
    window.clearTimeout(slowTimer);
    restoreButton(button, label);
    setJobProgress(job.job_id, "");
    showErrors([{ file: job.original_file, error: error.message }]);
  }
}

async function pollExportTask(taskId, job, button, label, slowTimer) {
  while (true) {
    const response = await fetch(`/api/export-status/${encodeURIComponent(taskId)}`);
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "读取导出进度失败");
    const progressText = formatProgress(data);
    setJobProgress(job.job_id, progressText);
    if (data.state === "done") {
      window.clearTimeout(slowTimer);
      restoreButton(button, label);
      if (data.job) {
        replaceJob(data.job);
        refreshJobCard(data.job);
      }
      setJobProgress(job.job_id, "生成完成，正在下载...");
      if (data.download_url) {
        window.location.href = data.download_url;
      }
      return;
    }
    if (data.state === "error") {
      throw new Error(data.message || "导出失败");
    }
    await wait(600);
  }
}

function openPreview(job, page) {
  if (!job || !page) {
    showErrors([{ file: "页面预览", error: "当前页信息不存在，请重新打开预览页。" }]);
    return;
  }
  activeJob = job;
  activePage = page;
  selection = null;
  dragStart = null;
  setManualStatus("");
  previewTitle.textContent = `${job.original_file} - 第 ${page.page} 页`;
  previewImage.src = `${page.image_url}?t=${Date.now()}`;
  modal.hidden = false;
}

function closePreview() {
  modal.hidden = true;
  activeJob = null;
  activePage = null;
  selection = null;
  setManualStatus("");
  drawSelection();
}

function syncCanvas() {
  if (modal.hidden || !previewImage.complete) return;
  canvas.width = previewImage.clientWidth;
  canvas.height = previewImage.clientHeight;
  canvas.style.width = `${previewImage.clientWidth}px`;
  canvas.style.height = `${previewImage.clientHeight}px`;
  drawSelection();
}

function drawSelection() {
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  if (!selection) return;
  ctx.fillStyle = "rgba(31, 122, 104, 0.18)";
  ctx.strokeStyle = "#1f7a68";
  ctx.lineWidth = 2;
  ctx.fillRect(selection.x, selection.y, selection.width, selection.height);
  ctx.strokeRect(selection.x, selection.y, selection.width, selection.height);
}

async function submitManualRedaction() {
  if (!activeJob || !activePage) {
    setManualStatus("当前页信息不存在，请重新打开预览页。", true);
    return;
  }
  if (!selection || selection.width < 6 || selection.height < 6) {
    setManualStatus("请先在页面上拖拽框选需要遮挡的区域。", true);
    return;
  }
  if (!previewImage.naturalWidth || !previewImage.naturalHeight || !previewImage.clientWidth || !previewImage.clientHeight) {
    setManualStatus("当前页信息不存在，请重新打开预览页。", true);
    return;
  }

  submitManualBtn.disabled = true;
  setManualStatus("正在保存手动遮挡...");
  const scaleX = previewImage.naturalWidth / previewImage.clientWidth;
  const scaleY = previewImage.naturalHeight / previewImage.clientHeight;
  const payload = {
    job_id: activeJob.job_id,
    page: activePage.page,
    x: selection.x * scaleX,
    y: selection.y * scaleY,
    width: selection.width * scaleX,
    height: selection.height * scaleY,
    image_width: previewImage.naturalWidth,
    image_height: previewImage.naturalHeight,
    method: maskMethod.value || "mosaic",
  };

  try {
    const response = await fetch("/api/manual-redaction", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "人工遮挡失败");
    const pageNumber = activePage.page;
    replaceJob(data.job);
    activeJob = findJob(data.job.job_id);
    activePage = activeJob?.pages?.find((page) => page.page === pageNumber) || null;
    refreshJobCard(data.job, pageNumber);
    reloadActivePreview(pageNumber);
    selection = null;
    drawSelection();
    setManualStatus("已添加 1 处手动遮挡，导出时会自动应用。");
  } catch (error) {
    setManualStatus(error.message, true);
  } finally {
    submitManualBtn.disabled = false;
  }
}

async function undoPageManualRedaction() {
  await runPageManualAction("/api/manual-redaction/undo", "已撤销当前页上一处手动遮挡。");
}

async function clearPageManualRedactions() {
  await runPageManualAction("/api/manual-redactions/clear", "已清空当前页手动遮挡。");
}

async function runPageManualAction(url, successMessage) {
  if (!activeJob || !activePage) {
    setManualStatus("当前页信息不存在，请重新打开预览页。", true);
    return;
  }
  const pageNumber = activePage.page;
  setManualStatus("正在更新当前页...");
  try {
    const response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ job_id: activeJob.job_id, page: pageNumber }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "操作失败");
    replaceJob(data.job);
    activeJob = findJob(data.job.job_id);
    activePage = activeJob?.pages?.find((page) => page.page === pageNumber) || null;
    refreshJobCard(data.job, pageNumber);
    reloadActivePreview(pageNumber);
    selection = null;
    drawSelection();
    setManualStatus(successMessage);
  } catch (error) {
    setManualStatus(error.message, true);
  }
}

function reloadActivePreview(pageNumber) {
  if (!activeJob || !activePage) {
    setManualStatus("当前页信息不存在，请重新打开预览页。", true);
    return;
  }
  previewTitle.textContent = `${activeJob.original_file} - 第 ${pageNumber} 页`;
  previewImage.src = `${activePage.image_url}?t=${Date.now()}`;
}

function refreshJobCard(job, pageNumber) {
  const section = jobSection(job.job_id);
  if (!section) return;
  const meta = section.querySelector('[data-role="job-meta"]');
  if (meta) meta.innerHTML = jobMetaHtml(job);
  if (pageNumber) {
    const page = (job.pages || []).find((item) => item.page === pageNumber);
    const image = section.querySelector(`.thumb[data-page="${pageNumber}"] img`);
    if (page && image) image.src = `${page.image_url}?t=${Date.now()}`;
  }
}

function replaceJob(updatedJob) {
  jobs = jobs.map((job) => (job.job_id === updatedJob.job_id ? updatedJob : job));
}

function findJob(jobId) {
  return jobs.find((job) => job.job_id === jobId) || null;
}

function jobSection(jobId) {
  return results.querySelector(`.job[data-job-id="${cssEscape(jobId)}"]`);
}

function setJobProgress(jobId, message) {
  const section = jobSection(jobId);
  const progress = section?.querySelector('[data-role="job-progress"]');
  if (progress) progress.textContent = message || "";
}

function setManualStatus(message, isError = false) {
  manualStatus.textContent = message;
  manualStatus.hidden = !message;
  manualStatus.classList.toggle("is-error", Boolean(isError));
}

function setButtonWorking(button, text) {
  if (!button) return;
  button.dataset.originalText = button.textContent;
  button.textContent = text;
  button.disabled = true;
}

function restoreButton(button, fallbackText) {
  if (!button) return;
  button.textContent = fallbackText || button.dataset.originalText || button.textContent;
  button.disabled = false;
}

function formatProgress(task) {
  if (!task) return "";
  if (task.current && task.total) {
    return `${task.message || "正在处理"}（${task.current} / ${task.total} 页）`;
  }
  return task.message || "";
}

function methodLabel(method) {
  return {
    black_box: "黑块",
    mosaic: "马赛克",
    blur: "模糊",
    gray_mosaic: "浅灰马赛克",
  }[method] || "马赛克";
}

function securityModeLabel(mode) {
  return {
    image_based: "图片化防复制版 Excel",
    locked_excel: "普通锁定版 Excel",
  }[mode] || "图片化防复制版 Excel";
}

function copyRiskLabel(level) {
  return {
    low: "低",
    medium: "中",
    high: "高",
  }[level] || "-";
}

function outputModeLabel(mode) {
  return {
    desktop: "输出到桌面",
    app_output: "程序 output 目录",
  }[mode] || "-";
}

function windowModeLabel(mode) {
  return {
    quiet: "安静模式",
    visible: "可见调试模式",
  }[mode] || "-";
}

function screenshotEngineLabel(engine) {
  return {
    "Excel COM": "Excel COM 真实截图",
    "WPS COM": "WPS COM 真实截图",
    fallback_openpyxl: "fallback 近似绘制",
    not_used: "未使用",
  }[engine] || engine;
}

function conversionLabel(method) {
  return {
    direct_pdf: "PDF 直接处理",
    libreoffice_auto: "LibreOffice 自动转换",
  }[method] || "-";
}

function pointerPosition(event) {
  const rect = canvas.getBoundingClientRect();
  return {
    x: Math.max(0, Math.min(canvas.width, event.clientX - rect.left)),
    y: Math.max(0, Math.min(canvas.height, event.clientY - rect.top)),
  };
}

function normalizeRect(x, y, width, height) {
  const nx = width < 0 ? x + width : x;
  const ny = height < 0 ? y + height : y;
  return {
    x: Math.max(0, nx),
    y: Math.max(0, ny),
    width: Math.abs(width),
    height: Math.abs(height),
  };
}

function showErrors(items) {
  errors.hidden = false;
  errors.innerHTML = items
    .map((item) => `<div><strong>${escapeHtml(item.file)}：</strong>${escapeHtml(item.error)}</div>`)
    .join("");
}

function formatSize(size) {
  if (size > 1024 * 1024) return `${(size / 1024 / 1024).toFixed(1)} MB`;
  if (size > 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${size} B`;
}

function cssEscape(value) {
  if (window.CSS && typeof window.CSS.escape === "function") {
    return window.CSS.escape(String(value));
  }
  return String(value).replaceAll('"', '\\"');
}

function wait(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
