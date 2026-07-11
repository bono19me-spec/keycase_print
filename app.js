const DEFAULT_SETTINGS = {
  pageWidthMm: 130,
  pageHeightMm: 110,
  printAreaStartX: 65,
  nameX: 98,
  nameY: 70,
  nameAlign: "left",
  nameFontSize: 10,
  roomX: 98,
  roomY: 88,
  roomAlign: "left",
  roomFontSize: 18,
  printName: true,
  printNameHonorific: true,
  printRoom: true,
  nameLineGap: 6,
  groupNameX: 98,
  groupNameY: 14,
  groupNameFontSize: 9,
  stayInfoX: 98,
  stayInfoY: 52,
  stayInfoFontSize: 9,
  printGroupName: false,
  printStayInfo: false,
  printCopyMode: "room",
  globalOffsetX: 0,
  globalOffsetY: 0,
  rotate180: false,
  dataStartRow: 13,
  roomColumn: "U",
  nameColumn: "X"
};

const STORAGE_KEY = "keycoverPrintSettings.v2.b6";
const SHEET_NAME = "団体メンバ一覧表";
const BUILD_VERSION = "20260712-customfields1";
const B6_WIDTH_MM = 182;
const B6_HEIGHT_MM = 128;
const NAME_AREA_WIDTH_MM = 58;
const MIN_NAME_FONT_SIZE_PT = 6.5;
const PT_TO_MM = 0.352778;

let settings = loadSettings();
let records = [];
let warnings = [];
let currentWorkbook = null;
let simpleStep = 1;

const els = {
  modeHome: document.getElementById("modeHome"),
  simpleApp: document.getElementById("simpleApp"),
  advancedApp: document.getElementById("advancedApp"),
  homeBuildVersion: document.getElementById("homeBuildVersion"),
  excelFile: document.getElementById("excelFile"),
  fileName: document.getElementById("fileName"),
  sheetSelect: document.getElementById("sheetSelect"),
  simpleExcelFile: document.getElementById("simpleExcelFile"),
  simpleFileName: document.getElementById("simpleFileName"),
  simpleSheetSelect: document.getElementById("simpleSheetSelect"),
  simpleStatus: document.getElementById("simpleStatus"),
  simpleRangeStart: document.getElementById("simpleRangeStart"),
  simpleRangeEnd: document.getElementById("simpleRangeEnd"),
  simpleCopyModePanel: document.getElementById("simpleCopyModePanel"),
  simpleCopyModeHint: document.getElementById("simpleCopyModeHint"),
  advancedCopyModePanel: document.getElementById("advancedCopyModePanel"),
  advancedCopyModeHint: document.getElementById("advancedCopyModeHint"),
  status: document.getElementById("status"),
  dataTable: document.getElementById("dataTable"),
  countBadge: document.getElementById("countBadge"),
  warningBadge: document.getElementById("warningBadge"),
  customGroupName: document.getElementById("customGroupName"),
  customStayInfo: document.getElementById("customStayInfo"),
  customEntries: document.getElementById("customEntries"),
  simpleCustomGroupName: document.getElementById("simpleCustomGroupName"),
  simpleCustomStayInfo: document.getElementById("simpleCustomStayInfo"),
  simpleCustomEntries: document.getElementById("simpleCustomEntries"),
  rangeStart: document.getElementById("rangeStart"),
  rangeEnd: document.getElementById("rangeEnd"),
  buildVersion: document.getElementById("buildVersion")
};

const settingInputs = [
  "globalOffsetX",
  "globalOffsetY",
  "nameX",
  "nameY",
  "roomX",
  "roomY",
  "nameFontSize",
  "roomFontSize",
  "printName",
  "printNameHonorific",
  "printRoom",
  "printGroupName",
  "printStayInfo",
  "rotate180"
];

const simplePositionInputMap = {
  simpleGlobalOffsetX: "globalOffsetX",
  simpleGlobalOffsetY: "globalOffsetY",
  simpleNameX: "nameX",
  simpleNameY: "nameY",
  simpleRoomX: "roomX",
  simpleRoomY: "roomY",
  simpleNameFontSize: "nameFontSize",
  simpleRoomFontSize: "roomFontSize",
  simpleRotate180: "rotate180"
};

init();

function init() {
  els.homeBuildVersion.textContent = `Build ${BUILD_VERSION}`;
  els.buildVersion.textContent = `Build ${BUILD_VERSION}`;
  bindSettingsToForm();
  bindSimpleSettingsToForm();
  bindCustomSettingsToForm();
  bindSimplePositionSettingsToForm();
  showMode("home");
  showSimpleStep(1);

  document.getElementById("startSimple").addEventListener("click", () => showMode("simple"));
  document.getElementById("startAdvanced").addEventListener("click", () => showMode("advanced"));
  document.getElementById("simpleBackHome").addEventListener("click", () => showMode("home"));
  document.getElementById("advancedBackHome").addEventListener("click", () => showMode("home"));
  document.getElementById("simpleNextUpload").addEventListener("click", () => showSimpleStep(2));
  document.getElementById("simplePrevInfo").addEventListener("click", () => showSimpleStep(1));
  document.getElementById("simpleNextInfo").addEventListener("click", () => {
    updateSettingsFromSimpleForm();
    if (!validateRequiredPrintField()) return;
    bindSettingsToForm();
    showSimpleStep(3);
  });
  document.getElementById("simplePrevPrinter").addEventListener("click", () => showSimpleStep(2));
  document.getElementById("simpleNextPrinter").addEventListener("click", () => showSimpleStep(4));
  document.getElementById("simplePrevPrint").addEventListener("click", () => showSimpleStep(3));

  settingInputs.forEach((key) => {
    document.getElementById(key).addEventListener("input", () => {
      updateSettingsFromForm();
      bindSimpleSettingsToForm();
      bindCustomSettingsToForm();
      bindSimplePositionSettingsToForm();
      renderTable();
    });
  });
  ["simplePrintName", "simplePrintNameHonorific", "simplePrintRoom", "simplePrintGroupName", "simplePrintStayInfo"].forEach((id) => {
    document.getElementById(id).addEventListener("input", () => {
      updateSettingsFromSimpleForm();
      bindSettingsToForm();
      bindCustomSettingsToForm();
      updatePrintFieldControls();
      renderTable();
    });
  });
  [
    "simpleCustomPrintName",
    "simpleCustomPrintNameHonorific",
    "simpleCustomPrintRoom",
    "customPrintName",
    "customPrintNameHonorific",
    "customPrintRoom"
  ].forEach((id) => {
    document.getElementById(id).addEventListener("input", () => {
      updateSettingsFromCustomForm(id.startsWith("simple"));
      bindSettingsToForm();
      bindSimpleSettingsToForm();
      bindCustomSettingsToForm();
      renderTable();
    });
  });
  document.querySelectorAll('input[name="simpleCopyMode"], input[name="advancedCopyMode"]').forEach((input) => {
    input.addEventListener("input", () => {
      updateCopyModeFromForm(input.value);
      bindCopyModeToForm();
    });
  });
  Object.keys(simplePositionInputMap).forEach((id) => {
    document.getElementById(id).addEventListener("input", () => {
      updateSettingsFromSimplePositionForm();
      bindSettingsToForm();
    });
  });
  document.getElementById("simpleSaveSettings").addEventListener("click", saveSimpleSettings);
  document.getElementById("simpleResetSettings").addEventListener("click", resetSimpleSettings);

  els.excelFile.addEventListener("change", handleFile);
  els.sheetSelect.addEventListener("change", parseSelectedSheet);
  els.simpleExcelFile.addEventListener("change", handleFile);
  els.simpleSheetSelect.addEventListener("change", parseSelectedSheet);
  document.getElementById("loadCustom").addEventListener("click", loadCustomEntries);
  document.getElementById("simpleLoadCustom").addEventListener("click", loadSimpleCustomEntries);
  document.getElementById("saveSettings").addEventListener("click", saveSettings);
  document.getElementById("resetSettings").addEventListener("click", resetSettings);
  document.getElementById("testPrint").addEventListener("click", () => openPrintWindow("test"));
  document.getElementById("rangePrint").addEventListener("click", () => openPrintWindow("range"));
  document.getElementById("allPrint").addEventListener("click", () => openPrintWindow("all"));
  document.getElementById("simpleTestPrint").addEventListener("click", () => openPrintWindow("test"));
  document.getElementById("simpleRangePrint").addEventListener("click", () => openPrintWindow("range"));
  document.getElementById("simpleAllPrint").addEventListener("click", () => openPrintWindow("all"));
  [els.rangeStart, els.rangeEnd, els.simpleRangeStart, els.simpleRangeEnd].forEach((input) => {
    input.addEventListener("input", () => syncRangeInputs(false, input));
    input.addEventListener("change", () => syncRangeInputs(false, input));
  });
  syncRangeInputs(true);
  updatePrintFieldControls();
  updateCustomEntryHints();
}

function showMode(mode) {
  els.modeHome.hidden = mode !== "home";
  els.simpleApp.hidden = mode !== "simple";
  els.advancedApp.hidden = mode !== "advanced";
  updateCopyModeVisibility();
}

function showSimpleStep(step) {
  simpleStep = step;
  document.querySelectorAll("[data-simple-step]").forEach((section) => {
    section.hidden = Number(section.dataset.simpleStep) !== step;
  });
  document.querySelectorAll("[data-step-dot]").forEach((dot) => {
    const dotStep = Number(dot.dataset.stepDot);
    dot.classList.toggle("active", dotStep === step);
    dot.classList.toggle("done", dotStep < step);
  });
}

function loadSettings() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
    return { ...DEFAULT_SETTINGS, ...(saved || {}) };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

function bindSettingsToForm() {
  settingInputs.forEach((key) => {
    const input = document.getElementById(key);
    if (input.type === "checkbox") {
      input.checked = Boolean(settings[key]);
    } else {
      input.value = settings[key];
    }
  });
  updatePrintFieldControls();
}

function bindSimpleSettingsToForm() {
  document.getElementById("simplePrintName").checked = Boolean(settings.printName);
  document.getElementById("simplePrintNameHonorific").checked = Boolean(settings.printNameHonorific);
  document.getElementById("simplePrintRoom").checked = Boolean(settings.printRoom);
  document.getElementById("simplePrintGroupName").checked = Boolean(settings.printGroupName);
  document.getElementById("simplePrintStayInfo").checked = Boolean(settings.printStayInfo);
  bindCopyModeToForm();
  updatePrintFieldControls();
}

function bindCustomSettingsToForm() {
  [
    ["simpleCustomPrintName", "printName"],
    ["simpleCustomPrintNameHonorific", "printNameHonorific"],
    ["simpleCustomPrintRoom", "printRoom"],
    ["customPrintName", "printName"],
    ["customPrintNameHonorific", "printNameHonorific"],
    ["customPrintRoom", "printRoom"]
  ].forEach(([id, key]) => {
    document.getElementById(id).checked = Boolean(settings[key]);
  });
  updateCustomEntryHints();
}

function bindSimplePositionSettingsToForm() {
  Object.entries(simplePositionInputMap).forEach(([id, key]) => {
    const input = document.getElementById(id);
    if (input.type === "checkbox") {
      input.checked = Boolean(settings[key]);
    } else {
      input.value = settings[key];
    }
  });
}

function bindCopyModeToForm() {
  const mode = settings.printCopyMode === "guest" ? "guest" : "room";
  document.querySelectorAll('input[name="simpleCopyMode"], input[name="advancedCopyMode"]').forEach((input) => {
    input.checked = input.value === mode;
  });
}

function updateCopyModeFromForm(value) {
  settings.printCopyMode = value === "guest" ? "guest" : "room";
}

function updateCopyModeFromActiveForm() {
  const selector = els.simpleApp.hidden
    ? 'input[name="advancedCopyMode"]:checked'
    : 'input[name="simpleCopyMode"]:checked';
  const selected = document.querySelector(selector);
  if (selected) updateCopyModeFromForm(selected.value);
}

function getMultiOccupancySummary() {
  const multiRooms = records.filter((record) => record.outputNames.length > 1);
  const extraCards = multiRooms.reduce((sum, record) => sum + record.outputNames.length - 1, 0);
  return { multiRooms, extraCards };
}

function updateCopyModeVisibility() {
  const { multiRooms, extraCards } = getMultiOccupancySummary();
  const hasMultiRooms = multiRooms.length > 0;
  const hint = hasMultiRooms
    ? `2名以上の部屋が${multiRooms.length}室あります。人数分の場合は追加で${extraCards}枚印刷されます。`
    : "";

  [els.simpleCopyModePanel, els.advancedCopyModePanel].forEach((panel) => {
    panel.hidden = !hasMultiRooms;
  });
  els.simpleCopyModeHint.textContent = hint;
  els.advancedCopyModeHint.textContent = hint;

  if (!hasMultiRooms) {
    settings.printCopyMode = "room";
    bindCopyModeToForm();
  }
}

function syncRangeInputs(resetToFullRange = false, changedInput = null) {
  const max = Math.max(1, records.length);
  const disabled = !records.length;

  [els.rangeStart, els.simpleRangeStart].forEach((input) => {
    input.min = 1;
    input.max = max;
    input.disabled = disabled;
  });
  [els.rangeEnd, els.simpleRangeEnd].forEach((input) => {
    input.min = 1;
    input.max = max;
    input.disabled = disabled;
  });

  if (resetToFullRange || disabled) {
    [els.rangeStart, els.simpleRangeStart].forEach((input) => {
      input.value = 1;
    });
    [els.rangeEnd, els.simpleRangeEnd].forEach((input) => {
      input.value = max;
    });
  }

  syncRangePair(els.rangeStart, els.rangeEnd, max, changedInput);
  syncRangePair(els.simpleRangeStart, els.simpleRangeEnd, max, changedInput);
}

function syncRangePair(startInput, endInput, max, changedInput) {
  let start = clampRangeValue(startInput.value, max);
  let end = clampRangeValue(endInput.value, max);

  if (start > end) {
    if (changedInput === startInput) {
      end = start;
    } else {
      start = end;
    }
  }

  startInput.value = start;
  endInput.value = end;
}

function clampRangeValue(value, max) {
  const numeric = Number(value || 1);
  if (!Number.isFinite(numeric)) return 1;
  return Math.min(max, Math.max(1, Math.trunc(numeric)));
}

function updateSettingsFromForm() {
  settingInputs.forEach((key) => {
    const input = document.getElementById(key);
    settings[key] = input.type === "checkbox" ? input.checked : Number(input.value);
  });
  const selected = document.querySelector('input[name="advancedCopyMode"]:checked');
  if (selected) settings.printCopyMode = selected.value;
  updatePrintFieldControls();
}

function updateSettingsFromSimpleForm() {
  settings.printName = document.getElementById("simplePrintName").checked;
  settings.printNameHonorific = document.getElementById("simplePrintNameHonorific").checked;
  settings.printRoom = document.getElementById("simplePrintRoom").checked;
  settings.printGroupName = document.getElementById("simplePrintGroupName").checked;
  settings.printStayInfo = document.getElementById("simplePrintStayInfo").checked;
  const selected = document.querySelector('input[name="simpleCopyMode"]:checked');
  if (selected) settings.printCopyMode = selected.value;
  updatePrintFieldControls();
}

function updateSettingsFromCustomForm(isSimple) {
  const prefix = isSimple ? "simpleCustom" : "custom";
  settings.printName = document.getElementById(`${prefix}PrintName`).checked;
  settings.printNameHonorific = document.getElementById(`${prefix}PrintNameHonorific`).checked;
  settings.printRoom = document.getElementById(`${prefix}PrintRoom`).checked;
  updatePrintFieldControls();
  updateCustomEntryHints();
  return {
    printName: settings.printName,
    printNameHonorific: settings.printNameHonorific,
    printRoom: settings.printRoom
  };
}

function hasRequiredPrintField() {
  return Boolean(settings.printName || settings.printRoom);
}

function updatePrintFieldControls() {
  const simpleHonorific = document.getElementById("simplePrintNameHonorific");
  const advancedHonorific = document.getElementById("printNameHonorific");
  const simpleCustomHonorific = document.getElementById("simpleCustomPrintNameHonorific");
  const customHonorific = document.getElementById("customPrintNameHonorific");
  const simpleError = document.getElementById("simplePrintFieldError");
  const advancedError = document.getElementById("advancedPrintFieldError");
  const simpleCustomError = document.getElementById("simpleCustomFieldError");
  const customError = document.getElementById("customFieldError");
  const simpleNext = document.getElementById("simpleNextInfo");
  const isValid = hasRequiredPrintField();

  [simpleHonorific, advancedHonorific, simpleCustomHonorific, customHonorific].forEach((input) => {
    if (!input) return;
    input.disabled = !settings.printName;
    input.closest(".checkbox-row")?.classList.toggle("is-disabled", !settings.printName);
  });
  if (simpleError) simpleError.hidden = isValid;
  if (advancedError) advancedError.hidden = isValid;
  if (simpleCustomError) simpleCustomError.hidden = isValid;
  if (customError) customError.hidden = isValid;
  if (simpleNext) simpleNext.disabled = !isValid;
}

function updateCustomEntryHints() {
  const label = getCustomEntriesLabel(settings);
  const placeholder = getCustomEntriesPlaceholder(settings);
  [
    ["simpleCustomEntriesLabel", "simpleCustomEntries"],
    ["customEntriesLabel", "customEntries"]
  ].forEach(([labelId, textareaId]) => {
    const labelEl = document.getElementById(labelId);
    const textarea = document.getElementById(textareaId);
    if (labelEl) labelEl.textContent = label;
    if (textarea) textarea.placeholder = placeholder;
  });
}

function getCustomEntriesLabel(fields) {
  if (fields.printName && fields.printRoom) return "部屋番号 / 氏名";
  if (fields.printRoom) return "部屋番号";
  if (fields.printName) return "氏名";
  return "入力内容";
}

function getCustomEntriesPlaceholder(fields) {
  if (fields.printName && fields.printRoom) {
    return "例:\n4082 イノウエ　トモオ/カワハラ　リッカ\n4083 カワカミ　ミコト";
  }
  if (fields.printRoom) {
    return "例:\n4082\n4083";
  }
  if (fields.printName) {
    return "例:\nイノウエ　トモオ/カワハラ　リッカ\nカワカミ　ミコト";
  }
  return "氏名または部屋番号を選択してください。";
}

function updateSettingsFromSimplePositionForm() {
  Object.entries(simplePositionInputMap).forEach(([id, key]) => {
    const input = document.getElementById(id);
    settings[key] = input.type === "checkbox" ? input.checked : Number(input.value);
  });
}

function validateRequiredPrintField() {
  updatePrintFieldControls();
  if (hasRequiredPrintField()) return true;

  const message = "氏名または部屋番号を少なくとも1つ選択してください。";
  setStatus(message, true);
  setSimpleStatus(message, true);
  return false;
}

function saveSettings() {
  updateSettingsFromForm();
  updateCopyModeFromActiveForm();
  bindSimplePositionSettingsToForm();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(pickSettings(settings)));
  setStatus("設定をブラウザに保存しました。");
}

function saveSimpleSettings() {
  updateSettingsFromSimpleForm();
  updateSettingsFromSimplePositionForm();
  bindSettingsToForm();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(pickSettings(settings)));
  setSimpleStatus("位置設定をブラウザに保存しました。");
}

function resetSettings() {
  settings = { ...DEFAULT_SETTINGS };
  bindSettingsToForm();
  bindSimpleSettingsToForm();
  bindCustomSettingsToForm();
  bindSimplePositionSettingsToForm();
  localStorage.removeItem(STORAGE_KEY);
  setStatus("初期位置に戻しました。");
}

function resetSimpleSettings() {
  settings = { ...DEFAULT_SETTINGS };
  bindSettingsToForm();
  bindSimpleSettingsToForm();
  bindCustomSettingsToForm();
  bindSimplePositionSettingsToForm();
  localStorage.removeItem(STORAGE_KEY);
  setSimpleStatus("位置設定を初期値に戻しました。");
}

function pickSettings(source) {
  return settingInputs.reduce((picked, key) => {
    picked[key] = source[key];
    return picked;
  }, { printCopyMode: source.printCopyMode });
}

async function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  els.fileName.textContent = file.name;
  els.simpleFileName.textContent = file.name;
  setStatus("Excelファイルを読み込んでいます。");
  setSimpleStatus("Excelファイルを読み込んでいます。");

  try {
    const buffer = await file.arrayBuffer();
    currentWorkbook = XLSX.read(buffer, { type: "array", cellDates: false });
    populateSheetSelect(currentWorkbook);
    parseSelectedSheet();
  } catch (error) {
    currentWorkbook = null;
    populateSheetSelect(null);
    records = [];
    warnings = [];
    renderTable();
    setStatus(error.message || "ファイルを読み込めません。", true);
    setSimpleStatus(error.message || "ファイルを読み込めません。", true);
  }
}

function populateSheetSelect(workbook) {
  const sheetNames = workbook?.SheetNames || [];
  [els.sheetSelect, els.simpleSheetSelect].forEach((select) => {
    select.disabled = !sheetNames.length;
    select.innerHTML = sheetNames.length
    ? sheetNames.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("")
    : '<option value="">Excelファイルを選択してください</option>';

    if (sheetNames.length) {
      select.value = sheetNames.includes(SHEET_NAME) ? SHEET_NAME : sheetNames[0];
    }
  });
}

function parseSelectedSheet() {
  if (!currentWorkbook) return;

  try {
    const activeSelect = els.simpleApp.hidden ? els.sheetSelect : els.simpleSheetSelect;
    const sheetName = activeSelect.value || currentWorkbook.SheetNames[0];
    els.sheetSelect.value = sheetName;
    els.simpleSheetSelect.value = sheetName;
    const sheet = currentWorkbook.Sheets[sheetName];

    if (!sheet) {
      throw new Error("選択したシートを読み込めません。");
    }

    const parsed = extractRecords(sheet);
    records = parsed.valid;
    warnings = parsed.warnings;

    if (!records.length) {
      throw new Error("13行目以降に有効な部屋番号・宿泊者名データがありません。");
    }

    renderTable();
    syncRangeInputs(true);
    document.getElementById("simpleNextUpload").disabled = false;
    const summary = getRecordsSummary(records);
    setStatus(`「${sheetName}」から ${summary}のデータを読み込みました。${warnings.length ? `除外・警告 ${warnings.length}件があります。` : ""}`);
    setSimpleStatus(`「${sheetName}」から ${summary}のデータを読み込みました。${warnings.length ? `除外・警告 ${warnings.length}件があります。` : ""}`);
  } catch (error) {
    records = [];
    warnings = [];
    renderTable();
    syncRangeInputs(true);
    document.getElementById("simpleNextUpload").disabled = true;
    setStatus(error.message || "シートを読み込めません。", true);
    setSimpleStatus(error.message || "シートを読み込めません。", true);
  }
}

function loadCustomEntries() {
  const fields = updateSettingsFromCustomForm(false);
  if (!validateRequiredPrintField()) return;
  const text = els.customEntries.value.trim();
  if (!text) {
    setStatus(`直接入力欄に${getCustomEntriesLabel(fields)}を入力してください。`, true);
    return;
  }

  bindSimpleSettingsToForm();
  const groupName = els.customGroupName.value.trim();
  const stayInfo = els.customStayInfo.value.trim();
  const parsed = parseCustomEntries(text, groupName, stayInfo, fields);
  records = parsed.valid;
  warnings = parsed.warnings;
  currentWorkbook = null;
  populateSheetSelect(null);

  if (!records.length) {
    renderTable();
    syncRangeInputs(true);
    setStatus("有効な直接入力データがありません。", true);
    return;
  }

  renderTable();
  syncRangeInputs(true);
  document.getElementById("simpleNextUpload").disabled = false;
  const summary = getRecordsSummary(records);
  setStatus(`直接入力から ${summary}のデータを読み込みました。${warnings.length ? `警告 ${warnings.length}件があります。` : ""}`);
  setSimpleStatus(`直接入力から ${summary}のデータを読み込みました。${warnings.length ? `警告 ${warnings.length}件があります。` : ""}`);
}

function loadSimpleCustomEntries() {
  const fields = updateSettingsFromCustomForm(true);
  if (!validateRequiredPrintField()) return;
  const text = els.simpleCustomEntries.value.trim();
  if (!text) {
    setSimpleStatus(`直接入力欄に${getCustomEntriesLabel(fields)}を入力してください。`, true);
    return;
  }

  bindSettingsToForm();
  bindSimpleSettingsToForm();
  const groupName = els.simpleCustomGroupName.value.trim();
  const stayInfo = els.simpleCustomStayInfo.value.trim();
  const parsed = parseCustomEntries(text, groupName, stayInfo, fields);
  records = parsed.valid;
  warnings = parsed.warnings;
  currentWorkbook = null;
  populateSheetSelect(null);

  if (!records.length) {
    renderTable();
    syncRangeInputs(true);
    setSimpleStatus("有効な直接入力データがありません。", true);
    return;
  }

  renderTable();
  syncRangeInputs(true);
  document.getElementById("simpleNextUpload").disabled = false;
  const summary = getRecordsSummary(records);
  setStatus(`直接入力から ${summary}のデータを読み込みました。${warnings.length ? `警告 ${warnings.length}件があります。` : ""}`);
  setSimpleStatus(`直接入力から ${summary}のデータを読み込みました。${warnings.length ? `警告 ${warnings.length}件があります。` : ""}`);
}

function parseCustomEntries(text, groupName, stayInfo, fields = DEFAULT_SETTINGS) {
  const valid = [];
  const rowWarnings = [];
  const needsRoom = Boolean(fields.printRoom);
  const needsName = Boolean(fields.printName);

  text.split(/\r?\n/).forEach((line, lineIndex) => {
    const trimmed = line.trim();
    if (!trimmed) return;

    const parsed = parseCustomLineForFields(trimmed, { needsRoom, needsName });
    if ((needsRoom && !parsed.room) || (needsName && !parsed.names.length)) {
      rowWarnings.push({
        excelRow: lineIndex + 1,
        room: parsed.room || "",
        rawName: parsed.rawNames || trimmed,
        reason: needsRoom && !parsed.room ? "部屋番号が空です" : "氏名が空です"
      });
      return;
    }

    const rawNames = parsed.names;
    const outputNames = rawNames.map(normalizeGuestName);
    valid.push({
      index: valid.length + 1,
      excelRow: `入力${lineIndex + 1}`,
      room: parsed.room,
      rawNames,
      outputNames,
      rawName: rawNames.join(" / "),
      outputName: outputNames.join(" / "),
      guestCount: outputNames.length,
      groupName,
      stayInfo
    });
  });

  return { valid, warnings: rowWarnings };
}

function parseCustomLineForFields(line, fields) {
  if (fields.needsRoom && fields.needsName) return parseCustomLine(line);
  if (fields.needsRoom) {
    return {
      room: line.trim(),
      names: [],
      rawNames: ""
    };
  }
  if (fields.needsName) {
    return {
      room: "",
      names: splitCustomNames(line),
      rawNames: line.trim()
    };
  }
  return { room: "", names: [], rawNames: line };
}

function parseCustomLine(line) {
  const tabParts = line.split(/\t+/).map((part) => part.trim()).filter(Boolean);
  if (tabParts.length >= 2) {
    return {
      room: tabParts[0],
      names: splitCustomNames(tabParts.slice(1).join("/")),
      rawNames: tabParts.slice(1).join(" / ")
    };
  }

  const commaParts = line.split(/\s*,\s*/).map((part) => part.trim()).filter(Boolean);
  if (commaParts.length >= 2) {
    return {
      room: commaParts[0],
      names: splitCustomNames(commaParts.slice(1).join("/")),
      rawNames: commaParts.slice(1).join(" / ")
    };
  }

  const match = line.match(/^(\S+)\s+(.+)$/);
  if (!match) return { room: line, names: [], rawNames: "" };

  return {
    room: match[1].trim(),
    names: splitCustomNames(match[2]),
    rawNames: match[2].trim()
  };
}

function splitCustomNames(value) {
  return String(value || "")
    .split(/\s*[\/／、,，;；]\s*/)
    .map((name) => name.trim())
    .filter(Boolean);
}

function extractRecords(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const valid = [];
  const rowWarnings = [];
  const startRowIndex = settings.dataStartRow - 1;
  const roomColumnIndex = XLSX.utils.decode_col(settings.roomColumn);
  const nameColumnIndex = XLSX.utils.decode_col(settings.nameColumn);
  const groupName = normalizeCell(sheet[XLSX.utils.encode_cell({ r: 7, c: 10 })]);
  let currentRecord = null;

  for (let row = startRowIndex; row <= range.e.r; row += 1) {
    const room = normalizeRoom(sheet[XLSX.utils.encode_cell({ r: row, c: roomColumnIndex })]);
    const rawName = normalizeCell(sheet[XLSX.utils.encode_cell({ r: row, c: nameColumnIndex })]);
    const arrivalCell = sheet[XLSX.utils.encode_cell({ r: row, c: XLSX.utils.decode_col("F") })];
    const nights = normalizeCell(sheet[XLSX.utils.encode_cell({ r: row, c: XLSX.utils.decode_col("K") })]);

    if (!room && !rawName) continue;

    if (room && !rawName) {
      rowWarnings.push({
        excelRow: row + 1,
        room,
        rawName,
        reason: "宿泊者名が空です"
      });
      continue;
    }

    if (!room && rawName && currentRecord) {
      currentRecord.rawNames.push(rawName);
      currentRecord.outputNames.push(normalizeGuestName(rawName));
      currentRecord.rawName = currentRecord.rawNames.join(" / ");
      currentRecord.outputName = currentRecord.outputNames.join(" / ");
      continue;
    }

    if (!room && rawName && !currentRecord) {
      rowWarnings.push({
        excelRow: row + 1,
        room,
        rawName,
        reason: "部屋番号が空です"
      });
      continue;
    }

    currentRecord = {
      index: valid.length + 1,
      excelRow: row + 1,
      room,
      rawNames: [rawName],
      outputNames: [normalizeGuestName(rawName)],
      rawName,
      outputName: normalizeGuestName(rawName),
      guestCount: 1,
      groupName,
      stayInfo: formatStayInfo(arrivalCell, nights)
    };
    valid.push(currentRecord);
  }

  valid.forEach((record, index) => {
    record.index = index + 1;
    record.guestCount = record.outputNames.length;
  });

  return { valid, warnings: rowWarnings };
}

function normalizeCell(cell) {
  if (!cell) return "";
  const value = cell.w ?? cell.v ?? "";
  if (typeof value === "number" && Number.isFinite(value)) {
    return Number.isInteger(value) ? String(value) : String(value).replace(/\.0$/, "");
  }
  return String(value).trim();
}

function normalizeRoom(cell) {
  const normalized = normalizeCell(cell);
  return normalized.replace(/\.0$/, "");
}

function normalizeGuestName(name) {
  const trimmed = String(name || "").trim();
  return trimmed.replace(/\s*様$/, "");
}

function formatGuestNameForPrint(name) {
  const normalized = normalizeGuestName(name);
  if (!settings.printNameHonorific) return normalized;
  return normalized.endsWith("様") ? normalized : `${normalized} 様`;
}

function formatStayInfo(arrivalCell, nights) {
  const dateText = formatExcelDate(arrivalCell);
  const nightsText = normalizeNights(nights);
  if (!dateText && !nightsText) return "";
  if (!dateText) return `${nightsText}泊`;
  if (!nightsText) return dateText;
  return `${dateText}〜${nightsText}泊`;
}

function formatExcelDate(cell) {
  if (!cell) return "";
  if (typeof cell.v === "number" && Number.isFinite(cell.v)) {
    const date = XLSX.SSF.parse_date_code(cell.v);
    if (date) return `${date.m}/${date.d}`;
  }
  const text = normalizeCell(cell);
  const match = text.match(/(\d{1,2})[/-](\d{1,2})/);
  return match ? `${Number(match[1])}/${Number(match[2])}` : text;
}

function normalizeNights(value) {
  const text = String(value || "").trim().replace(/\.0$/, "");
  return text;
}

function getRecordsSummary(list) {
  const guestTotal = list.reduce((sum, record) => sum + record.outputNames.length, 0);
  const unit = list.length && list.every((record) => record.room) ? "室" : "件";
  return guestTotal ? `${list.length}${unit} / ${guestTotal}名` : `${list.length}${unit}`;
}

function renderTable() {
  els.countBadge.textContent = getRecordsSummary(records);
  els.warningBadge.textContent = warnings.length ? `警告 ${warnings.length}件` : "";

  const rows = records.map((record) => `
    <tr>
      <td>${record.index}</td>
      <td>${record.excelRow}</td>
      <td>${escapeHtml(record.room)}</td>
      <td>${escapeHtml(record.rawName)}</td>
      <td>${record.outputNames.length}</td>
      <td>${escapeHtml(record.stayInfo)}</td>
      <td>${escapeHtml(getPrintableNames(record).join(" / "))}</td>
      <td>使用</td>
    </tr>
  `);

  const warningRows = warnings.map((warning) => `
    <tr class="bad-row">
      <td>-</td>
      <td>${warning.excelRow}</td>
      <td>${escapeHtml(warning.room)}</td>
      <td>${escapeHtml(warning.rawName)}</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>${warning.reason}</td>
    </tr>
  `);

  els.dataTable.innerHTML = rows.concat(warningRows).join("") || '<tr><td colspan="8" class="empty">Excelファイルを選択するとデータが表示されます。</td></tr>';
  updateCopyModeVisibility();
}

function ensureData() {
  if (records.length) return true;
  setStatus("先にExcelファイルを選択してください。", true);
  setSimpleStatus("先にExcelファイルを選択してください。", true);
  return false;
}

function getRecordsForMode(mode) {
  if (mode === "test") return [records[0]];
  if (mode === "all") return records;

  syncRangeInputs();
  const rangeStartInput = els.simpleApp.hidden ? els.rangeStart : els.simpleRangeStart;
  const rangeEndInput = els.simpleApp.hidden ? els.rangeEnd : els.simpleRangeEnd;
  const start = Math.max(1, Number(rangeStartInput.value || 1));
  const end = Math.min(records.length, Number(rangeEndInput.value || records.length));

  if (start > end) {
    setStatus("開始番号が終了番号より大きくなっています。", true);
    setSimpleStatus("開始番号が終了番号より大きくなっています。", true);
    return [];
  }

  return records.slice(start - 1, end);
}

function openPrintWindow(mode) {
  if (!ensureData()) return;
  if (els.simpleApp.hidden) {
    updateSettingsFromForm();
    bindSimpleSettingsToForm();
  } else {
    updateSettingsFromSimpleForm();
    updateSettingsFromSimplePositionForm();
    bindSettingsToForm();
  }
  if (!validateRequiredPrintField()) return;

  const selected = getRecordsForMode(mode);
  if (!selected.length) return;
  const printableRecords = getPrintableRecords(selected, mode);

  const printWindow = window.open("", "_blank");
  if (!printWindow) {
    setStatus("印刷画面を開けません。ポップアップを許可してください。", true);
    return;
  }

  printWindow.document.open();
  printWindow.document.write(buildPrintHtml(printableRecords));
  printWindow.document.close();
  printWindow.focus();
  const copyNote = printableRecords.length === selected.length ? "" : ` ${printableRecords.length}枚分を作成しました。`;
  setStatus(`印刷画面を開きました。用紙サイズ B6（横）、倍率100%で印刷してください。${copyNote}`);
  setSimpleStatus(`印刷画面を開きました。レジカード用プリンターを選択してから印刷してください。${copyNote}`);
}

function getPrintableRecords(selected, mode) {
  if (mode === "test" || settings.printCopyMode !== "guest") return selected;

  return selected.flatMap((record) => {
    const copyCount = Math.max(1, record.outputNames.length);
    return Array.from({ length: copyCount }, () => record);
  });
}

function buildPrintHtml(selected) {
  const pages = selected.map((record) => buildPrintPage(record)).join("");
  return `<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <title>キーカバー印刷</title>
  <style>
    @page {
      size: ${B6_WIDTH_MM}mm ${B6_HEIGHT_MM}mm;
      margin: 0;
    }
    * {
      box-sizing: border-box;
    }
    html,
    body {
      margin: 0;
      padding: 0;
      background: #fff;
      font-family: "Yu Gothic", "Meiryo", "Noto Sans JP", sans-serif;
    }
    .toolbar {
      position: fixed;
      top: 12px;
      left: 12px;
      z-index: 10;
      display: flex;
      gap: 8px;
      align-items: center;
      padding: 10px 12px;
      border: 1px solid #c8d0d9;
      border-radius: 6px;
      background: #fff;
      box-shadow: 0 2px 10px rgba(15, 23, 42, 0.14);
      font-size: 13px;
    }
    button {
      border: 1px solid #176b87;
      border-radius: 5px;
      background: #176b87;
      color: #fff;
      padding: 7px 11px;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
    }
    .page {
      position: relative;
      width: ${B6_WIDTH_MM}mm;
      height: ${B6_HEIGHT_MM}mm;
      overflow: hidden;
      page-break-after: always;
      break-after: page;
      background: #fff;
    }
    .guide {
      position: absolute;
      top: 0;
      left: 50%;
      width: ${settings.pageWidthMm}mm;
      height: ${settings.pageHeightMm}mm;
      transform: translateX(-50%);
      border: 0.4mm dashed #8b98a8;
      pointer-events: none;
    }
    .name,
    .room,
    .group-name,
    .stay-info {
      position: absolute;
      margin: 0;
      padding: 0;
      color: #000;
      line-height: 1;
      white-space: nowrap;
      font-weight: 400;
    }
    .name {
      font-size: ${settings.nameFontSize}pt;
      transform: ${printTransform()};
    }
    .room {
      left: ${printX(settings.roomX)}mm;
      top: ${printY(settings.roomY)}mm;
      font-size: ${settings.roomFontSize}pt;
      transform: ${printTransform()};
    }
    .group-name {
      left: ${printX(settings.groupNameX)}mm;
      top: ${printY(settings.groupNameY)}mm;
      font-size: ${settings.groupNameFontSize}pt;
      transform: ${printTransform()};
    }
    .stay-info {
      left: ${printX(settings.stayInfoX)}mm;
      top: ${printY(settings.stayInfoY)}mm;
      font-size: ${settings.stayInfoFontSize}pt;
      transform: ${printTransform()};
    }
    @media screen {
      body {
        background: #eef1f4;
        padding: 70px 20px 20px;
      }
      .page {
        margin: 0 auto 16px;
        outline: 1px solid #9aa8b7;
      }
    }
    @media print {
      .toolbar {
        display: none;
      }
      body {
        width: ${B6_WIDTH_MM}mm;
      }
      .page {
        margin: 0;
        outline: none;
      }
      .guide {
        display: none;
      }
    }
  </style>
</head>
<body>
  <div class="toolbar">
    <button type="button" onclick="window.print()">印刷</button>
    <span>レジカード用プリンターを選択 / 用紙 B6（横） / 倍率100%</span>
  </div>
  ${pages}
</body>
</html>`;
}

function buildPrintPage(record) {
  const printableNames = getPrintableNames(record);
  const nameLayout = getNameLayout(record);
  const groupName = settings.printGroupName && record.groupName
    ? `<p class="group-name">${escapeHtml(record.groupName)}</p>`
    : "";
  const stayInfo = settings.printStayInfo && record.stayInfo
    ? `<p class="stay-info" style="top:${printY(nameLayout.stayInfoY)}mm;">${escapeHtml(record.stayInfo)}</p>`
    : "";
  const names = printableNames.map((name, index) => (
    `<p class="name" style="left:${printX(settings.nameX)}mm; top:${printY(nameLayout.firstNameY + nameLayout.lineGap * index)}mm; font-size:${nameLayout.fontSize}pt;">${escapeHtml(name)}</p>`
  )).join("");
  const room = settings.printRoom
    ? `<p class="room">${escapeHtml(record.room)}</p>`
    : "";

  return `<section class="page">
    <div class="guide"></div>
    ${groupName}
    ${stayInfo}
    ${names}
    ${room}
  </section>`;
}

function getNameLayout(record) {
  const names = getPrintableNames(record);
  const count = Math.max(1, names.length);
  const fontSize = getAutoNameFontSize(names);
  const fontHeightMm = fontSize * PT_TO_MM;
  const lineGap = roundToTenth(Math.max(settings.nameLineGap, fontHeightMm * 1.35));
  const extraLines = Math.max(0, count - 2);
  const baseShift = extraLines * lineGap * 0.55;
  const projectedLastNameY = Number(settings.nameY) - baseShift + (count - 1) * lineGap;
  const roomClearanceY = settings.printRoom
    ? Number(settings.roomY) - Number(settings.roomFontSize) * PT_TO_MM * 0.85
    : Number(settings.roomY);
  const clearanceShift = Math.max(0, projectedLastNameY - roomClearanceY);
  const upwardShift = roundToTenth(baseShift + clearanceShift);

  return {
    fontSize,
    lineGap,
    firstNameY: Number(settings.nameY) - upwardShift,
    stayInfoY: Number(settings.stayInfoY) - upwardShift
  };
}

function getPrintableNames(record) {
  if (!settings.printName) return [];
  return record.outputNames.map(formatGuestNameForPrint);
}

function getAutoNameFontSize(names) {
  const baseSize = Number(settings.nameFontSize);
  const longestWidth = names.reduce((max, name) => (
    Math.max(max, estimateTextWidthMm(name, baseSize))
  ), 0);

  if (longestWidth <= NAME_AREA_WIDTH_MM) return baseSize;

  const fittedSize = baseSize * (NAME_AREA_WIDTH_MM / longestWidth);
  return roundToTenth(Math.max(MIN_NAME_FONT_SIZE_PT, fittedSize));
}

function estimateTextWidthMm(text, fontPt) {
  const units = Array.from(String(text || "")).reduce((sum, char) => {
    if (/\s/.test(char)) return sum + 0.35;
    if (/[\u0000-\u007f]/.test(char)) return sum + 0.58;
    if (/[\uff61-\uff9f]/.test(char)) return sum + 0.62;
    return sum + 1;
  }, 0);
  return units * fontPt * PT_TO_MM;
}

function roundToTenth(value) {
  return Math.round(Number(value) * 10) / 10;
}

function printX(baseX) {
  const x = Number(baseX) + Number(settings.globalOffsetX);
  return settings.rotate180 ? B6_WIDTH_MM - x : x;
}

function printY(baseY) {
  const y = Number(baseY) + Number(settings.globalOffsetY);
  return settings.rotate180 ? B6_HEIGHT_MM - y : y;
}

function printTransform() {
  return settings.rotate180 ? "translateY(-50%) rotate(180deg)" : "translateY(-50%)";
}

function setStatus(message, isError = false) {
  els.status.textContent = message;
  els.status.classList.toggle("error", isError);
}

function setSimpleStatus(message, isError = false) {
  els.simpleStatus.textContent = message;
  els.simpleStatus.classList.toggle("error", isError);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
