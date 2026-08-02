const DEFAULT_SETTINGS = {
  pageWidthMm: 130,
  pageHeightMm: 110,
  printAreaStartX: 65,
  nameX: 98,
  nameY: 70,
  nameAlign: "left",
  nameFontSize: 12,
  roomX: 98,
  roomY: 88,
  roomAlign: "left",
  roomFontSize: 20,
  printName: true,
  printNameHonorific: true,
  printRoom: true,
  nameLineGap: 6,
  groupNameX: 98,
  groupNameY: 14,
  groupNameFontSize: 9,
  stayInfoX: 98,
  stayInfoY: 44,
  stayInfoFontSize: 11,
  cleaningInfoX: 98,
  cleaningInfoY: 52,
  cleaningInfoFontSize: 11,
  rcInfoX: 98,
  rcInfoY: 60,
  rcInfoFontSize: 11,
  printGroupName: false,
  printStayInfo: false,
  printStaySchedule: false,
  printCleaningInfo: false,
  cleaningInfoCustomText: "",
  printRcInfo: false,
  rcInfoCustomText: "",
  printCopyMode: "room",
  printOrder: "load",
  globalOffsetX: 0,
  globalOffsetY: 0,
  rotate180: false,
  dataStartRow: 13,
  roomColumn: "U",
  nameColumn: "X"
};

const STORAGE_KEY = "keycoverPrintSettings.v2.b6";
const SHEET_NAME = "団体メンバ一覧表";
const BUILD_VERSION = "20260803-fontmenu";
const B6_WIDTH_MM = 182;
const B6_HEIGHT_MM = 128;
const NAME_AREA_WIDTH_MM = 58;
const MIN_WRAPPED_TEXT_WIDTH_MM = 1;
const MIN_NAME_FONT_SIZE_PT = 6.5;
const PT_TO_MM = 0.352778;
const FONT_SIZE_KEYS = [
  "nameFontSize",
  "roomFontSize",
  "groupNameFontSize",
  "stayInfoFontSize",
  "cleaningInfoFontSize",
  "rcInfoFontSize"
];

let settings = loadSettings();
let records = [];
let warnings = [];
let currentWorkbook = null;
let simpleStep = 1;
let selectedRecordIndexes = new Set();

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
  simplePrintSelection: document.getElementById("simplePrintSelection"),
  simplePrintOptionsPanel: document.getElementById("simplePrintOptionsPanel"),
  simpleCopyModeHint: document.getElementById("simpleCopyModeHint"),
  advancedPrintOptionsPanel: document.getElementById("advancedPrintOptionsPanel"),
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
  simpleGroupNameInput: document.getElementById("simpleGroupNameInput"),
  advancedGroupNameInput: document.getElementById("advancedGroupNameInput"),
  simpleGroupNamePreviewBadge: document.getElementById("simpleGroupNamePreviewBadge"),
  simpleNamePreviewChip: document.getElementById("simpleNamePreviewChip"),
  simpleRoomPreviewChip: document.getElementById("simpleRoomPreviewChip"),
  simpleStayInfoPreviewChip: document.getElementById("simpleStayInfoPreviewChip"),
  staySchedulePreview: document.getElementById("staySchedulePreview"),
  simpleAdditionalInfo: document.getElementById("simpleAdditionalInfo"),
  advancedAdditionalInfo: document.getElementById("advancedAdditionalInfo"),
  simpleStayInfoDetails: document.getElementById("simpleStayInfoDetails"),
  advancedStayInfoDetails: document.getElementById("advancedStayInfoDetails"),
  simpleAdditionalSummary: document.getElementById("simpleAdditionalSummary"),
  advancedAdditionalSummary: document.getElementById("advancedAdditionalSummary"),
  simpleStaySelectionSummary: document.getElementById("simpleStaySelectionSummary"),
  advancedStaySelectionSummary: document.getElementById("advancedStaySelectionSummary"),
  simpleLiveGroup: document.getElementById("simpleLiveGroup"),
  simpleLiveStay: document.getElementById("simpleLiveStay"),
  simpleLiveCleaning: document.getElementById("simpleLiveCleaning"),
  simpleLiveRc: document.getElementById("simpleLiveRc"),
  simpleLiveName: document.getElementById("simpleLiveName"),
  simpleLiveRoom: document.getElementById("simpleLiveRoom"),
  advancedPrintSelection: document.getElementById("advancedPrintSelection"),
  recordEditModal: document.getElementById("recordEditModal"),
  saveRecordEdit: document.getElementById("saveRecordEdit"),
  selectAllRows: document.getElementById("selectAllRows"),
  buildVersion: document.getElementById("buildVersion")
};

const settingInputs = [
  "nameFontSize",
  "roomFontSize",
  "groupNameFontSize",
  "stayInfoFontSize",
  "cleaningInfoFontSize",
  "rcInfoFontSize",
  "printName",
  "printNameHonorific",
  "printRoom",
  "printGroupName",
  "printStayInfo",
  "printStaySchedule",
  "printCleaningInfo",
  "printRcInfo"
];

const simplePositionInputMap = {
  simpleNameFontSize: "nameFontSize",
  simpleRoomFontSize: "roomFontSize",
  simpleGroupNameFontSize: "groupNameFontSize",
  simpleStayInfoFontSize: "stayInfoFontSize",
  simpleCleaningInfoFontSize: "cleaningInfoFontSize",
  simpleRcInfoFontSize: "rcInfoFontSize"
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

  document.getElementById("startSimple").addEventListener("click", () => {
    showSimpleStep(1);
    showMode("simple");
  });
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
  ["simplePrintName", "simplePrintNameHonorific", "simplePrintRoom", "simplePrintGroupName", "simplePrintStayInfo", "simplePrintStaySchedule", "simplePrintCleaningInfo", "simplePrintRcInfo"].forEach((id) => {
    document.getElementById(id).addEventListener("input", () => {
      updateSettingsFromSimpleForm();
      bindSettingsToForm();
      bindCustomSettingsToForm();
      bindCleaningSettingsToForm();
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

  const handleGroupNameChange = (e) => {
    const val = e.target.value;
    if (els.simpleGroupNameInput && els.simpleGroupNameInput !== e.target) els.simpleGroupNameInput.value = val;
    if (els.advancedGroupNameInput && els.advancedGroupNameInput !== e.target) els.advancedGroupNameInput.value = val;
    records.forEach((r) => { r.groupName = val; });
    renderTable();
  };
  if (els.simpleGroupNameInput) els.simpleGroupNameInput.addEventListener("input", handleGroupNameChange);
  if (els.advancedGroupNameInput) els.advancedGroupNameInput.addEventListener("input", handleGroupNameChange);

  getDatePickerConfigs().forEach((config) => {
    document.getElementById(config.monthId).addEventListener("input", () => renderDatePicker(config));
    document.getElementById(config.clearButtonId).addEventListener("click", () => {
      settings[config.textKey] = "";
      bindCleaningSettingsToForm();
      renderTable();
    });
    document.getElementById(config.openButtonId).addEventListener("click", () => {
      openDatePickerModal(config.panelId);
    });
    document.getElementById(config.inputId).addEventListener("change", (event) => {
      if (event.target.checked) openDatePickerModal(config.panelId);
    });
  });
  document.querySelectorAll("[data-close-date-picker]").forEach((button) => {
    button.addEventListener("click", () => closeDatePickerModal(button.dataset.closeDatePicker));
  });
  document.querySelectorAll(".date-picker-modal").forEach((dialog) => {
    dialog.addEventListener("click", (event) => {
      if (event.target === dialog) dialog.close();
    });
  });
  document.querySelectorAll('input[name="simpleCopyMode"], input[name="advancedCopyMode"]').forEach((input) => {
    input.addEventListener("input", () => {
      updateCopyModeFromForm(input.value);
      bindCopyModeToForm();
    });
  });
  document.querySelectorAll('input[name="simplePrintOrder"], input[name="advancedPrintOrder"]').forEach((input) => {
    input.addEventListener("input", () => {
      updatePrintOrderFromForm(input.value);
      bindPrintOrderToForm();
      updatePrintSelectionPanels();
    });
  });
  Object.keys(simplePositionInputMap).forEach((id) => {
    document.getElementById(id).addEventListener("input", () => {
      updateSettingsFromSimplePositionForm();
      bindSettingsToForm();
      updatePreviewChips();
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
  [els.simplePrintSelection, els.advancedPrintSelection].forEach((panel) => {
    panel.addEventListener("change", handlePrintSelectionChange);
    panel.addEventListener("click", handlePrintSelectionAction);
    panel.addEventListener("click", handleEditButtonClick);
  });
  els.recordEditModal.querySelector("[data-close-record-edit]").addEventListener("click", () => els.recordEditModal.close());
  els.saveRecordEdit.addEventListener("click", saveIndividualRecordEdit);
  els.dataTable.addEventListener("change", handlePrintSelectionChange);
  els.selectAllRows.addEventListener("change", () => {
    setAllPrintSelection(els.selectAllRows.checked);
    renderTable();
  });
  updatePrintFieldControls();
  bindCleaningSettingsToForm();
  updateCustomEntryHints();
}

function showMode(mode) {
  els.modeHome.hidden = mode !== "home";
  els.simpleApp.hidden = mode !== "simple";
  els.advancedApp.hidden = mode !== "advanced";
  updateCopyModeVisibility();
  if (mode === "advanced") updatePrintSelectionPanels();
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
  if (step === 4) {
    updateCopyModeVisibility();
    updatePrintSelectionPanels();
  }
}

function loadSettings() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
    const migrated = { ...DEFAULT_SETTINGS, ...(saved || {}) };
    if (saved && typeof saved.printStaySchedule !== "boolean") {
      migrated.printStaySchedule = Boolean(saved.printStayInfo);
    }
    [
      "globalOffsetX",
      "globalOffsetY",
      "nameX",
      "nameY",
      "roomX",
      "roomY",
      "rotate180"
    ].forEach((key) => {
      migrated[key] = DEFAULT_SETTINGS[key];
    });
    return migrated;
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

function bindSettingsToForm() {
  syncStayInfoMaster();
  settingInputs.forEach((key) => {
    const input = document.getElementById(key);
    if (input.type === "checkbox") {
      input.checked = Boolean(settings[key]);
    } else {
      input.value = settings[key];
    }
  });
  if (settings.printGroupName || settings.printStayInfo) {
    els.advancedAdditionalInfo.open = true;
  }
  if (settings.printStayInfo) {
    els.advancedStayInfoDetails.open = true;
  }
  updatePrintFieldControls();
}

function bindSimpleSettingsToForm() {
  syncStayInfoMaster();
  document.getElementById("simplePrintName").checked = Boolean(settings.printName);
  document.getElementById("simplePrintNameHonorific").checked = Boolean(settings.printNameHonorific);
  document.getElementById("simplePrintRoom").checked = Boolean(settings.printRoom);
  document.getElementById("simplePrintGroupName").checked = Boolean(settings.printGroupName);
  document.getElementById("simplePrintStayInfo").checked = Boolean(settings.printStayInfo);
  document.getElementById("simplePrintStaySchedule").checked = Boolean(settings.printStaySchedule);
  document.getElementById("simplePrintCleaningInfo").checked = Boolean(settings.printCleaningInfo);
  document.getElementById("simplePrintRcInfo").checked = Boolean(settings.printRcInfo);
  if (settings.printGroupName || settings.printStayInfo) {
    els.simpleAdditionalInfo.open = true;
  }
  if (settings.printStayInfo) {
    els.simpleStayInfoDetails.open = true;
  }
  bindCopyModeToForm();
  bindPrintOrderToForm();
  updatePrintFieldControls();
  bindCleaningSettingsToForm();
}

function bindCleaningSettingsToForm() {
  getDatePickerConfigs().forEach((config) => {
    const monthInput = document.getElementById(config.monthId);
    if (!monthInput.value) monthInput.value = getDefaultPickerMonth();
    renderDatePicker(config);
  });
  updateCleaningControls();
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

function bindPrintOrderToForm() {
  const order = settings.printOrder === "room" ? "room" : "load";
  document.querySelectorAll('input[name="simplePrintOrder"], input[name="advancedPrintOrder"]').forEach((input) => {
    input.checked = input.value === order;
  });
}

function updateCopyModeFromForm(value) {
  settings.printCopyMode = value === "guest" ? "guest" : "room";
}

function updatePrintOrderFromForm(value) {
  settings.printOrder = value === "room" ? "room" : "load";
}

function updateCopyModeFromActiveForm() {
  const selector = els.simpleApp.hidden
    ? 'input[name="advancedCopyMode"]:checked'
    : 'input[name="simpleCopyMode"]:checked';
  const selected = document.querySelector(selector);
  if (selected) updateCopyModeFromForm(selected.value);
}

function updatePrintOrderFromActiveForm() {
  const selector = els.simpleApp.hidden
    ? 'input[name="advancedPrintOrder"]:checked'
    : 'input[name="simplePrintOrder"]:checked';
  const selected = document.querySelector(selector);
  if (selected) updatePrintOrderFromForm(selected.value);
}

function compareRoomNumbers(a, b) {
  const roomA = String(a.room || "");
  const roomB = String(b.room || "");
  const numA = Number(roomA);
  const numB = Number(roomB);
  if (roomA !== "" && roomB !== "" && !Number.isNaN(numA) && !Number.isNaN(numB)) {
    return numA - numB;
  }
  return roomA.localeCompare(roomB, "ja", { numeric: true });
}

function sortRecordsForPrint(list) {
  if (settings.printOrder !== "room") return list;
  return [...list].sort((a, b) => {
    const roomCompare = compareRoomNumbers(a, b);
    if (roomCompare !== 0) return roomCompare;
    return a.index - b.index;
  });
}

function getRecordsInDisplayOrder() {
  return sortRecordsForPrint(records);
}

function getMultiOccupancySummary() {
  const multiRooms = records.filter((record) => record.outputNames.length > 1);
  const extraCards = multiRooms.reduce((sum, record) => sum + record.outputNames.length - 1, 0);
  return { multiRooms, extraCards };
}

function updateCopyModeVisibility() {
  const { multiRooms, extraCards } = getMultiOccupancySummary();
  const hasRecords = records.length > 0;
  const hasMultiRooms = multiRooms.length > 0;
  const hint = hasMultiRooms
    ? `2名以上の部屋が${multiRooms.length}室あります。人数分の場合は追加で${extraCards}枚印刷されます。`
    : "";

  [els.simplePrintOptionsPanel, els.advancedPrintOptionsPanel].forEach((panel) => {
    if (panel) panel.hidden = !hasRecords;
  });

  // 印刷枚数セクション：1室1人のみの場合は非表示
  ["simpleCopyModeRow", "advancedCopyModeRow"].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.hidden = !hasMultiRooms;
  });

  els.simpleCopyModeHint.textContent = hint;
  els.advancedCopyModeHint.textContent = hint;
}

function resetPrintSelection() {
  selectedRecordIndexes = new Set(records.map((record) => String(record.index)));
}

function setAllPrintSelection(isSelected) {
  selectedRecordIndexes = isSelected
    ? new Set(records.map((record) => String(record.index)))
    : new Set();
}

function handlePrintSelectionChange(event) {
  const checkbox = event.target.closest("[data-print-select-index]");
  if (!checkbox) return;

  const index = checkbox.dataset.printSelectIndex;
  if (checkbox.checked) {
    selectedRecordIndexes.add(index);
  } else {
    selectedRecordIndexes.delete(index);
  }

  // スクロール位置を維持するため、チェックボックス変更時はDOMを全再描画せず
  // 対象行のCSS更新とツールバーカウントのみ更新する
  const selectedRecords = getSelectedRecords();


  // 全ての選択リストパネル의 行スタイルとツールバー를 軽量更新
  [els.simplePrintSelection, els.advancedPrintSelection].forEach((panel) => {
    if (!panel) return;

    // 行のunselected-rowクラス를 更新
    panel.querySelectorAll("[data-print-select-index]").forEach((cb) => {
      const row = cb.closest(".selection-row");
      if (!row) return;
      const isSelected = selectedRecordIndexes.has(cb.dataset.printSelectIndex);
      cb.checked = isSelected;
      row.classList.toggle("unselected-row", !isSelected);
    });

    // ツールバー의 カウント表示를 更新
    const countEl = panel.querySelector(".selection-toolbar strong");
    if (countEl) countEl.textContent = `選択中 ${selectedRecords.length}件 / ${records.length}件`;
  });


  // dataTable의 チェックボックスから来た場合はテーブル全体の選択状態を更新
  if (event.target.closest("#dataTable") || event.currentTarget === els.dataTable) {
    updateTableSelectionControl(selectedRecords);
  }
}

function handlePrintSelectionAction(event) {
  const button = event.target.closest("[data-selection-action]");
  if (!button) return;

  setAllPrintSelection(button.dataset.selectionAction === "all");
  renderTable();
}

function isRecordSelected(record) {
  return selectedRecordIndexes.has(String(record.index));
}

function getSelectedRecords() {
  return records.filter(isRecordSelected);
}

function updateSettingsFromForm() {
  settingInputs.forEach((key) => {
    const input = document.getElementById(key);
    settings[key] = input.type === "checkbox" ? input.checked : Number(input.value);
  });
  syncStayInfoMaster();
  const selected = document.querySelector('input[name="advancedCopyMode"]:checked');
  if (selected) settings.printCopyMode = selected.value;
  const printOrder = document.querySelector('input[name="advancedPrintOrder"]:checked');
  if (printOrder) settings.printOrder = printOrder.value;
  updatePrintFieldControls();
  updateCleaningControls();
}

function updateSettingsFromSimpleForm() {
  settings.printName = document.getElementById("simplePrintName").checked;
  settings.printNameHonorific = document.getElementById("simplePrintNameHonorific").checked;
  settings.printRoom = document.getElementById("simplePrintRoom").checked;
  settings.printGroupName = document.getElementById("simplePrintGroupName").checked;
  settings.printStaySchedule = document.getElementById("simplePrintStaySchedule").checked;
  settings.printCleaningInfo = document.getElementById("simplePrintCleaningInfo").checked;
  settings.printRcInfo = document.getElementById("simplePrintRcInfo").checked;
  syncStayInfoMaster();
  const selected = document.querySelector('input[name="simpleCopyMode"]:checked');
  if (selected) settings.printCopyMode = selected.value;
  const printOrder = document.querySelector('input[name="simplePrintOrder"]:checked');
  if (printOrder) settings.printOrder = printOrder.value;
  updatePrintFieldControls();
  updateCleaningControls();
}

function syncStayInfoMaster() {
  settings.printStayInfo = Boolean(
    settings.printStaySchedule
    || settings.printCleaningInfo
    || settings.printRcInfo
  );
  ["simplePrintStayInfo", "printStayInfo"].forEach((id) => {
    const input = document.getElementById(id);
    if (input) input.checked = settings.printStayInfo;
  });
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

function updatePreviewChips() {
  const firstRecord = records[0];

  // 氏名 Preview Chip
  if (els.simpleNamePreviewChip) {
    if (firstRecord && firstRecord.outputNames && firstRecord.outputNames.length) {
      const formatted = formatGuestNameForPrint(firstRecord.outputNames[0]);
      els.simpleNamePreviewChip.textContent = `例: ${formatted}`;
    } else {
      els.simpleNamePreviewChip.textContent = settings.printNameHonorific ? "例: 山田 太郎 様" : "例: 山田 太郎";
    }
  }

  // 部屋番号 Preview Chip
  if (els.simpleRoomPreviewChip) {
    if (firstRecord && firstRecord.room) {
      els.simpleRoomPreviewChip.textContent = `例: ${firstRecord.room}`;
    } else {
      els.simpleRoomPreviewChip.textContent = "例: 704";
    }
  }

  // 団体名 (Group Name) Input & Preview
  const groupNameVal = (firstRecord && firstRecord.groupName) ? firstRecord.groupName : (els.simpleGroupNameInput ? els.simpleGroupNameInput.value : "");
  if (els.simpleGroupNameInput && document.activeElement !== els.simpleGroupNameInput) {
    els.simpleGroupNameInput.value = groupNameVal;
  }
  if (els.advancedGroupNameInput && document.activeElement !== els.advancedGroupNameInput) {
    els.advancedGroupNameInput.value = groupNameVal;
  }
  if (els.simpleGroupNamePreviewBadge) {
    els.simpleGroupNamePreviewBadge.hidden = !settings.printGroupName;
  }

  // 宿泊日程 (Stay Schedule) Preview
  const stayInfoText = (firstRecord && firstRecord.stayInfo) ? firstRecord.stayInfo : "";
  if (els.simpleStayInfoPreviewChip) {
    if (stayInfoText) {
      els.simpleStayInfoPreviewChip.textContent = `例: ${stayInfoText}`;
      els.simpleStayInfoPreviewChip.classList.add("has-value");
    } else {
      els.simpleStayInfoPreviewChip.textContent = "例: 7/7〜1泊";
      els.simpleStayInfoPreviewChip.classList.remove("has-value");
    }
  }
  if (els.staySchedulePreview) {
    els.staySchedulePreview.textContent = `例: ${stayInfoText || "7/7〜1泊"}`;
  }

  updateAdditionalInformationControls();
  updateLivePrintPreview(firstRecord);
}

function updateCleaningControls() {
  syncStayInfoMaster();

  getDatePickerConfigs().forEach((config) => {
    const enabled = Boolean(settings[config.printKey]);
    const panel = document.getElementById(config.panelId);
    const monthInput = document.getElementById(config.monthId);
    const clearBtn = document.getElementById(config.clearButtonId);
    const selectedEl = document.getElementById(config.selectedId);
    const previewEl = document.getElementById(config.previewId);
    const openButton = document.getElementById(config.openButtonId);

    if (panel) panel.hidden = !enabled;
    if (panel && !enabled && panel.open) panel.close();
    if (monthInput) monthInput.disabled = !enabled;
    if (clearBtn) clearBtn.disabled = !enabled;
    if (openButton) {
      openButton.hidden = !enabled;
      openButton.disabled = !enabled;
    }
    if (selectedEl) selectedEl.textContent = formatSelectedDates(settings[config.textKey]);

    if (previewEl) {
      const dates = parseDateInfoDates(settings[config.textKey]);
      const prefix = config.label === "清掃" ? "清掃日：" : "部屋変更：";
      if (enabled && dates.length) {
        previewEl.textContent = `${prefix}${dates.join(", ")}`;
        previewEl.classList.add("has-value");
      } else {
        previewEl.textContent = enabled ? "日付を選択" : "未選択";
        previewEl.classList.remove("has-value");
      }
    }
  });

  updatePreviewChips();
}

function openDatePickerModal(panelId) {
  const dialog = document.getElementById(panelId);
  if (!dialog || dialog.hidden || dialog.open) return;
  dialog.showModal();
}

function closeDatePickerModal(panelId) {
  const dialog = document.getElementById(panelId);
  if (dialog?.open) dialog.close();
}

function updateAdditionalInformationControls() {
  const simpleGroupBody = document.getElementById("simpleGroupNameBody");
  const advancedGroupBody = document.getElementById("advancedGroupNameBody");
  if (simpleGroupBody) simpleGroupBody.hidden = !settings.printGroupName;
  if (advancedGroupBody) advancedGroupBody.hidden = !settings.printGroupName;

  const stayCount = [
    settings.printStaySchedule,
    settings.printCleaningInfo,
    settings.printRcInfo
  ].filter(Boolean).length;
  const additionalCount = stayCount + (settings.printGroupName ? 1 : 0);
  const staySummary = stayCount ? `${stayCount}項目選択` : "未選択";
  const additionalSummary = additionalCount ? `${additionalCount}項目選択` : "未選択";

  [els.simpleStaySelectionSummary, els.advancedStaySelectionSummary].forEach((element) => {
    if (element) element.textContent = staySummary;
  });
  [els.simpleAdditionalSummary, els.advancedAdditionalSummary].forEach((element) => {
    if (element) element.textContent = additionalSummary;
  });
}

function updateLivePrintPreview(firstRecord) {
  const record = firstRecord || {};
  const previewRecord = {
    ...record,
    outputNames: record.outputNames?.length ? record.outputNames : ["山田 太郎"]
  };
  const printableNames = previewRecord.outputNames.map(formatGuestNameForPrint);
  const nameLayout = getNameLayout(previewRecord);
  const groupName = record.groupName || els.simpleGroupNameInput?.value || "東横商事 御一行様";
  const stayInfo = record.stayInfo || "7/7〜1泊";
  const room = record.room || "704";

  setLivePreviewLine(els.simpleLiveGroup, settings.printGroupName, groupName);
  positionLivePreviewLine(
    els.simpleLiveGroup,
    settings.groupNameX,
    settings.groupNameY,
    settings.groupNameFontSize
  );
  setLivePreviewLine(els.simpleLiveStay, settings.printStayInfo && settings.printStaySchedule, stayInfo);
  positionLivePreviewLine(
    els.simpleLiveStay,
    settings.stayInfoX,
    nameLayout.stayInfoY,
    settings.stayInfoFontSize
  );
  setLivePreviewLine(
    els.simpleLiveCleaning,
    settings.printStayInfo && settings.printCleaningInfo,
    settings.cleaningInfoCustomText || "清掃日：未設定"
  );
  positionLivePreviewLine(
    els.simpleLiveCleaning,
    settings.cleaningInfoX,
    nameLayout.cleaningInfoY,
    settings.cleaningInfoFontSize
  );
  setLivePreviewLine(
    els.simpleLiveRc,
    settings.printStayInfo && settings.printRcInfo,
    settings.rcInfoCustomText || "部屋変更：未設定"
  );
  positionLivePreviewLine(
    els.simpleLiveRc,
    settings.rcInfoX,
    nameLayout.rcInfoY,
    settings.rcInfoFontSize
  );
  setLivePreviewLine(els.simpleLiveName, settings.printName, printableNames.join("\n"));
  positionLivePreviewLine(
    els.simpleLiveName,
    settings.nameX,
    nameLayout.firstNameY,
    nameLayout.fontSize,
    nameLayout.lineGap
  );
  setLivePreviewLine(els.simpleLiveRoom, settings.printRoom, room);
  positionLivePreviewLine(
    els.simpleLiveRoom,
    settings.roomX,
    settings.roomY,
    settings.roomFontSize
  );
}

function setLivePreviewLine(element, visible, text) {
  if (!element) return;
  element.hidden = !visible;
  element.textContent = text;
}

function positionLivePreviewLine(element, xMm, yMm, fontPt, lineGapMm = null) {
  if (!element) return;
  const faceWidthMm = Number(settings.pageWidthMm) - Number(settings.printAreaStartX);
  const faceStartX = getPrintFaceStartX();
  const faceHeightMm = Number(settings.pageHeightMm);
  const xPercent = ((printX(xMm) - faceStartX) / faceWidthMm) * 100;
  const yPercent = (printY(yMm) / faceHeightMm) * 100;
  const fontSizeCqw = (Number(fontPt) * PT_TO_MM / faceWidthMm) * 100;
  const maxWidthCqw = (getPrintFaceRemainingWidth(xMm) / faceWidthMm) * 100;

  element.style.left = `${xPercent}%`;
  element.style.top = `${yPercent}%`;
  element.style.fontSize = `${fontSizeCqw}cqw`;
  element.style.maxWidth = `${maxWidthCqw}cqw`;
  element.style.transform = settings.rotate180
    ? "translateY(-50%) rotate(180deg)"
    : "translateY(-50%)";
  element.style.lineHeight = lineGapMm
    ? `${(Number(lineGapMm) / faceWidthMm) * 100}cqw`
    : "1.15";
}

function getDatePickerConfigs() {
  return [
    {
      label: "清掃",
      maxDates: 6,
      printKey: "printCleaningInfo",
      inputId: "printCleaningInfo",
      textKey: "cleaningInfoCustomText",
      panelId: "cleaningInfoPicker",
      openButtonId: "openCleaningInfoPicker",
      monthId: "cleaningInfoMonth",
      calendarId: "cleaningInfoCalendar",
      clearButtonId: "clearCleaningInfoDates",
      selectedId: "cleaningInfoSelectedDates",
      previewId: "cleaningInfoPreview"
    },
    {
      label: "清掃",
      maxDates: 6,
      printKey: "printCleaningInfo",
      inputId: "simplePrintCleaningInfo",
      textKey: "cleaningInfoCustomText",
      panelId: "simpleCleaningInfoPicker",
      openButtonId: "simpleOpenCleaningInfoPicker",
      monthId: "simpleCleaningInfoMonth",
      calendarId: "simpleCleaningInfoCalendar",
      clearButtonId: "simpleClearCleaningInfoDates",
      selectedId: "simpleCleaningInfoSelectedDates",
      previewId: "simpleCleaningInfoPreview"
    },
    {
      label: "部屋変更",
      maxDates: 1,
      printKey: "printRcInfo",
      inputId: "printRcInfo",
      textKey: "rcInfoCustomText",
      panelId: "rcInfoPicker",
      openButtonId: "openRcInfoPicker",
      monthId: "rcInfoMonth",
      calendarId: "rcInfoCalendar",
      clearButtonId: "clearRcInfoDates",
      selectedId: "rcInfoSelectedDates",
      previewId: "rcInfoPreview"
    },
    {
      label: "部屋変更",
      maxDates: 1,
      printKey: "printRcInfo",
      inputId: "simplePrintRcInfo",
      textKey: "rcInfoCustomText",
      panelId: "simpleRcInfoPicker",
      openButtonId: "simpleOpenRcInfoPicker",
      monthId: "simpleRcInfoMonth",
      calendarId: "simpleRcInfoCalendar",
      clearButtonId: "simpleClearRcInfoDates",
      selectedId: "simpleRcInfoSelectedDates",
      previewId: "simpleRcInfoPreview"
    }
  ];
}

function getDefaultPickerMonth() {
  const date = records.find((record) => record.arrivalDate)?.arrivalDate || new Date();
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function renderDatePicker(config) {
  const calendar = document.getElementById(config.calendarId);
  const monthInput = document.getElementById(config.monthId);
  const monthValue = monthInput.value || getDefaultPickerMonth();
  monthInput.value = monthValue;
  const [year, month] = monthValue.split("-").map(Number);
  if (!year || !month) return;

  const selected = new Set(parseDateInfoDates(settings[config.textKey]));
  const firstDay = new Date(year, month - 1, 1).getDay();
  const daysInMonth = new Date(year, month, 0).getDate();
  const cells = ["日", "月", "火", "水", "木", "金", "土"].map((day) => `<span class="calendar-weekday">${day}</span>`);

  for (let index = 0; index < firstDay; index += 1) {
    cells.push('<span class="calendar-empty"></span>');
  }

  for (let day = 1; day <= daysInMonth; day += 1) {
    const dateText = `${month}/${day}`;
    const active = selected.has(dateText) ? " active" : "";
    cells.push(`<button type="button" class="calendar-day${active}" data-date="${dateText}">${day}</button>`);
  }

  calendar.innerHTML = cells.join("");
  calendar.querySelectorAll(".calendar-day").forEach((button) => {
    button.addEventListener("click", () => {
      toggleSavedDate(config, button.dataset.date);
      bindCleaningSettingsToForm();
      renderTable();
    });
  });
}

function toggleSavedDate(config, dateText) {
  const dates = new Set(parseDateInfoDates(settings[config.textKey]));
  if (dates.has(dateText)) {
    dates.delete(dateText);
  } else {
    if (config.maxDates === 1) {
      dates.clear();
    } else if (dates.size >= config.maxDates) {
      return;
    }
    dates.add(dateText);
  }
  const sortedDates = sortMonthDayDates([...dates]);
  const prefix = config.label === "清掃" ? "清掃日：" : "部屋変更：";
  settings[config.textKey] = sortedDates.length ? `${prefix}${sortedDates.join("、")}` : "";
}

function parseDateInfoDates(text) {
  return String(text || "")
    .replace(/^(清掃日：|清掃日:|清掃|部屋変更：|部屋変更:|R\/C|部屋変更)\s*/, "")
    .split(/[、,\s]+/)
    .map((date) => date.trim())
    .filter((date) => /^\d{1,2}\/\d{1,2}$/.test(date));
}

function sortMonthDayDates(dates) {
  return [...new Set(dates)].sort((a, b) => {
    const [monthA, dayA] = a.split("/").map(Number);
    const [monthB, dayB] = b.split("/").map(Number);
    return monthA - monthB || dayA - dayB;
  });
}

function formatSelectedDates(text) {
  const dates = parseDateInfoDates(text);
  return dates.length ? `選択済み: ${dates.join("、")}` : "日付未選択";
}

function getDateInfoText(config) {
  return String(settings[config.textKey] || "").trim();
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
  updatePrintOrderFromActiveForm();
  bindSimplePositionSettingsToForm();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(pickSettings(settings)));
  setStatus("設定をブラウザに保存しました。");
}

function saveSimpleSettings() {
  updateSettingsFromSimpleForm();
  updatePrintOrderFromActiveForm();
  updateSettingsFromSimplePositionForm();
  bindSettingsToForm();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(pickSettings(settings)));
  setSimpleStatus("文字サイズをブラウザに保存しました。");
}

function resetSettings() {
  resetFontSizesToDefault();
  bindSettingsToForm();
  bindSimpleSettingsToForm();
  bindSimplePositionSettingsToForm();
  updatePreviewChips();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(pickSettings(settings)));
  setStatus("文字サイズを初期値に戻しました。");
}

function resetSimpleSettings() {
  resetFontSizesToDefault();
  bindSettingsToForm();
  bindSimpleSettingsToForm();
  bindSimplePositionSettingsToForm();
  updatePreviewChips();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(pickSettings(settings)));
  setSimpleStatus("文字サイズを初期値に戻しました。");
}

function resetFontSizesToDefault() {
  FONT_SIZE_KEYS.forEach((key) => {
    settings[key] = DEFAULT_SETTINGS[key];
  });
}

function pickSettings(source) {
  return settingInputs.reduce((picked, key) => {
    picked[key] = source[key];
    return picked;
  }, {
    printCopyMode: source.printCopyMode,
    printOrder: source.printOrder === "room" ? "room" : "load",
    cleaningInfoCustomText: source.cleaningInfoCustomText,
    rcInfoCustomText: source.rcInfoCustomText
  });
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
    resetPrintSelection();
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

    const firstGroup = records.find((r) => r.groupName)?.groupName || "";
    if (els.simpleGroupNameInput) els.simpleGroupNameInput.value = firstGroup;
    if (els.advancedGroupNameInput) els.advancedGroupNameInput.value = firstGroup;

    resetPrintSelection();
    renderTable();
    document.getElementById("simpleNextUpload").disabled = false;
    const summary = getRecordsSummary(records);
    setStatus(`「${sheetName}」から ${summary}のデータを読み込みました。${warnings.length ? `除外・警告 ${warnings.length}件があります。` : ""}`);
    setSimpleStatus(`「${sheetName}」から ${summary}のデータを読み込みました。${warnings.length ? `除外・警告 ${warnings.length}件があります。` : ""}`);
  } catch (error) {
    records = [];
    warnings = [];
    resetPrintSelection();
    renderTable();
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
    resetPrintSelection();
    renderTable();
    setStatus("有効な直接入力データがありません。", true);
    return;
  }

  const firstGroup = records.find((r) => r.groupName)?.groupName || "";
  if (els.simpleGroupNameInput) els.simpleGroupNameInput.value = firstGroup;
  if (els.advancedGroupNameInput) els.advancedGroupNameInput.value = firstGroup;

  resetPrintSelection();
  renderTable();
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
    resetPrintSelection();
    renderTable();
    setSimpleStatus("有効な直接入力データがありません。", true);
    return;
  }

  const firstGroup = records.find((r) => r.groupName)?.groupName || "";
  if (els.simpleGroupNameInput) els.simpleGroupNameInput.value = firstGroup;
  if (els.advancedGroupNameInput) els.advancedGroupNameInput.value = firstGroup;

  resetPrintSelection();
  renderTable();
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
  const arrivalDate = parseMonthDayDate(stayInfo);

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
      stayInfo,
      arrivalDate
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
    const arrivalDate = parseExcelArrivalDate(arrivalCell);

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
      stayInfo: formatStayInfo(arrivalCell, nights),
      arrivalDate
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
  const fullDateMatch = text.match(/\d{4}[/-](\d{1,2})[/-](\d{1,2})/);
  if (fullDateMatch) return `${Number(fullDateMatch[1])}/${Number(fullDateMatch[2])}`;
  const match = text.match(/(\d{1,2})[/-](\d{1,2})/);
  return match ? `${Number(match[1])}/${Number(match[2])}` : text;
}

function parseExcelArrivalDate(cell) {
  if (!cell) return null;
  if (typeof cell.v === "number" && Number.isFinite(cell.v)) {
    const date = XLSX.SSF.parse_date_code(cell.v);
    if (date) return new Date(date.y, date.m - 1, date.d);
  }
  return parseMonthDayDate(normalizeCell(cell));
}

function parseMonthDayDate(value) {
  const text = String(value || "");
  const fullDateMatch = text.match(/(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
  const monthDayMatch = text.match(/(\d{1,2})[/-](\d{1,2})/);
  const year = fullDateMatch ? Number(fullDateMatch[1]) : new Date().getFullYear();
  const month = Number(fullDateMatch ? fullDateMatch[2] : monthDayMatch?.[1]);
  const day = Number(fullDateMatch ? fullDateMatch[3] : monthDayMatch?.[2]);
  if (!month || !day || month > 12 || day > 31) return null;
  return new Date(year, month - 1, day);
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

let editingRecordIndex = null;

function handleEditButtonClick(event) {
  const button = event.target.closest("[data-edit-index]");
  if (!button) return;

  const index = Number(button.dataset.editIndex);
  const record = records.find((r) => r.index === index);
  if (!record) return;

  editingRecordIndex = index;

  // 인쇄 설정에 따른 필드 활성화 상태 정의
  const fields = [
    { id: "editRecRoom", active: settings.printRoom },
    { id: "editRecNames", active: settings.printName },
    { id: "editRecGroupName", active: settings.printGroupName },
    { id: "editRecStayInfo", active: settings.printStayInfo && settings.printStaySchedule },
    { id: "editRecCleaning", active: settings.printStayInfo && settings.printCleaningInfo },
    { id: "editRecRc", active: settings.printStayInfo && settings.printRcInfo }
  ];

  fields.forEach(f => {
    const input = document.getElementById(f.id);
    const label = input.closest("label");

    if (f.active) {
      input.disabled = false;
      label.style.color = ""; // 기본색
      label.style.opacity = "1";
      input.style.backgroundColor = ""; // 기본색
    } else {
      input.disabled = true;
      label.style.color = "#999"; // 회색 텍스트
      label.style.opacity = "0.6";
      input.style.backgroundColor = "#f0f0f0"; // 회색 배경
    }
  });

  document.getElementById("editRecRoom").value = record.room || "";
  document.getElementById("editRecNames").value = record.outputNames.join("/");
  document.getElementById("editRecGroupName").value = record.groupName || "";
  document.getElementById("editRecStayInfo").value = record.stayInfo || "";
  document.getElementById("editRecCleaning").value = record.customCleaning || getCleaningInfo(record) || "";
  document.getElementById("editRecRc").value = record.customRc || getRcInfo(record) || "";

  els.recordEditModal.showModal();
}

function saveIndividualRecordEdit() {
  const record = records.find((r) => r.index === editingRecordIndex);
  if (!record) return;

  if (!document.getElementById("editRecRoom").disabled) record.room = document.getElementById("editRecRoom").value.trim();
  if (!document.getElementById("editRecNames").disabled) {
    const namesText = document.getElementById("editRecNames").value.trim();
    record.outputNames = namesText.split("/").map((n) => n.trim()).filter(Boolean);
    record.rawName = record.outputNames.join(" / ");
    record.outputName = record.rawName;
  }
  if (!document.getElementById("editRecGroupName").disabled) record.groupName = document.getElementById("editRecGroupName").value.trim();
  if (!document.getElementById("editRecStayInfo").disabled) record.stayInfo = document.getElementById("editRecStayInfo").value.trim();
  if (!document.getElementById("editRecCleaning").disabled) record.customCleaning = document.getElementById("editRecCleaning").value.trim();
  if (!document.getElementById("editRecRc").disabled) record.customRc = document.getElementById("editRecRc").value.trim();

  renderTable();
  els.recordEditModal.close();
}

function getCleaningInfo(record) {
  if (record.customCleaning !== undefined) return record.customCleaning;
  if (!settings.printStayInfo || !settings.printCleaningInfo) return "";
  return settings.cleaningInfoCustomText.trim();
}

function getRcInfo(record) {
  if (record.customRc !== undefined) return record.customRc;
  if (!settings.printStayInfo || !settings.printRcInfo) return "";
  return settings.rcInfoCustomText.trim();
}

function renderTable() {
  const selectedRecords = getSelectedRecords();
  els.countBadge.textContent = records.length
    ? `${getRecordsSummary(records)} / 選択 ${selectedRecords.length}件`
    : getRecordsSummary(records);
  els.warningBadge.textContent = warnings.length ? `警告 ${warnings.length}件` : "";

  const rows = records.map((record) => {
    const selected = isRecordSelected(record);
    return `
    <tr class="${selected ? "" : "unselected-row"}">
      <td><input type="checkbox" data-print-select-index="${record.index}" ${selected ? "checked" : ""} aria-label="印刷対象"></td>
      <td>${record.index}</td>
      <td>${record.excelRow}</td>
      <td>${escapeHtml(record.room)}</td>
      <td>${escapeHtml(record.rawName)}</td>
      <td>${record.outputNames.length}</td>
      <td>${escapeHtml(record.stayInfo)}</td>
      <td>${escapeHtml(getCleaningInfo(record))}</td>
      <td>${escapeHtml(getRcInfo(record))}</td>
      <td>${escapeHtml(getPrintableNames(record).join(" / "))}</td>
      <td>使用</td>
    </tr>
  `;
  });

  const warningRows = warnings.map((warning) => `
    <tr class="bad-row">
      <td>-</td>
      <td>-</td>
      <td>${warning.excelRow}</td>
      <td>${escapeHtml(warning.room)}</td>
      <td>${escapeHtml(warning.rawName)}</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>${warning.reason}</td>
    </tr>
  `);

  els.dataTable.innerHTML = rows.concat(warningRows).join("") || '<tr><td colspan="11" class="empty">Excelファイルを選択するとデータが表示されます。</td></tr>';
  updatePrintSelectionPanels(selectedRecords);
  updateTableSelectionControl(selectedRecords);
  updateCopyModeVisibility();
  updateCleaningControls();
}

function updatePrintSelectionPanels(selectedRecords = getSelectedRecords()) {
  const panels = [els.simplePrintSelection, els.advancedPrintSelection].filter(Boolean);
  const content = buildPrintSelectionPanelHtml(selectedRecords);
  panels.forEach((panel) => {
    panel.innerHTML = content;
  });
}

function buildPrintSelectionPanelHtml(selectedRecords) {
  if (!records.length) {
    return '<p class="empty compact-empty">データを読み込むと、ここで印刷する行を選択できます。</p>';
  }

  const orderedRecords = getRecordsInDisplayOrder();
  const rows = orderedRecords.map((record, position) => {
    const selected = isRecordSelected(record);
    const title = [
      record.room ? `部屋 ${record.room}` : "",
      record.rawName
    ].filter(Boolean).join(" / ");
    return `
      <label class="selection-row ${selected ? "" : "unselected-row"}">
        <input type="checkbox" data-print-select-index="${record.index}" ${selected ? "checked" : ""}>
        <span class="selection-index">${position + 1}</span>
        <span class="selection-main">${escapeHtml(title || "印刷 データ")}</span>
        <small>Excel行 ${escapeHtml(record.excelRow)}</small>
        <button type="button" class="secondary edit-trigger-button" data-edit-index="${record.index}">修正</button>
      </label>
    `;
  }).join("");

  return `
    <div class="selection-toolbar">
      <strong>選択中 ${selectedRecords.length}件 / ${records.length}件</strong>
      <span class="selection-actions">
        <button type="button" class="text-button" data-selection-action="all">全て選択</button>
        <button type="button" class="text-button" data-selection-action="none">全て解除</button>
      </span>
    </div>
    <div class="print-selection-list">${rows}</div>
  `;
}

function updateTableSelectionControl(selectedRecords) {
  if (!els.selectAllRows) return;
  els.selectAllRows.checked = records.length > 0 && selectedRecords.length === records.length;
  els.selectAllRows.indeterminate = selectedRecords.length > 0 && selectedRecords.length < records.length;
  els.selectAllRows.disabled = !records.length;
}

function ensureData() {
  if (records.length) return true;
  setStatus("先にExcelファイルを選択してください。", true);
  setSimpleStatus("先にExcelファイルを選択してください。", true);
  return false;
}

function getRecordsForMode(mode) {
  if (mode === "test") {
    const sorted = sortRecordsForPrint(records);
    return sorted.length ? [sorted[0]] : [];
  }
  if (mode === "all") return sortRecordsForPrint(records);

  const selected = getSelectedRecords();
  if (!selected.length) {
    setStatus("印刷するデータを選択してください。", true);
    setSimpleStatus("印刷するデータを選択してください。", true);
    return [];
  }

  return sortRecordsForPrint(selected);
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
  updateCopyModeFromActiveForm();
  updatePrintOrderFromActiveForm();
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
    .stay-info,
    .cleaning-info,
    .rc-info {
      position: absolute;
      margin: 0;
      padding: 0;
      color: #000;
      line-height: 1;
      white-space: nowrap;
      font-weight: 700;
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
    .cleaning-info {
      left: ${printX(settings.cleaningInfoX)}mm;
      top: ${printY(settings.cleaningInfoY)}mm;
      max-width: ${getPrintFaceRemainingWidth(settings.cleaningInfoX)}mm;
      font-size: ${settings.cleaningInfoFontSize}pt;
      line-height: 1.15;
      white-space: normal;
      overflow-wrap: anywhere;
      transform: ${printTransform()};
    }
    .rc-info {
      left: ${printX(settings.rcInfoX)}mm;
      top: ${printY(settings.rcInfoY)}mm;
      max-width: ${getPrintFaceRemainingWidth(settings.rcInfoX)}mm;
      font-size: ${settings.rcInfoFontSize}pt;
      line-height: 1.15;
      white-space: normal;
      overflow-wrap: anywhere;
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
  const stayInfo = settings.printStayInfo && settings.printStaySchedule && record.stayInfo
    ? `<p class="stay-info" style="top:${printY(nameLayout.stayInfoY)}mm;">${escapeHtml(record.stayInfo)}</p>`
    : "";
  const cleaningInfoText = getCleaningInfo(record);
  const cleaningInfo = cleaningInfoText
    ? `<p class="cleaning-info" style="top:${printY(nameLayout.cleaningInfoY)}mm;">${escapeHtml(cleaningInfoText)}</p>`
    : "";
  const rcInfoText = getRcInfo(record);
  const rcInfo = rcInfoText
    ? `<p class="rc-info" style="top:${printY(nameLayout.rcInfoY)}mm;">${escapeHtml(rcInfoText)}</p>`
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
    ${cleaningInfo}
    ${rcInfo}
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










  const safeGap = 2; // 최소 안전 거리 2mm

  // 방 번호와 성함 사이 거리 확보
  const projectedLastNameY = Number(settings.nameY) + (count - 1) * lineGap;
  const roomTopY = Number(settings.roomY) - (Number(settings.roomFontSize) * PT_TO_MM);
  const shiftUp = Math.max(0, (projectedLastNameY + safeGap) - roomTopY);
  return {
    fontSize,
    lineGap,




    firstNameY: Number(settings.nameY) - shiftUp,
    stayInfoY: Number(settings.stayInfoY) - shiftUp,
    cleaningInfoY: Number(settings.cleaningInfoY) - shiftUp,
    rcInfoY: Number(settings.rcInfoY) - shiftUp
  };
}

function getPrintableNames(record) {
  if (!settings.printName) return [];
  return record.outputNames.map(formatGuestNameForPrint);
}

function getAutoNameFontSize(names) {
  const baseSize = Number(settings.nameFontSize);
  const remainingWidth = getPrintFaceRemainingWidth(settings.nameX); // 실제 남은 폭 계산

  const longestWidth = names.reduce((max, name) => (
    Math.max(max, estimateTextWidthMm(name, baseSize))
  ), 0);




  if (longestWidth <= remainingWidth) return baseSize;
  const fittedSize = baseSize * (remainingWidth / longestWidth);
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

function getPrintFaceStartX() {
  return (B6_WIDTH_MM - Number(settings.pageWidthMm)) / 2 + Number(settings.printAreaStartX);
}

function getPrintFaceEndX() {
  return getPrintFaceStartX() + Number(settings.pageWidthMm) - Number(settings.printAreaStartX);
}

function getPrintFaceRemainingWidth(baseX) {
  const x = Number(baseX) + Number(settings.globalOffsetX);
  const width = settings.rotate180
    ? x - getPrintFaceStartX()
    : getPrintFaceEndX() - x;
  return roundToTenth(Math.max(MIN_WRAPPED_TEXT_WIDTH_MM, width));
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
