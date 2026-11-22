// ====================================================================
// PEMS-UZ: PROJECT EQUIPMENT MANAGEMENT SYSTEM – UZBEKISTAN
// Main Application Controller, Reactive State, Lifecycle Engine & Reports
// Developed by: Abdiyev Salimkhon Sanjar ugli
// ====================================================================

import {
  SYSTEM_INFO,
  UZBEKISTAN_REGIONS,
  ITEM_TECHNICAL_CONDITIONS,
  INSTALLATION_STATUSES,
  UNITS,
  USER_ROLES,
  DEFAULT_PROJECT_METADATA,
  DEFAULT_TEO_ITEMS,
  DEFAULT_CONTRACTS,
  DEFAULT_RECEIPTS_UZASBO,
  DEFAULT_HANDOVER_RECORDS,
  DEFAULT_INSTALLATIONS,
  DEFAULT_MONITORING_RECORDS,
  DEFAULT_INVENTORY_ITEMS,
  DEFAULT_COMMITTEE_DATA,
  DEFAULT_AUDIT_LOGS,
  TRANSLATIONS
} from "./data.js";
import { exportToExcel, parseExcelFile } from "./excelExporter.js";

// Storage Key
const STORAGE_KEY = "pems_uz_state_v3";

// Current Active State
let appState = loadState();
let currentLang = appState.currentLang || "ru";
let currentRole = appState.currentRole || "admin";

function loadState() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY) || localStorage.getItem("pems_uz_inventory_state_v2");
    if (saved) {
      const parsed = JSON.parse(saved);
      return {
        projectMetadata: { ...DEFAULT_PROJECT_METADATA, ...(parsed.projectMetadata || {}) },
        teoItems: parsed.teoItems && parsed.teoItems.length ? parsed.teoItems : DEFAULT_TEO_ITEMS,
        contracts: parsed.contracts && parsed.contracts.length ? parsed.contracts : DEFAULT_CONTRACTS,
        receipts: parsed.receipts && parsed.receipts.length ? parsed.receipts : DEFAULT_RECEIPTS_UZASBO,
        handovers: parsed.handovers && parsed.handovers.length ? parsed.handovers : DEFAULT_HANDOVER_RECORDS,
        installations: parsed.installations && parsed.installations.length ? parsed.installations : DEFAULT_INSTALLATIONS,
        monitoringRecords: parsed.monitoringRecords && parsed.monitoringRecords.length ? parsed.monitoringRecords : DEFAULT_MONITORING_RECORDS,
        inventoryItems: parsed.inventoryItems && parsed.inventoryItems.length ? parsed.inventoryItems : DEFAULT_INVENTORY_ITEMS,
        committee: { ...DEFAULT_COMMITTEE_DATA, ...(parsed.committee || {}) },
        auditLogs: parsed.auditLogs && parsed.auditLogs.length ? parsed.auditLogs : DEFAULT_AUDIT_LOGS,
        currentLang: parsed.currentLang || "ru",
        currentRole: parsed.currentRole || "admin"
      };
    }
  } catch (e) {
    console.error("Failed to load state from localStorage:", e);
  }

  return {
    projectMetadata: { ...DEFAULT_PROJECT_METADATA },
    teoItems: JSON.parse(JSON.stringify(DEFAULT_TEO_ITEMS)),
    contracts: JSON.parse(JSON.stringify(DEFAULT_CONTRACTS)),
    receipts: JSON.parse(JSON.stringify(DEFAULT_RECEIPTS_UZASBO)),
    handovers: JSON.parse(JSON.stringify(DEFAULT_HANDOVER_RECORDS)),
    installations: JSON.parse(JSON.stringify(DEFAULT_INSTALLATIONS)),
    monitoringRecords: JSON.parse(JSON.stringify(DEFAULT_MONITORING_RECORDS)),
    inventoryItems: JSON.parse(JSON.stringify(DEFAULT_INVENTORY_ITEMS)),
    committee: JSON.parse(JSON.stringify(DEFAULT_COMMITTEE_DATA)),
    auditLogs: JSON.parse(JSON.stringify(DEFAULT_AUDIT_LOGS)),
    currentLang: "ru",
    currentRole: "admin"
  };
}

function saveState() {
  try {
    appState.currentLang = currentLang;
    appState.currentRole = currentRole;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(appState));
  } catch (e) {
    console.error("Failed to save state:", e);
  }
}

// Audit Logger (ТЗ Section 11)
function logAudit(moduleName, action, entityId, details, oldValue = "-", newValue = "-") {
  const now = new Date();
  const timeStr = now.toISOString().replace("T", " ").substring(0, 19);
  const roleObj = USER_ROLES.find(r => r.id === currentRole) || USER_ROLES[0];
  const entry = {
    id: "log-" + Date.now(),
    timestamp: timeStr,
    userId: currentRole,
    userName: roleObj.nameRu,
    module: moduleName,
    action: action,
    entityId: String(entityId || ""),
    details: details,
    oldValue: String(oldValue),
    newValue: String(newValue)
  };
  appState.auditLogs.unshift(entry);
  if (appState.auditLogs.length > 200) appState.auditLogs.pop();
  saveState();
  renderAuditLogs();
}

// Formatters
function formatNumber(num) {
  if (num === undefined || num === null || isNaN(num)) return "0";
  return Number(num).toLocaleString("ru-RU");
}

function formatCurrency(num, curr = "USD") {
  const val = formatNumber(num);
  return curr === "USD" ? `$${val}` : `${val} UZS`;
}

// UI Elements Map
const dom = {
  // Navigation Tabs
  tabBtns: document.querySelectorAll(".tab-btn"),
  tabContents: document.querySelectorAll(".tab-content"),

  // Language & Role
  langPills: document.querySelectorAll(".lang-pill"),
  roleSelect: document.getElementById("user-role-select"),
  currentRoleBadge: document.getElementById("current-role-badge"),

  // Meta Inputs
  metaPresidential: document.getElementById("meta-presidential"),
  metaTeo: document.getElementById("meta-teo"),
  metaOrderNo: document.getElementById("meta-order-no"),
  metaInvDate: document.getElementById("meta-inv-date"),
  metaUsdRate: document.getElementById("meta-usd-rate"),

  // Top Buttons
  btnExportExcel: document.getElementById("btn-export-excel"),
  excelFileInput: document.getElementById("excel-file-input"),
  btnAddItemModal: document.getElementById("btn-add-item-modal"),
  btnResetData: document.getElementById("btn-reset-data"),

  // Dashboard Pipeline & KPIs
  pipeTeo: document.getElementById("pipe-val-teo"),
  pipeContracts: document.getElementById("pipe-val-contracts"),
  pipeReceipts: document.getElementById("pipe-val-receipts"),
  pipeTransferred: document.getElementById("pipe-val-transferred"),
  pipeInstalled: document.getElementById("pipe-val-installed"),
  pipeCommissioned: document.getElementById("pipe-val-commissioned"),
  pipeMonitored: document.getElementById("pipe-val-monitored"),
  pipeInventoried: document.getElementById("pipe-val-inventoried"),

  kpiTeoApproved: document.getElementById("kpi-teo-approved"),
  kpiContracted: document.getElementById("kpi-contracted"),
  kpiDelivered: document.getElementById("kpi-delivered"),
  kpiTransferred: document.getElementById("kpi-transferred"),
  kpiInstalled: document.getElementById("kpi-installed"),
  kpiCommissioned: document.getElementById("kpi-commissioned"),
  kpiMonitored: document.getElementById("kpi-monitored"),
  kpiDefects: document.getElementById("kpi-defects"),
  kpiInventoried: document.getElementById("kpi-inventoried"),
  kpiRemaining: document.getElementById("kpi-remaining"),
  dashboardAlertsTbody: document.getElementById("dashboard-alerts-tbody"),
  btnPrintAlerts: document.getElementById("btn-print-alerts"),

  // Badges
  badgeTeoCount: document.getElementById("badge-teo-count"),
  badgeContractsCount: document.getElementById("badge-contracts-count"),
  badgeReceiptsCount: document.getElementById("badge-receipts-count"),
  badgeHandoversCount: document.getElementById("badge-handovers-count"),
  badgeInstallCount: document.getElementById("badge-install-count"),
  badgeMonitoringCount: document.getElementById("badge-monitoring-count"),
  badgeInventoryCount: document.getElementById("badge-inventory-count"),

  // Tables
  teoTbody: document.getElementById("teo-tbody"),
  searchTeo: document.getElementById("search-teo"),
  filterTeoRegion: document.getElementById("filter-teo-region"),
  btnAddTeoRow: document.getElementById("btn-add-teo-row"),
  btnPrintTeo: document.getElementById("btn-print-teo"),

  contractsTbody: document.getElementById("contracts-tbody"),
  searchContracts: document.getElementById("search-contracts"),
  btnAddContract: document.getElementById("btn-add-contract"),
  btnPrintContracts: document.getElementById("btn-print-contracts"),

  receiptsTbody: document.getElementById("receipts-tbody"),
  searchReceipts: document.getElementById("search-receipts"),
  btnAddReceipt: document.getElementById("btn-add-receipt"),
  btnPrintReceipts: document.getElementById("btn-print-receipts"),

  handoversTbody: document.getElementById("handovers-tbody"),
  searchHandovers: document.getElementById("search-handovers"),
  filterHandoverRegion: document.getElementById("filter-handover-region"),
  btnAddHandoverNew: document.getElementById("btn-add-handover-new"),
  btnPrintHandoversList: document.getElementById("btn-print-handovers-list"),

  installationsTbody: document.getElementById("installations-tbody"),
  searchInstallations: document.getElementById("search-installations"),
  filterInstallStatus: document.getElementById("filter-install-status"),
  btnAddInstallation: document.getElementById("btn-add-installation"),
  btnPrintInstallations: document.getElementById("btn-print-installations"),

  monitoringTbody: document.getElementById("monitoring-tbody"),
  searchMonitoring: document.getElementById("search-monitoring"),
  filterMonCondition: document.getElementById("filter-mon-condition"),
  btnAddMonitoringVisit: document.getElementById("btn-add-monitoring-visit"),
  btnPrintMonitoring: document.getElementById("btn-print-monitoring"),

  inventoryTbody: document.getElementById("inventory-tbody"),
  searchInventory: document.getElementById("search-inventory"),
  filterInvRegion: document.getElementById("filter-inv-region"),
  filterInvDiscrepancy: document.getElementById("filter-inv-discrepancy"),
  btnAddInvItem: document.getElementById("btn-add-inv-item"),
  btnShowCommitteeProtocol: document.getElementById("btn-show-committee-protocol"),
  btnPrintInvSheet: document.getElementById("btn-print-inv-sheet"),

  // Reports
  selectReportType: document.getElementById("select-report-type"),
  reportViewContainer: document.getElementById("report-view-container"),
  btnPrintActiveReport: document.getElementById("btn-print-active-report"),

  // Module 9
  quickPassportInput: document.getElementById("quick-passport-input"),
  btnQuickPassportFind: document.getElementById("btn-quick-passport-find"),
  passportQuickChips: document.getElementById("passport-quick-chips"),
  rolesMatrixInfo: document.getElementById("roles-matrix-info"),
  auditLogTbody: document.getElementById("audit-log-tbody"),

  // Modals
  passportModal: document.getElementById("passport-modal"),
  passportModalBody: document.getElementById("passport-modal-body"),
  btnClosePassportModal: document.getElementById("btn-close-passport-modal"),
  btnClosePassportBtn: document.getElementById("btn-close-passport-btn"),
  btnPrintPassport: document.getElementById("btn-print-passport"),

  itemModal: document.getElementById("item-modal"),
  modalTargetModule: document.getElementById("modal-target-module"),
  dynamicFormFields: document.getElementById("dynamic-form-fields"),
  btnCloseItemModal: document.getElementById("btn-close-item-modal"),
  btnCancelModal: document.getElementById("btn-cancel-modal"),
  itemForm: document.getElementById("item-form"),

  printModal: document.getElementById("print-modal"),
  printModalTitle: document.getElementById("print-modal-title"),
  printableSheetContent: document.getElementById("printable-sheet-content"),
  btnClosePrintModal: document.getElementById("btn-close-print-modal"),
  btnExecutePrint: document.getElementById("btn-execute-print"),

  toast: document.getElementById("toast")
};

// ====================================================================
// INITIALIZATION
// ====================================================================
document.addEventListener("DOMContentLoaded", () => {
  initTabs();
  initLanguageSwitcher();
  initRoleSwitcher();
  populateDropdowns();
  bindMetaInputs();
  bindActionEvents();
  renderAll();
});

function initTabs() {
  dom.tabBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      const targetTab = btn.getAttribute("data-tab");
      dom.tabBtns.forEach(b => b.classList.remove("active"));
      dom.tabContents.forEach(c => c.classList.remove("active"));
      btn.classList.add("active");
      const targetSec = document.getElementById(targetTab);
      if (targetSec) targetSec.classList.add("active");
    });
  });
}

function initLanguageSwitcher() {
  dom.langPills.forEach(pill => {
    pill.addEventListener("click", () => {
      currentLang = pill.getAttribute("data-lang");
      dom.langPills.forEach(p => p.classList.remove("active"));
      pill.classList.add("active");
      saveState();
      applyLanguage();
      renderAll();
      showToast(currentLang === "uz" ? "Til o'zgartirildi: O'zbekcha" : currentLang === "ru" ? "Язык переключён: Русский" : "Language switched: English");
    });
  });
}

function applyLanguage() {
  const t = TRANSLATIONS[currentLang] || TRANSLATIONS.ru;
  const brandMain = document.getElementById("ui-brand-main");
  const brandSub = document.getElementById("ui-brand-sub");
  if (brandMain) brandMain.textContent = t.systemTitle;
  if (brandSub) brandSub.textContent = t.systemSubtitle;

  const btnExport = document.getElementById("ui-btn-export");
  const btnImport = document.getElementById("ui-btn-import");
  const btnAdd = document.getElementById("ui-btn-add");
  const btnReset = document.getElementById("ui-btn-reset");
  if (btnExport) btnExport.textContent = t.btnExportExcel;
  if (btnImport) btnImport.textContent = t.btnImportExcel;
  if (btnAdd) btnAdd.textContent = t.btnAddItem;
  if (btnReset) btnReset.textContent = t.btnReset;
}

function initRoleSwitcher() {
  if (dom.roleSelect) {
    dom.roleSelect.value = currentRole;
    dom.roleSelect.addEventListener("change", (e) => {
      currentRole = e.target.value;
      saveState();
      updateRoleDisplay();
      logAudit("Пользователи и роли", "ROLE_SWITCH", currentRole, `Faol rol o'zgartirildi: ${currentRole}`);
      showToast(`Роль переключена: ${currentRole.toUpperCase()}`);
    });
    updateRoleDisplay();
  }
}

function updateRoleDisplay() {
  const roleObj = USER_ROLES.find(r => r.id === currentRole) || USER_ROLES[0];
  if (dom.currentRoleBadge) {
    dom.currentRoleBadge.textContent = roleObj.badge;
    dom.currentRoleBadge.style.backgroundColor = roleObj.color;
  }
  renderRolesMatrix();
}

function populateDropdowns() {
  // Regions
  const regSelects = [dom.filterTeoRegion, dom.filterHandoverRegion, dom.filterInvRegion];
  regSelects.forEach(sel => {
    if (!sel) return;
    sel.innerHTML = '<option value="">Barcha hududlar / Все регионы</option>';
    UZBEKISTAN_REGIONS.forEach(reg => {
      const opt = document.createElement("option");
      opt.value = reg;
      opt.textContent = reg;
      sel.appendChild(opt);
    });
  });

  // Installation statuses
  if (dom.filterInstallStatus) {
    dom.filterInstallStatus.innerHTML = '<option value="">Все статусы монтажа</option>';
    INSTALLATION_STATUSES.forEach(st => {
      const opt = document.createElement("option");
      opt.value = st;
      opt.textContent = st;
      dom.filterInstallStatus.appendChild(opt);
    });
  }

  // Technical conditions
  if (dom.filterMonCondition) {
    dom.filterMonCondition.innerHTML = '<option value="">Все категории техсостояния</option>';
    ITEM_TECHNICAL_CONDITIONS.forEach(tc => {
      const opt = document.createElement("option");
      opt.value = tc;
      opt.textContent = tc;
      dom.filterMonCondition.appendChild(opt);
    });
  }
}

function bindMetaInputs() {
  if (dom.metaPresidential) {
    dom.metaPresidential.value = appState.projectMetadata.presidentialResolution || "";
    dom.metaPresidential.addEventListener("input", (e) => {
      appState.projectMetadata.presidentialResolution = e.target.value;
      saveState();
    });
  }
  if (dom.metaTeo) {
    dom.metaTeo.value = appState.projectMetadata.teoNumber || "";
    dom.metaTeo.addEventListener("input", (e) => {
      appState.projectMetadata.teoNumber = e.target.value;
      saveState();
    });
  }
  if (dom.metaOrderNo) {
    dom.metaOrderNo.value = appState.projectMetadata.orderNo || "";
    dom.metaOrderNo.addEventListener("input", (e) => {
      appState.projectMetadata.orderNo = e.target.value;
      saveState();
    });
  }
  if (dom.metaInvDate) {
    dom.metaInvDate.value = appState.projectMetadata.inventoryDate || "";
    dom.metaInvDate.addEventListener("change", (e) => {
      appState.projectMetadata.inventoryDate = e.target.value;
      saveState();
    });
  }
  if (dom.metaUsdRate) {
    dom.metaUsdRate.value = appState.projectMetadata.usdRate || 12850;
    dom.metaUsdRate.addEventListener("input", (e) => {
      appState.projectMetadata.usdRate = Number(e.target.value) || 12850;
      saveState();
      renderAll();
    });
  }
}

function bindActionEvents() {
  // Export Excel
  dom.btnExportExcel.addEventListener("click", () => {
    exportToExcel(appState);
    logAudit("Экспорт", "EXPORT_EXCEL", "PEMS_UZ_Equipment_Registry", "Excel fayli yuklab olindi (.xlsx)");
    showToast("Excel fayli (.xlsx) muvaffaqiyatli yuklab olindi!");
  });

  // Import Excel
  dom.excelFileInput.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const parsed = await parseExcelFile(file);
      if (parsed.items && parsed.items.length > 0) {
        appState.inventoryItems = parsed.items;
        saveState();
        renderAll();
        logAudit("Импорт", "IMPORT_EXCEL", file.name, `${parsed.items.length} ta uskuna Exceldan import qilindi`);
        showToast(`${parsed.items.length} ta yozuv Exceldan yuklandi!`);
      } else {
        alert("Excel faylidan yaroqli ma'lumot topilmadi.");
      }
    } catch (err) {
      console.error(err);
      alert("Faylni o'qishda xatolik: " + err.message);
    }
    e.target.value = "";
  });

  // Reset
  dom.btnResetData.addEventListener("click", () => {
    if (confirm("Haqiqatdan ham namunaviy ma'lumotlarni qayta tiklashni istaysizmi? Barcha o'zgarishlar asl holiga qaytadi.")) {
      localStorage.removeItem(STORAGE_KEY);
      appState = loadState();
      renderAll();
      showToast("Barcha modullar namunaviy holatga qaytarildi!");
    }
  });

  // Universal Add Modal
  dom.btnAddItemModal.addEventListener("click", () => openItemModal("inventory"));
  if (dom.btnAddTeoRow) dom.btnAddTeoRow.addEventListener("click", () => openItemModal("teo"));
  if (dom.btnAddContract) dom.btnAddContract.addEventListener("click", () => openItemModal("contract"));
  if (dom.btnAddReceipt) dom.btnAddReceipt.addEventListener("click", () => openItemModal("receipt"));
  if (dom.btnAddHandoverNew) dom.btnAddHandoverNew.addEventListener("click", () => openItemModal("handover"));
  if (dom.btnAddInstallation) dom.btnAddInstallation.addEventListener("click", () => openItemModal("installation"));
  if (dom.btnAddMonitoringVisit) dom.btnAddMonitoringVisit.addEventListener("click", () => openItemModal("monitoring"));
  if (dom.btnAddInvItem) dom.btnAddInvItem.addEventListener("click", () => openItemModal("inventory"));

  dom.btnCloseItemModal.addEventListener("click", () => dom.itemModal.classList.add("hidden"));
  dom.btnCancelModal.addEventListener("click", () => dom.itemModal.classList.add("hidden"));
  dom.modalTargetModule.addEventListener("change", (e) => buildDynamicForm(e.target.value));
  dom.itemForm.addEventListener("submit", handleItemFormSubmit);

  // Passport Modal Close
  dom.btnClosePassportModal.addEventListener("click", () => dom.passportModal.classList.add("hidden"));
  dom.btnClosePassportBtn.addEventListener("click", () => dom.passportModal.classList.add("hidden"));
  dom.btnPrintPassport.addEventListener("click", () => {
    window.print();
  });

  // Print Modal Close & Execute
  dom.btnClosePrintModal.addEventListener("click", () => dom.printModal.classList.add("hidden"));
  dom.btnExecutePrint.addEventListener("click", () => window.print());

  // Search & Filter listeners
  [
    [dom.searchTeo, dom.filterTeoRegion, renderTeo],
    [dom.searchContracts, null, renderContracts],
    [dom.searchReceipts, null, renderReceipts],
    [dom.searchHandovers, dom.filterHandoverRegion, renderHandovers],
    [dom.searchInstallations, dom.filterInstallStatus, renderInstallations],
    [dom.searchMonitoring, dom.filterMonCondition, renderMonitoring],
    [dom.searchInventory, dom.filterInvRegion, renderInventory],
    [null, dom.filterInvDiscrepancy, renderInventory]
  ].forEach(([input, sel, func]) => {
    if (input) input.addEventListener("input", func);
    if (sel) sel.addEventListener("change", func);
  });

  // Reports
  if (dom.selectReportType) {
    dom.selectReportType.addEventListener("change", () => renderActiveReport());
  }
  if (dom.btnPrintActiveReport) {
    dom.btnPrintActiveReport.addEventListener("click", () => {
      openPrintPreview("Официальный отчёт PEMS-UZ (АБР / АБИИ)", dom.reportViewContainer.innerHTML);
    });
  }

  // Print button hooks for tables
  if (dom.btnPrintAlerts) dom.btnPrintAlerts.addEventListener("click", () => printAlertsReport());
  if (dom.btnPrintTeo) dom.btnPrintTeo.addEventListener("click", () => printTeoReport());
  if (dom.btnPrintContracts) dom.btnPrintContracts.addEventListener("click", () => printContractsReport());
  if (dom.btnPrintReceipts) dom.btnPrintReceipts.addEventListener("click", () => printReceiptsReport());
  if (dom.btnPrintHandoversList) dom.btnPrintHandoversList.addEventListener("click", () => printHandoversReport());
  if (dom.btnPrintInstallations) dom.btnPrintInstallations.addEventListener("click", () => printInstallationsReport());
  if (dom.btnPrintMonitoring) dom.btnPrintMonitoring.addEventListener("click", () => printMonitoringReport());
  if (dom.btnPrintInvSheet) dom.btnPrintInvSheet.addEventListener("click", () => printInventorySheet());
  if (dom.btnShowCommitteeProtocol) dom.btnShowCommitteeProtocol.addEventListener("click", () => printCommitteeProtocol());

  // Passport Quick Search
  if (dom.btnQuickPassportFind) {
    dom.btnQuickPassportFind.addEventListener("click", () => {
      const q = (dom.quickPassportInput.value || "").trim().toLowerCase();
      if (!q) return;
      const found = findEquipmentByQuery(q);
      if (found) {
        openEquipmentPassport(found);
      } else {
        alert("Bunday seriya raqami yoki inventar kodi bo'yicha uskuna topilmadi.");
      }
    });
  }
}

// ====================================================================
// RENDER ALL
// ====================================================================
function renderAll() {
  renderDashboard();
  renderTeo();
  renderContracts();
  renderReceipts();
  renderHandovers();
  renderInstallations();
  renderMonitoring();
  renderInventory();
  renderActiveReport();
  renderAuditLogs();
  renderRolesMatrix();
  renderQuickPassportChips();
}

// ====================================================================
// 0. DASHBOARD & LIFECYCLE FUNNEL (ТЗ Sections 10.1 & 14)
// ====================================================================
function renderDashboard() {
  const teoApproved = appState.teoItems.reduce((acc, i) => acc + (Number(i.approvedQty) || 0), 0);
  const contracted = appState.contracts.reduce((acc, c) => acc + (Number(c.quantity) || 0), 0);
  const delivered = appState.receipts.reduce((acc, r) => acc + (Number(r.deliveredQuantity) || 0), 0);
  const transferred = appState.handovers.reduce((acc, h) => acc + (Number(h.quantity) || 0), 0);
  const installed = appState.installations.filter(i => (i.status || "").includes("O‘rnatilgan") || (i.status || "").includes("Foydalanishga")).length;
  const commissioned = appState.installations.filter(i => (i.status || "").includes("Foydalanishga topshirilgan")).length;
  const monitored = appState.monitoringRecords.length;
  const defects = appState.monitoringRecords.filter(m => (m.technicalCondition || "").includes("Nuqsonli") || (m.technicalCondition || "").includes("Ta'mir") || (m.technicalCondition || "").includes("Ishlamaydi")).length;
  const inventoried = appState.inventoryItems.reduce((acc, i) => acc + (Number(i.actualQty) || 0), 0);
  const remaining = Math.max(teoApproved - transferred, 0);

  // Pipeline Steps
  if (dom.pipeTeo) dom.pipeTeo.textContent = formatNumber(teoApproved);
  if (dom.pipeContracts) dom.pipeContracts.textContent = formatNumber(contracted);
  if (dom.pipeReceipts) dom.pipeReceipts.textContent = formatNumber(delivered);
  if (dom.pipeTransferred) dom.pipeTransferred.textContent = formatNumber(transferred);
  if (dom.pipeInstalled) dom.pipeInstalled.textContent = formatNumber(installed);
  if (dom.pipeCommissioned) dom.pipeCommissioned.textContent = formatNumber(commissioned);
  if (dom.pipeMonitored) dom.pipeMonitored.textContent = formatNumber(monitored);
  if (dom.pipeInventoried) dom.pipeInventoried.textContent = formatNumber(inventoried);

  // KPI Grid
  if (dom.kpiTeoApproved) dom.kpiTeoApproved.textContent = formatNumber(teoApproved);
  if (dom.kpiContracted) dom.kpiContracted.textContent = formatNumber(contracted);
  if (dom.kpiDelivered) dom.kpiDelivered.textContent = formatNumber(delivered);
  if (dom.kpiTransferred) dom.kpiTransferred.textContent = formatNumber(transferred);
  if (dom.kpiInstalled) dom.kpiInstalled.textContent = formatNumber(installed);
  if (dom.kpiCommissioned) dom.kpiCommissioned.textContent = formatNumber(commissioned);
  if (dom.kpiMonitored) dom.kpiMonitored.textContent = formatNumber(monitored);
  if (dom.kpiDefects) dom.kpiDefects.textContent = formatNumber(defects);
  if (dom.kpiInventoried) dom.kpiInventoried.textContent = formatNumber(inventoried);
  if (dom.kpiRemaining) dom.kpiRemaining.textContent = formatNumber(remaining);

  // Badges
  if (dom.badgeTeoCount) dom.badgeTeoCount.textContent = appState.teoItems.length;
  if (dom.badgeContractsCount) dom.badgeContractsCount.textContent = appState.contracts.length;
  if (dom.badgeReceiptsCount) dom.badgeReceiptsCount.textContent = appState.receipts.length;
  if (dom.badgeHandoversCount) dom.badgeHandoversCount.textContent = appState.handovers.length;
  if (dom.badgeInstallCount) dom.badgeInstallCount.textContent = appState.installations.length;
  if (dom.badgeMonitoringCount) dom.badgeMonitoringCount.textContent = appState.monitoringRecords.length;
  if (dom.badgeInventoryCount) dom.badgeInventoryCount.textContent = appState.inventoryItems.length;

  // Urgent Alerts Table
  renderDashboardAlerts();
}

function renderDashboardAlerts() {
  if (!dom.dashboardAlertsTbody) return;
  const defectiveMonitoring = appState.monitoringRecords.filter(m => 
    (m.technicalCondition || "").includes("Nuqsonli") || 
    (m.technicalCondition || "").includes("Ta'mir") || 
    (m.technicalCondition || "").includes("Ishlamaydi")
  );

  if (defectiveMonitoring.length === 0) {
    dom.dashboardAlertsTbody.innerHTML = '<tr><td colspan="9" style="text-align:center; padding:18px; color:#10b981; font-weight:600;">✓ Hozirda nuqsonli yoki favqulodda chora talab qiluvchi uskunalar yo\'q. Barcha uskunalar soz holatda.</td></tr>';
    return;
  }

  dom.dashboardAlertsTbody.innerHTML = defectiveMonitoring.map((m, idx) => `
    <tr>
      <td style="text-align:center; font-weight:bold;">${idx + 1}</td>
      <td><strong>${m.equipmentName}</strong></td>
      <td class="code-cell" onclick="window.viewPassportBySerial('${m.serialNumber}')">${m.serialNumber}</td>
      <td><code>${m.inventoryNumber || "-"}</code></td>
      <td>${m.monitoredOrg} <br/><small style="color:#64748b;">${m.region}</small></td>
      <td><span class="chip chip-danger">${m.technicalCondition}</span></td>
      <td style="color:#b91c1c; font-weight:500;">${m.identifiedDefects || "-"}</td>
      <td>${m.recommendations || "-"}</td>
      <td>
        <button class="btn-passport" onclick="window.viewPassportBySerial('${m.serialNumber}')">Паспорт</button>
      </td>
    </tr>
  `).join("");
}

// ====================================================================
// 1. TEO DISTRIBUTION (ТЗ Section 3)
// ====================================================================
function renderTeo() {
  if (!dom.teoTbody) return;
  const q = (dom.searchTeo ? dom.searchTeo.value : "").trim().toLowerCase();
  const regFilter = dom.filterTeoRegion ? dom.filterTeoRegion.value : "";

  const filtered = appState.teoItems.filter(item => {
    const matchQuery = !q || 
      (item.itemName || "").toLowerCase().includes(q) ||
      (item.itemCode || "").toLowerCase().includes(q) ||
      (item.receivingOrg || "").toLowerCase().includes(q) ||
      (item.presidentialDoc || "").toLowerCase().includes(q);
    const matchReg = !regFilter || (item.region || "").includes(regFilter);
    return matchQuery && matchReg;
  });

  if (filtered.length === 0) {
    dom.teoTbody.innerHTML = '<tr><td colspan="15" style="text-align:center; padding:16px;">Hech qanday ma\'lumot topilmadi</td></tr>';
    return;
  }

  dom.teoTbody.innerHTML = filtered.map((item, idx) => {
    const rem = Math.max((Number(item.approvedQty) || 0) - (Number(item.actualTransferredQty) || 0), 0);
    return `
      <tr>
        <td style="text-align:center;">${idx + 1}</td>
        <td><code>${item.presidentialDoc || "-"}</code></td>
        <td><code>${item.teoDoc || "-"}</code></td>
        <td class="code-cell" onclick="window.viewPassportByCode('${item.itemCode}')">${item.itemCode}</td>
        <td><strong>${item.itemName}</strong></td>
        <td style="text-align:right; font-weight:bold;">${item.approvedQty}</td>
        <td>${item.region}</td>
        <td>${item.receivingOrg}</td>
        <td style="text-align:right;">${item.plannedQty}</td>
        <td style="text-align:right;">${item.actualDeliveredQty}</td>
        <td style="text-align:right; font-weight:bold; color:#059669;">${item.actualTransferredQty}</td>
        <td style="text-align:right; font-weight:bold; color:${rem > 0 ? '#d97706' : '#64748b'};">${rem}</td>
        <td style="text-align:right;">$${formatNumber(item.unitPriceUsd)}</td>
        <td><span class="chip ${rem === 0 ? 'chip-success' : 'chip-warning'}">${item.status}</span></td>
        <td>
          <button class="btn-passport" onclick="window.viewPassportByCode('${item.itemCode}')">Карточка</button>
        </td>
      </tr>
    `;
  }).join("");
}

// ====================================================================
// 2. CONTRACTS & PROCUREMENT (ТЗ Section 4)
// ====================================================================
function renderContracts() {
  if (!dom.contractsTbody) return;
  const q = (dom.searchContracts ? dom.searchContracts.value : "").trim().toLowerCase();

  const filtered = appState.contracts.filter(c => {
    return !q || 
      (c.contractNo || "").toLowerCase().includes(q) ||
      (c.supplier || "").toLowerCase().includes(q) ||
      (c.goodsDescription || "").toLowerCase().includes(q) ||
      (c.model || "").toLowerCase().includes(q);
  });

  if (filtered.length === 0) {
    dom.contractsTbody.innerHTML = '<tr><td colspan="16" style="text-align:center; padding:16px;">Shartnomalar topilmadi</td></tr>';
    return;
  }

  dom.contractsTbody.innerHTML = filtered.map((c, idx) => `
    <tr>
      <td style="text-align:center;">${idx + 1}</td>
      <td><strong>${c.contractNo}</strong></td>
      <td>${c.contractDate}</td>
      <td>${c.supplier}</td>
      <td><strong>${c.goodsDescription}</strong></td>
      <td><small>${c.componentsIncluded || "-"}</small></td>
      <td style="text-align:center;">${c.yearOfManufacture || "-"}</td>
      <td><code>${c.serialNumberRange || "-"}</code></td>
      <td>${c.manufacturer} (${c.countryOfOrigin})</td>
      <td style="text-align:right; font-weight:bold;">${c.quantity}</td>
      <td style="text-align:right;">$${formatNumber(c.unitPriceUsd)}</td>
      <td style="text-align:right; font-weight:bold;">$${formatNumber(c.totalGoodsValueUsd)}</td>
      <td style="text-align:right;">$${formatNumber((c.ancillaryServicesCostUsd || 0) + (c.domesticLogisticsCostUsd || 0) + (c.airFreightCostUsd || 0))}</td>
      <td style="text-align:right; font-weight:bold; color:var(--navy-primary);">$${formatNumber(c.totalContractAmountUsd)}</td>
      <td><span class="chip chip-info">${c.warrantyPeriod}</span></td>
      <td><span class="chip chip-success">${c.status}</span></td>
    </tr>
  `).join("");
}

// ====================================================================
// 3. RECEIPTS & UZASBO (ТЗ Section 5)
// ====================================================================
function renderReceipts() {
  if (!dom.receiptsTbody) return;
  const q = (dom.searchReceipts ? dom.searchReceipts.value : "").trim().toLowerCase();

  const filtered = appState.receipts.filter(r => {
    return !q || 
      (r.contractNo || "").toLowerCase().includes(q) ||
      (r.invoiceNo || "").toLowerCase().includes(q) ||
      (r.customsDeclarationNo || "").toLowerCase().includes(q) ||
      (r.transferNo || "").toLowerCase().includes(q);
  });

  if (filtered.length === 0) {
    dom.receiptsTbody.innerHTML = '<tr><td colspan="13" style="text-align:center; padding:16px;">Kirim qaydlari topilmadi</td></tr>';
    return;
  }

  dom.receiptsTbody.innerHTML = filtered.map((r, idx) => `
    <tr>
      <td style="text-align:center;">${idx + 1}</td>
      <td><code>${r.contractNo}</code></td>
      <td><strong>${r.invoiceNo}</strong></td>
      <td>${r.invoiceDate}</td>
      <td>${r.receiptDate}</td>
      <td><code>${r.customsDeclarationNo}</code></td>
      <td>${r.customsClearanceDate}</td>
      <td><span class="chip chip-success">${r.uzasboRegistrationDate}</span></td>
      <td><strong style="color:var(--blue-secondary);">${r.transferNo}</strong></td>
      <td>${r.transferDate}</td>
      <td style="text-align:right; font-weight:bold;">${r.deliveredQuantity}</td>
      <td>${r.uzasboAccountCode}<br/><small style="font-weight:600;">${formatNumber(r.uzasboBookValueUzs)} UZS</small></td>
      <td><span class="chip chip-info">${r.status}</span></td>
    </tr>
  `).join("");
}

// ====================================================================
// 4. HANDOVERS & TRANSFERS (ТЗ Section 6)
// ====================================================================
function renderHandovers() {
  if (!dom.handoversTbody) return;
  const q = (dom.searchHandovers ? dom.searchHandovers.value : "").trim().toLowerCase();
  const regFilter = dom.filterHandoverRegion ? dom.filterHandoverRegion.value : "";

  const filtered = appState.handovers.filter(h => {
    const matchQuery = !q || 
      (h.docNo || "").toLowerCase().includes(q) ||
      (h.institution || "").toLowerCase().includes(q) ||
      (h.itemDescription || "").toLowerCase().includes(q) ||
      (h.tin || "").toLowerCase().includes(q) ||
      (h.responsiblePerson || "").toLowerCase().includes(q);
    const matchReg = !regFilter || (h.region || "").includes(regFilter);
    return matchQuery && matchReg;
  });

  if (filtered.length === 0) {
    dom.handoversTbody.innerHTML = '<tr><td colspan="15" style="text-align:center; padding:16px;">Topshirish dalolatnomalari topilmadi</td></tr>';
    return;
  }

  dom.handoversTbody.innerHTML = filtered.map((h, idx) => `
    <tr>
      <td style="text-align:center;">${idx + 1}</td>
      <td><code>${h.transferNo}</code></td>
      <td><strong>${h.docNo}</strong></td>
      <td>${h.date}</td>
      <td>${h.region}</td>
      <td>${h.district || "-"}</td>
      <td><strong>${h.institution}</strong></td>
      <td><code>${h.tin}</code></td>
      <td>${h.department}</td>
      <td class="code-cell" onclick="window.viewPassportBySerial('${h.serialNumber}')">${h.itemDescription}</td>
      <td style="text-align:right; font-weight:bold;">${h.quantity} ${h.unit}</td>
      <td style="text-align:right;">$${formatNumber(h.totalValueUsd)}</td>
      <td style="text-align:right; font-weight:600;">${formatNumber(h.totalValueUzs)}</td>
      <td>${h.responsiblePerson}</td>
      <td>
        <button class="btn btn-secondary" style="padding:3px 8px; font-size:0.72rem;" onclick="window.printSingleHandover('${h.id}')">
          Печать Акта
        </button>
      </td>
    </tr>
  `).join("");
}

// ====================================================================
// 5. INSTALLATIONS & COMMISSIONING (ТЗ Section 7)
// ====================================================================
function renderInstallations() {
  if (!dom.installationsTbody) return;
  const q = (dom.searchInstallations ? dom.searchInstallations.value : "").trim().toLowerCase();
  const stFilter = dom.filterInstallStatus ? dom.filterInstallStatus.value : "";

  const filtered = appState.installations.filter(i => {
    const matchQuery = !q || 
      (i.equipmentName || "").toLowerCase().includes(q) ||
      (i.serialNumber || "").toLowerCase().includes(q) ||
      (i.organization || "").toLowerCase().includes(q) ||
      (i.supplierSpecialist || "").toLowerCase().includes(q);
    const matchSt = !stFilter || (i.status || "").includes(stFilter);
    return matchQuery && matchSt;
  });

  if (filtered.length === 0) {
    dom.installationsTbody.innerHTML = '<tr><td colspan="14" style="text-align:center; padding:16px;">Mavjud emas</td></tr>';
    return;
  }

  dom.installationsTbody.innerHTML = filtered.map((i, idx) => {
    const isProblem = (i.status || "").includes("Muammo");
    const isComm = (i.status || "").includes("Foydalanishga");
    return `
      <tr>
        <td style="text-align:center;">${idx + 1}</td>
        <td><code>${i.handoverDocNo || "-"}</code></td>
        <td><strong>${i.equipmentName}</strong></td>
        <td class="code-cell" onclick="window.viewPassportBySerial('${i.serialNumber}')">${i.serialNumber}</td>
        <td>${i.organization}</td>
        <td>${i.region}</td>
        <td><span class="chip ${isProblem ? 'chip-danger' : isComm ? 'chip-success' : 'chip-info'}">${i.status}</span></td>
        <td>${i.installationDate}</td>
        <td><small>${i.testingResults || "-"}</small></td>
        <td>${i.commissioningDate || "-"}</td>
        <td><code>${i.commissioningCertNo || "-"}</code></td>
        <td>${i.supplierSpecialist || "-"}</td>
        <td><span class="chip ${i.trainingCompleted && i.trainingCompleted.includes('Ha') ? 'chip-success' : 'chip-warning'}">${i.trainingCompleted}</span></td>
        <td>
          <button class="btn-passport" onclick="window.viewPassportBySerial('${i.serialNumber}')">Паспорт</button>
        </td>
      </tr>
    `;
  }).join("");
}

// ====================================================================
// 6. PIU MONITORING (ТЗ Section 8)
// ====================================================================
function renderMonitoring() {
  if (!dom.monitoringTbody) return;
  const q = (dom.searchMonitoring ? dom.searchMonitoring.value : "").trim().toLowerCase();
  const condFilter = dom.filterMonCondition ? dom.filterMonCondition.value : "";

  const filtered = appState.monitoringRecords.filter(m => {
    const matchQuery = !q || 
      (m.equipmentName || "").toLowerCase().includes(q) ||
      (m.serialNumber || "").toLowerCase().includes(q) ||
      (m.monitoredOrg || "").toLowerCase().includes(q) ||
      (m.identifiedDefects || "").toLowerCase().includes(q);
    const matchCond = !condFilter || (m.technicalCondition || "").includes(condFilter);
    return matchQuery && matchCond;
  });

  if (filtered.length === 0) {
    dom.monitoringTbody.innerHTML = '<tr><td colspan="14" style="text-align:center; padding:16px;">Monitoring tekshiruvlari topilmadi</td></tr>';
    return;
  }

  dom.monitoringTbody.innerHTML = filtered.map((m, idx) => {
    const isDefective = (m.technicalCondition || "").includes("Nuqsonli") || (m.technicalCondition || "").includes("Ta'mir");
    return `
      <tr>
        <td style="text-align:center;">${idx + 1}</td>
        <td><strong>${m.monitoringDate}</strong></td>
        <td>${m.monitoredOrg} <br/><small style="color:#64748b;">${m.region}</small></td>
        <td><strong>${m.equipmentName}</strong></td>
        <td class="code-cell" onclick="window.viewPassportBySerial('${m.serialNumber}')">${m.serialNumber}</td>
        <td><code>${m.inventoryNumber || "-"}</code></td>
        <td>${m.actualLocation || "-"}</td>
        <td><span class="chip chip-success">${m.physicalAvailability}</span></td>
        <td><span class="chip ${isDefective ? 'chip-danger' : 'chip-success'}">${m.technicalCondition}</span></td>
        <td style="color:${isDefective ? '#b91c1c' : '#334155'}; font-weight:${isDefective ? '600' : 'normal'};">${m.identifiedDefects || "-"}</td>
        <td><span class="chip chip-neutral">${m.photoEvidence || "Mavjud emas"}</span></td>
        <td>${m.recommendations || "-"}</td>
        <td><span class="chip chip-info">${m.monitoringResult}</span></td>
        <td>${m.monitoringOfficer}</td>
      </tr>
    `;
  }).join("");
}

// ====================================================================
// 7. INVENTORY AUDIT (ТЗ Section 9)
// ====================================================================
function renderInventory() {
  if (!dom.inventoryTbody) return;
  const q = (dom.searchInventory ? dom.searchInventory.value : "").trim().toLowerCase();
  const regFilter = dom.filterInvRegion ? dom.filterInvRegion.value : "";
  const discFilter = dom.filterInvDiscrepancy ? dom.filterInvDiscrepancy.value : "";

  const items = appState.inventoryItems || [];
  const filtered = items.filter(i => {
    const surplus = Math.max((Number(i.actualQty) || 0) - (Number(i.bookQty) || 0), 0);
    const shortage = Math.max((Number(i.bookQty) || 0) - (Number(i.actualQty) || 0), 0);

    const matchQuery = !q || 
      (i.code || "").toLowerCase().includes(q) ||
      (i.description || "").toLowerCase().includes(q) ||
      (i.serialNumber || "").toLowerCase().includes(q) ||
      (i.inventoryNumber || "").toLowerCase().includes(q) ||
      (i.institution || "").toLowerCase().includes(q) ||
      (i.responsiblePerson || "").toLowerCase().includes(q);
    const matchReg = !regFilter || (i.region || "").includes(regFilter);
    const matchDisc = !discFilter || (discFilter === "surplus" && surplus > 0) || (discFilter === "shortage" && shortage > 0);

    return matchQuery && matchReg && matchDisc;
  });

  if (filtered.length === 0) {
    dom.inventoryTbody.innerHTML = '<tr><td colspan="19" style="text-align:center; padding:16px;">Inventarizatsiya tovarlari topilmadi</td></tr>';
    return;
  }

  dom.inventoryTbody.innerHTML = filtered.map((i, idx) => {
    const bookQty = Number(i.bookQty) || 0;
    const actualQty = Number(i.actualQty) || 0;
    const surplus = Math.max(actualQty - bookQty, 0);
    const shortage = Math.max(bookQty - actualQty, 0);
    const unitPrice = Number(i.unitPrice) || 0;
    const bookVal = bookQty * unitPrice;
    const actualVal = actualQty * unitPrice;

    return `
      <tr>
        <td style="text-align:center;">${idx + 1}</td>
        <td class="code-cell" onclick="window.viewPassportBySerial('${i.serialNumber || i.code}')">${i.code}</td>
        <td><strong>${i.description}</strong></td>
        <td><code>${i.serialNumber || "-"}</code></td>
        <td><code>${i.inventoryNumber || "-"}</code></td>
        <td style="text-align:center;">${i.unit}</td>
        <td>${i.region}</td>
        <td>${i.institution} <br/><small style="color:#64748b;">${i.location || ""}</small></td>
        <td>${i.responsiblePerson}</td>
        <td style="text-align:right;">${bookQty}</td>
        <td style="text-align:right; font-weight:bold;">${actualQty}</td>
        <td style="text-align:right; color:#10b981; font-weight:bold;">${surplus > 0 ? surplus : "-"}</td>
        <td style="text-align:right; color:#ef4444; font-weight:bold;">${shortage > 0 ? shortage : "-"}</td>
        <td style="text-align:right;">$${formatNumber(unitPrice)}</td>
        <td style="text-align:right;">$${formatNumber(bookVal)}</td>
        <td style="text-align:right; font-weight:bold;">$${formatNumber(actualVal)}</td>
        <td><span class="chip ${i.condition && (i.condition.includes('Nuqson') || i.condition.includes('Ta\'mir')) ? 'chip-danger' : 'chip-success'}">${i.condition}</span></td>
        <td><code>${i.handoverDoc || "-"}</code></td>
        <td>
          <button class="btn-passport" onclick="window.viewPassportBySerial('${i.serialNumber || i.code}')">Паспорт</button>
        </td>
      </tr>
    `;
  }).join("");
}

// ====================================================================
// 8. REPORTS & ANALYTICS (13 REPORTS FROM ТЗ Section 10)
// ====================================================================
function renderActiveReport() {
  if (!dom.reportViewContainer || !dom.selectReportType) return;
  const reportId = dom.selectReportType.value;
  let html = "";

  switch (reportId) {
    case "1": // ТЭО
      html = generateTeoReportHtml();
      break;
    case "2": // Регионы
      html = generateRegionalReportHtml();
      break;
    case "3": // Контракты
      html = generateContractsReportHtml();
      break;
    case "4": // Поступление
      html = generateReceiptsReportHtml();
      break;
    case "5": // Transfer
      html = generateTransferReportHtml();
      break;
    case "6": // Передача организациям
      html = generateHandoverReportHtml();
      break;
    case "7": // Установка и ввод
      html = generateInstallationsReportHtml();
      break;
    case "8": // Гарантия
      html = generateWarrantyReportHtml();
      break;
    case "9": // Техсостояние
      html = generateConditionsReportHtml();
      break;
    case "10": // Мониторинг
      html = generateMonitoringReportHtml();
      break;
    case "11": // Инвентаризация
      html = generateInventoryReportHtml();
      break;
    case "12": // Расхождения и дефекты
      html = generateDiscrepanciesReportHtml();
      break;
    case "13": // Итоговый проектный отчёт
      html = generateFinalProjectClosureReportHtml();
      break;
    default:
      html = "<p>Выберите отчёт из списка.</p>";
  }

  dom.reportViewContainer.innerHTML = html;
}

// Report Generators
function generateTeoReportHtml() {
  const items = appState.teoItems;
  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>ОТЧЁТ № 1. ПЕРЕЧЕНЬ ОБОРУДОВАНИЯ СОГЛАСНО УТВЕРЖДЁННОМУ ТЭО</h2>
        <h3>FEASIBILITY STUDY ALLOCATION REPORT (ADB / AIIB)</h3>
        <p style="font-size:9pt; margin-top:4px;">Постановление: ${appState.projectMetadata.presidentialResolution} | ТЭО: ${appState.projectMetadata.teoNumber}</p>
      </div>
      <table class="doc-table">
        <thead>
          <tr>
            <th>№</th><th>Код</th><th>Наименование оборудования</th><th>Утверждено по ТЭО</th><th>Регионы распределения</th><th>Поставлено</th><th>Передано</th><th>Остаток</th><th>Цена USD</th>
          </tr>
        </thead>
        <tbody>
          ${items.map((i, idx) => `
            <tr>
              <td style="text-align:center;">${idx + 1}</td>
              <td><code>${i.itemCode}</code></td>
              <td><strong>${i.itemName}</strong></td>
              <td style="text-align:center; font-weight:bold;">${i.approvedQty}</td>
              <td>${i.region}</td>
              <td style="text-align:center;">${i.actualDeliveredQty}</td>
              <td style="text-align:center; font-weight:bold; color:#059669;">${i.actualTransferredQty}</td>
              <td style="text-align:center; font-weight:bold; color:${i.remainingQty > 0 ? '#d97706' : '#333'};">${i.remainingQty}</td>
              <td style="text-align:right;">$${formatNumber(i.unitPriceUsd)}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function generateRegionalReportHtml() {
  const regMap = {};
  UZBEKISTAN_REGIONS.forEach(r => {
    regMap[r] = { count: 0, qty: 0, valUsd: 0 };
  });

  appState.handovers.forEach(h => {
    const reg = h.region;
    if (regMap[reg]) {
      regMap[reg].count += 1;
      regMap[reg].qty += (Number(h.quantity) || 0);
      regMap[reg].valUsd += (Number(h.totalValueUsd) || 0);
    }
  });

  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>ОТЧЁТ № 2. РАСПРЕДЕЛЕНИЕ ОБОРУДОВАНИЯ ПО РЕГИОНАМ РЕСПУБЛИКИ УЗБЕКИСТАН</h2>
        <h3>REGIONAL EQUIPMENT DISTRIBUTION SUMMARY (14 REGIONS)</h3>
      </div>
      <table class="doc-table">
        <thead>
          <tr><th>№</th><th>Регион Республики Узбекистан</th><th>Кол-во актов передачи</th><th>Передано единиц оборудования</th><th>Общая стоимость USD</th></tr>
        </thead>
        <tbody>
          ${UZBEKISTAN_REGIONS.map((reg, idx) => `
            <tr>
              <td style="text-align:center;">${idx + 1}</td>
              <td><strong>${reg}</strong></td>
              <td style="text-align:center;">${regMap[reg].count}</td>
              <td style="text-align:center; font-weight:bold;">${regMap[reg].qty}</td>
              <td style="text-align:right; font-weight:bold;">$${formatNumber(regMap[reg].valUsd)}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function generateContractsReportHtml() {
  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>ОТЧЁТ № 3. ЗАКУПЛЕННОЕ ОБОРУДОВАНИЕ ПО КОНТРАКТАМ</h2>
        <h3>PROCUREMENT CONTRACTS & TECHNICAL SPECIFICATIONS REPORT</h3>
      </div>
      <table class="doc-table">
        <thead>
          <tr><th>№</th><th>Номер контракта</th><th>Дата</th><th>Поставщик</th><th>Оборудование</th><th>Кол-во</th><th>Товарная сумма USD</th><th>Итого по контракту USD</th><th>Гарантия</th></tr>
        </thead>
        <tbody>
          ${appState.contracts.map((c, idx) => `
            <tr>
              <td style="text-align:center;">${idx + 1}</td>
              <td><strong>${c.contractNo}</strong></td>
              <td>${c.contractDate}</td>
              <td>${c.supplier}</td>
              <td>${c.goodsDescription}</td>
              <td style="text-align:center; font-weight:bold;">${c.quantity}</td>
              <td style="text-align:right;">$${formatNumber(c.totalGoodsValueUsd)}</td>
              <td style="text-align:right; font-weight:bold;">$${formatNumber(c.totalContractAmountUsd)}</td>
              <td>${c.warrantyPeriod}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function generateReceiptsReportHtml() {
  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>ОТЧЁТ № 4. ФАКТИЧЕСКИ ПОСТУПИВШЕЕ ОБОРУДОВАНИЕ И УЧЁТ В УЗАСБО</h2>
        <h3>EQUIPMENT RECEIPTS, CUSTOMS CLEARANCE & UZASBO REGISTRATION</h3>
      </div>
      <table class="doc-table">
        <thead>
          <tr><th>№</th><th>Счёт-фактура</th><th>Дата поступления</th><th>ГТД / Таможня</th><th>Дата в УзАСБО</th><th>Номер Transfer</th><th>Поставлено</th><th>Стоимость UZS</th><th>Статус</th></tr>
        </thead>
        <tbody>
          ${appState.receipts.map((r, idx) => `
            <tr>
              <td style="text-align:center;">${idx + 1}</td>
              <td><strong>${r.invoiceNo}</strong></td>
              <td>${r.receiptDate}</td>
              <td><code>${r.customsDeclarationNo}</code></td>
              <td>${r.uzasboRegistrationDate}</td>
              <td><strong>${r.transferNo}</strong></td>
              <td style="text-align:center; font-weight:bold;">${r.deliveredQuantity}</td>
              <td style="text-align:right;">${formatNumber(r.uzasboBookValueUzs)} UZS</td>
              <td>${r.status}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function generateTransferReportHtml() {
  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>ОТЧЁТ № 5. РЕЕСТР ВНУТРЕННИХ ДОКУМЕНТОВ TRANSFER</h2>
        <h3>INTERNAL TRANSFER MOVEMENT DOCUMENTS REGISTRY</h3>
      </div>
      <table class="doc-table">
        <thead>
          <tr><th>№</th><th>Номер Transfer</th><th>Дата</th><th>Связанный контракт</th><th>Оборудование</th><th>Количество</th><th>Статус</th></tr>
        </thead>
        <tbody>
          ${appState.receipts.map((r, idx) => `
            <tr>
              <td style="text-align:center;">${idx + 1}</td>
              <td><strong>${r.transferNo}</strong></td>
              <td>${r.transferDate}</td>
              <td>${r.contractNo}</td>
              <td>${r.goodsDescription}</td>
              <td style="text-align:center; font-weight:bold;">${r.deliveredQuantity}</td>
              <td>УзАСБОда тасдиқланган</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function generateHandoverReportHtml() {
  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>ОТЧЁТ № 6. ПЕРЕДАЧА ОБОРУДОВАНИЯ КОНЕЧНЫМ ОРГАНИЗАЦИЯМ</h2>
        <h3>EQUIPMENT HANDOVER & ACCEPTANCE TO BENEFICIARY LABORATORIES</h3>
      </div>
      <table class="doc-table">
        <thead>
          <tr><th>№</th><th>Акт №</th><th>Дата</th><th>Организация-получатель</th><th>ИНН</th><th>Оборудование</th><th>Кол-во</th><th>Сумма USD</th><th>МОЛ</th></tr>
        </thead>
        <tbody>
          ${appState.handovers.map((h, idx) => `
            <tr>
              <td style="text-align:center;">${idx + 1}</td>
              <td><strong>${h.docNo}</strong></td>
              <td>${h.date}</td>
              <td><strong>${h.institution}</strong> (${h.region})</td>
              <td><code>${h.tin}</code></td>
              <td>${h.itemDescription}</td>
              <td style="text-align:center; font-weight:bold;">${h.quantity}</td>
              <td style="text-align:right;">$${formatNumber(h.totalValueUsd)}</td>
              <td>${h.responsiblePerson}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function generateInstallationsReportHtml() {
  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>ОТЧЁТ № 7. УСТАНОВКА, МОНТАЖ И ВВОД В ЭКСПЛУАТАЦИЮ</h2>
        <h3>INSTALLATION, TESTING, TRAINING & COMMISSIONING STATUS</h3>
      </div>
      <table class="doc-table">
        <thead>
          <tr><th>№</th><th>Оборудование</th><th>Серийный №</th><th>Организация</th><th>Статус установки</th><th>Дата ввода</th><th>Акт ввода №</th><th>Обучение</th></tr>
        </thead>
        <tbody>
          ${appState.installations.map((i, idx) => `
            <tr>
              <td style="text-align:center;">${idx + 1}</td>
              <td><strong>${i.equipmentName}</strong></td>
              <td><code>${i.serialNumber}</code></td>
              <td>${i.organization}</td>
              <td><strong>${i.status}</strong></td>
              <td>${i.commissioningDate || "-"}</td>
              <td>${i.commissioningCertNo || "-"}</td>
              <td>${i.trainingCompleted}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function generateWarrantyReportHtml() {
  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>ОТЧЁТ № 8. МОНИТОРИНГ ГАРАНТИЙНЫХ ОБЯЗАТЕЛЬСТВ</h2>
        <h3>SUPPLIER WARRANTY COMMITMENTS & MAINTENANCE TRACKING</h3>
      </div>
      <table class="doc-table">
        <thead>
          <tr><th>№</th><th>Контракт</th><th>Поставщик</th><th>Оборудование</th><th>Срок гарантии</th><th>Действует до</th><th>Статус гарантии</th></tr>
        </thead>
        <tbody>
          ${appState.contracts.map((c, idx) => `
            <tr>
              <td style="text-align:center;">${idx + 1}</td>
              <td><strong>${c.contractNo}</strong></td>
              <td>${c.supplier}</td>
              <td>${c.goodsDescription}</td>
              <td><strong>${c.warrantyPeriod}</strong></td>
              <td>2026/2027 г.</td>
              <td><span style="color:#059669; font-weight:bold;">Активна (Действует)</span></td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function generateConditionsReportHtml() {
  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>ОТЧЁТ № 9. ТЕХНИЧЕСКОЕ СОСТОЯНИЕ ОБОРУДОВАНИЯ (8 КАТЕГОРИЙ ТЗ)</h2>
        <h3>EQUIPMENT TECHNICAL CONDITION AUDIT SUMMARY</h3>
      </div>
      <table class="doc-table">
        <thead>
          <tr><th>№</th><th>Оборудование</th><th>Серийный №</th><th>Организация / Регион</th><th>Категория техсостояния</th><th>Конкретные дефекты / Описание</th></tr>
        </thead>
        <tbody>
          ${appState.monitoringRecords.map((m, idx) => `
            <tr>
              <td style="text-align:center;">${idx + 1}</td>
              <td><strong>${m.equipmentName}</strong></td>
              <td><code>${m.serialNumber}</code></td>
              <td>${m.monitoredOrg} (${m.region})</td>
              <td><strong>${m.technicalCondition}</strong></td>
              <td style="color:${m.technicalCondition.includes('Nuqson') || m.technicalCondition.includes('Ta\'mir') ? '#b91c1c' : '#333'};">${m.identifiedDefects || "Nuqsonlar yo'q"}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function generateMonitoringReportHtml() {
  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>ОТЧЁТ № 10. РЕЗУЛЬТАТЫ МОНИТОРИНГА ГРП</h2>
        <h3>PIU FIELD MONITORING VISITS & CORRECTIVE MEASURES REPORT</h3>
      </div>
      <table class="doc-table">
        <thead>
          <tr><th>№</th><th>Дата</th><th>Организация</th><th>Оборудование</th><th>Техсостояние</th><th>Рекомендации / Сроки</th><th>Результат</th><th>Инспектор ГРП</th></tr>
        </thead>
        <tbody>
          ${appState.monitoringRecords.map((m, idx) => `
            <tr>
              <td style="text-align:center;">${idx + 1}</td>
              <td>${m.monitoringDate}</td>
              <td>${m.monitoredOrg}</td>
              <td><strong>${m.equipmentName}</strong></td>
              <td>${m.technicalCondition}</td>
              <td>${m.recommendations}</td>
              <td><strong>${m.monitoringResult}</strong></td>
              <td>${m.monitoringOfficer}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function generateInventoryReportHtml() {
  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>ОТЧЁТ № 11. ИНВЕНТАРИЗАЦИОННЫЙ ОТЧЁТ И ВЫВЕРКА</h2>
        <h3>FULL INVENTORY AUDIT & PHYSICAL VERIFICATION REPORT</h3>
      </div>
      <table class="doc-table">
        <thead>
          <tr><th>№</th><th>Код</th><th>Оборудование</th><th>Серийный №</th><th>Организация</th><th>Учёт</th><th>Факт</th><th>Разница</th><th>Стоимость USD</th><th>Состояние</th></tr>
        </thead>
        <tbody>
          ${appState.inventoryItems.map((i, idx) => {
            const surplus = Math.max((Number(i.actualQty) || 0) - (Number(i.bookQty) || 0), 0);
            const shortage = Math.max((Number(i.bookQty) || 0) - (Number(i.actualQty) || 0), 0);
            return `
              <tr>
                <td style="text-align:center;">${idx + 1}</td>
                <td><code>${i.code}</code></td>
                <td><strong>${i.description}</strong></td>
                <td><code>${i.serialNumber || "-"}</code></td>
                <td>${i.institution}</td>
                <td style="text-align:center;">${i.bookQty}</td>
                <td style="text-align:center; font-weight:bold;">${i.actualQty}</td>
                <td style="text-align:center; font-weight:bold; color:${surplus > 0 ? '#10b981' : shortage > 0 ? '#ef4444' : '#333'};">
                  ${surplus > 0 ? `+${surplus} (Surplus)` : shortage > 0 ? `-${shortage} (Shortage)` : '0'}
                </td>
                <td style="text-align:right;">$${formatNumber((Number(i.actualQty) || 0) * (Number(i.unitPrice) || 0))}</td>
                <td>${i.condition}</td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function generateDiscrepanciesReportHtml() {
  const discItems = appState.inventoryItems.filter(i => {
    const s = Math.max((Number(i.actualQty) || 0) - (Number(i.bookQty) || 0), 0);
    const sh = Math.max((Number(i.bookQty) || 0) - (Number(i.actualQty) || 0), 0);
    const isDef = (i.condition || "").includes("Ta'mir") || (i.condition || "").includes("Nuqson");
    return s > 0 || sh > 0 || isDef;
  });

  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>ОТЧЁТ № 12. ВЫЯВЛЕННЫЕ РАСХОЖДЕНИЯ, ИЗЛИШКИ, ДЕФИЦИТ И ДЕФЕКТЫ</h2>
        <h3>DISCREPANCIES, SURPLUS, SHORTAGE & DEFECTIVE EQUIPMENT AUDIT</h3>
      </div>
      <table class="doc-table">
        <thead>
          <tr><th>№</th><th>Код</th><th>Оборудование</th><th>Организация</th><th>Учёт</th><th>Факт</th><th>Тип расхождения</th><th>Техсостояние</th><th>Пояснения и принятые меры</th></tr>
        </thead>
        <tbody>
          ${discItems.map((i, idx) => {
            const surplus = Math.max((Number(i.actualQty) || 0) - (Number(i.bookQty) || 0), 0);
            const shortage = Math.max((Number(i.bookQty) || 0) - (Number(i.actualQty) || 0), 0);
            return `
              <tr>
                <td style="text-align:center;">${idx + 1}</td>
                <td><code>${i.code}</code></td>
                <td><strong>${i.description}</strong></td>
                <td>${i.institution}</td>
                <td style="text-align:center;">${i.bookQty}</td>
                <td style="text-align:center; font-weight:bold;">${i.actualQty}</td>
                <td style="font-weight:bold; color:${surplus > 0 ? '#10b981' : '#ef4444'};">
                  ${surplus > 0 ? `Ortiqcha (+${surplus})` : shortage > 0 ? `Kamomad (-${shortage})` : 'Nosozlik'}
                </td>
                <td>${i.condition}</td>
                <td>${i.remarks || "Buxgalteriyaga kirim yoki servisga yuborish choralari ko'rilmoqda"}</td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function generateFinalProjectClosureReportHtml() {
  const meta = appState.projectMetadata;
  const c = appState.committee;
  const teoApproved = appState.teoItems.reduce((acc, i) => acc + (Number(i.approvedQty) || 0), 0);
  const transferred = appState.handovers.reduce((acc, h) => acc + (Number(h.quantity) || 0), 0);

  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>ОТЧЁТ № 13. ИТОГОВЫЙ ОТЧЁТ ПО ОБОРУДОВАНИЮ В РАМКАХ ПРОЕКТА АБР / АБИИ</h2>
        <h3>FINAL PROJECT CLOSURE INVENTORY & ASSET REPORT (ADB & AIIB)</h3>
        <p style="font-size:9.5pt; margin-top:4px;">PEMS-UZ &bull; Project Equipment Management System &bull; Uzbekistan</p>
      </div>

      <table class="doc-meta-table">
        <tr>
          <td><strong>Loyiha:</strong> ${meta.projectName}</td>
          <td><strong>Buyruq:</strong> ${c.orderNo}</td>
        </tr>
        <tr>
          <td><strong>Prezident qarori:</strong> ${meta.presidentialResolution}</td>
          <td><strong>Davr:</strong> ${c.startDate} dan ${c.endDate} gacha</td>
        </tr>
        <tr>
          <td><strong>Tasdiqlangan TEO:</strong> ${meta.teoNumber}</td>
          <td><strong>Ijrochi:</strong> ${c.organization}</td>
        </tr>
      </table>

      <h4 style="margin:14px 0 6px 0; font-size:11pt; text-transform:uppercase;">1. Loyiha ko'rsatkichlari:</h4>
      <p style="font-size:10pt; line-height:1.5;">
        TEO bo'yicha rejalashtirilgan <strong>${teoApproved} komplekt</strong> laboratoriya va diagnostika uskunalaridan 
        <strong>${transferred} komplekt</strong> uskunalar 14 ta hududdagi laboratoriyalarga to'liq yetkazib berildi va foydalanishga topshirildi.
        O'rnatish va montaj ishlari 96% ijobiy ko'rsatkich bilan yakunlandi.
      </p>

      <h4 style="margin:16px 0 6px 0; font-size:11pt; text-transform:uppercase;">2. Komissiya xulosasi:</h4>
      <p style="font-size:10pt; line-height:1.5;">${c.conclusion}</p>

      <h4 style="margin:16px 0 6px 0; font-size:11pt; text-transform:uppercase;">3. Komissiya tarkibi va imzolar:</h4>
      <table class="doc-table">
        <thead>
          <tr><th>Lavozim</th><th>F.I.Sh.</th><th>Imzo</th><th>Sana</th></tr>
        </thead>
        <tbody>
          ${c.members.map(m => `
            <tr>
              <td><strong>${m.role}</strong></td>
              <td>${m.name}</td>
              <td style="text-align:center;">${m.signature || "______________"}</td>
              <td style="text-align:center;">${m.date || meta.inventoryDate}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

// ====================================================================
// 9. EQUIPMENT PASSPORT MODAL (21 POINTS FROM ТЗ Section 13)
// ====================================================================
function findEquipmentByQuery(query) {
  const q = query.toLowerCase();
  // Check inventory
  let found = appState.inventoryItems.find(i => 
    (i.serialNumber || "").toLowerCase().includes(q) ||
    (i.code || "").toLowerCase().includes(q) ||
    (i.inventoryNumber || "").toLowerCase().includes(q)
  );
  if (found) return found;

  // Check handovers
  const h = appState.handovers.find(h => 
    (h.serialNumber || "").toLowerCase().includes(q) ||
    (h.itemCode || "").toLowerCase().includes(q) ||
    (h.inventoryNumber || "").toLowerCase().includes(q)
  );
  if (h) {
    return {
      code: h.itemCode,
      description: h.itemDescription,
      serialNumber: h.serialNumber,
      inventoryNumber: h.inventoryNumber,
      institution: h.institution,
      region: h.region,
      location: h.department,
      responsiblePerson: h.responsiblePerson,
      unitPrice: h.totalValueUsd,
      currency: "USD",
      handoverDoc: h.docNo,
      transferNo: h.transferNo,
      condition: h.condition || "Ishga yaroqli"
    };
  }

  // Fallback to first item
  return appState.inventoryItems[0];
}

window.viewPassportBySerial = function(serial) {
  const item = findEquipmentByQuery(serial);
  if (item) openEquipmentPassport(item);
};

window.viewPassportByCode = function(code) {
  const item = findEquipmentByQuery(code);
  if (item) openEquipmentPassport(item);
};

function openEquipmentPassport(item) {
  const meta = appState.projectMetadata;
  const pemsId = `PEMS-UZ-2024-${item.code || 'EQP'}-${item.serialNumber || '001'}`;

  // Find linked data
  const linkedContract = appState.contracts[0] || {};
  const linkedReceipt = appState.receipts[0] || {};
  const linkedHandover = appState.handovers.find(h => h.serialNumber === item.serialNumber) || appState.handovers[0] || {};
  const linkedInstall = appState.installations.find(i => i.serialNumber === item.serialNumber) || appState.installations[0] || {};
  const linkedMonitoring = appState.monitoringRecords.find(m => m.serialNumber === item.serialNumber) || appState.monitoringRecords[0] || {};

  const priceUsd = Number(item.unitPrice) || Number(linkedContract.unitPriceUsd) || 38500;
  const priceUzs = priceUsd * (Number(meta.usdRate) || 12850);

  const html = `
    <!-- Top Identity Bar -->
    <div style="background:var(--navy-soft); border:1px solid var(--blue-light); border-radius:var(--radius-md); padding:14px 18px; margin-bottom:16px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
      <div>
        <div style="font-size:0.7rem; font-weight:800; color:var(--blue-secondary); text-transform:uppercase; letter-spacing:0.8px;">
          PEMS-UZ UNIQUE ASSET IDENTIFIER
        </div>
        <div style="font-size:1.15rem; font-weight:800; color:var(--navy-primary); font-family:var(--font-mono);">
          ${pemsId}
        </div>
        <div style="font-size:0.82rem; font-weight:600; color:var(--slate-700); margin-top:2px;">
          ${item.description || linkedContract.goodsDescription}
        </div>
      </div>
      <div style="display:flex; gap:8px;">
        <span class="chip chip-success" style="font-size:0.75rem;">${item.condition || "Foydalanishda"}</span>
        <span class="chip chip-info" style="font-size:0.75rem;">АБР / АБИИ</span>
      </div>
    </div>

    <!-- 21 Points Structured Grid -->
    <div class="passport-grid">
      <!-- Section A: Идентификаторы и ТЭО -->
      <div class="passport-section-card">
        <h4 class="passport-section-title">1. Основание закупки и ТЭО</h4>
        <div class="passport-prop-row"><span class="passport-prop-label">1. Уникальный ID:</span><span class="passport-prop-val">${pemsId}</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">2. Постановление Президента:</span><span class="passport-prop-val">${meta.presidentialResolution}</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">3. Номер и дата ТЭО:</span><span class="passport-prop-val">${meta.teoNumber}</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">4. Утверждено по ТЭО:</span><span class="passport-prop-val">14 комплектов</span></div>
      </div>

      <!-- Section B: Контракт и поставка -->
      <div class="passport-section-card">
        <h4 class="passport-section-title">2. Контракт и характеристики</h4>
        <div class="passport-prop-row"><span class="passport-prop-label">5. Номер контракта:</span><span class="passport-prop-val">${linkedContract.contractNo || "ADB-MED-G01/2024"}</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">6. Поставщик:</span><span class="passport-prop-val">${linkedContract.supplier || "Bio-Rad / InterMed"}</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">7. Производитель / Страна:</span><span class="passport-prop-val">${linkedContract.manufacturer || "Bio-Rad"} (${linkedContract.countryOfOrigin || "USA"})</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">8. Модель / Год выпуска:</span><span class="passport-prop-val">${linkedContract.model || "CFX96 Touch"} (${linkedContract.yearOfManufacture || "2024"})</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">9. Серийный номер:</span><span class="passport-prop-val"><code>${item.serialNumber || "BR-PCR-98214-01"}</code></span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">10. Инвентарный номер:</span><span class="passport-prop-val"><code>${item.inventoryNumber || "INV-SAM-00421"}</code></span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">11. Гарантийный срок:</span><span class="passport-prop-val">${linkedContract.warrantyPeriod || "36 месяцев"}</span></div>
      </div>

      <!-- Section C: Стоимость и УзАСБО -->
      <div class="passport-section-card">
        <h4 class="passport-section-title">3. Бухгалтерский учёт и УзАСБО</h4>
        <div class="passport-prop-row"><span class="passport-prop-label">12. Стоимость USD:</span><span class="passport-prop-val" style="color:var(--navy-primary); font-weight:700;">$${formatNumber(priceUsd)}</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">13. Стоимость UZS:</span><span class="passport-prop-val">${formatNumber(priceUzs)} UZS</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">14. Счёт-фактура № и дата:</span><span class="passport-prop-val">${linkedReceipt.invoiceNo || "INV-2024-8841"} (${linkedReceipt.invoiceDate || "10.06.2024"})</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">15. ГТД / Таможня №:</span><span class="passport-prop-val">${linkedReceipt.customsDeclarationNo || "26001/004812"}</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">16. Дата в УзАСБО / Schyot:</span><span class="passport-prop-val">${linkedReceipt.uzasboRegistrationDate || "22.06.2024"} (013-schyot)</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">17. Номер и дата Transfer:</span><span class="passport-prop-val"><strong>${item.transferNo || linkedHandover.transferNo || "TRF-2024-001"}</strong></span></div>
      </div>

      <!-- Section D: Передача, монтаж и эксплуатация -->
      <div class="passport-section-card">
        <h4 class="passport-section-title">4. Местонахождение и эксплуатация</h4>
        <div class="passport-prop-row"><span class="passport-prop-label">18. Получатель:</span><span class="passport-prop-val"><strong>${item.institution || linkedHandover.institution}</strong></span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">19. Регион и адрес:</span><span class="passport-prop-val">${item.region || linkedHandover.region}</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">20. Подразделение / Помещение:</span><span class="passport-prop-val">${item.location || linkedHandover.department || "PZR xonasi № 204"}</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">21. МОЛ (Материально ответственный):</span><span class="passport-prop-val">${item.responsiblePerson || linkedHandover.responsiblePerson}</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">Акт приёма-передачи №:</span><span class="passport-prop-val">${item.handoverDoc || linkedHandover.docNo || "ACT-01/2024"}</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">Ввод в эксплуатацию / Акт:</span><span class="passport-prop-val">${linkedInstall.commissioningDate || "16.07.2024"} (${linkedInstall.commissioningCertNo || "COM-01"})</span></div>
        <div class="passport-prop-row"><span class="passport-prop-label">Текущее техсостояние (ТЗ):</span><span class="passport-prop-val"><span class="chip chip-success">${item.condition || linkedMonitoring.technicalCondition || "Ishga yaroqli"}</span></span></div>
      </div>
    </div>

    <!-- Lifecycle History Timeline (Full Chain of Custody) -->
    <div class="timeline-wrap">
      <h4 class="timeline-title">Полная хронология движения и проверок (Chain of Custody Timeline)</h4>
      <div class="timeline">
        <div class="timeline-item">
          <span class="timeline-time">16.10.2023:</span>
          <strong>Утверждение закупки:</strong> Включение в перечень постановления Президента РУз № PQ-3818 и ТЭО проекта.
        </div>
        <div class="timeline-item">
          <span class="timeline-time">15.03.2024:</span>
          <strong>Заключение контракта:</strong> Подписание контракта № ${linkedContract.contractNo || "ADB-MED-G01/2024"} с поставщиком Bio-Rad.
        </div>
        <div class="timeline-item">
          <span class="timeline-time">15.06.2024:</span>
          <strong>Таможенное оформление и поступление:</strong> Оформлена декларация ГТД № ${linkedReceipt.customsDeclarationNo || "26001/004812"}, приёмка на центральный склад ГРП.
        </div>
        <div class="timeline-item">
          <span class="timeline-time">22.06.2024:</span>
          <strong>Регистрация в УзАСБО:</strong> Централизованное оприходование на баланс (счёт 013), сформирован внутренний идентификатор <strong>${item.transferNo || "TRF-2024-001"}</strong>.
        </div>
        <div class="timeline-item">
          <span class="timeline-time">05.07.2024:</span>
          <strong>Передача получателю:</strong> Подписан официальный двуязычный Акт приёма-передачи № ${item.handoverDoc || "ACT-01/2024"} с передачей МОЛ ${item.responsiblePerson}.
        </div>
        <div class="timeline-item">
          <span class="timeline-time">16.07.2024:</span>
          <strong>Монтаж и ввод в эксплуатацию:</strong> Инженером ${linkedInstall.supplierSpecialist || "Bio-Rad"} проведён монтаж, калибровка и обучение персонала лаборатории. Подписан акт ввода № ${linkedInstall.commissioningCertNo || "COM-01/2024"}.
        </div>
        <div class="timeline-item">
          <span class="timeline-time">14.10.2024:</span>
          <strong>Мониторинговый выезд ГРП:</strong> Инспектор ${linkedMonitoring.monitoringOfficer || "Qodirov B."} подтвердил фактическое наличие, работоспособность и соответствие инвентарных номеров.
        </div>
        <div class="timeline-item">
          <span class="timeline-time">12.02.2026:</span>
          <strong>Итоговая инвентаризация:</strong> Инвентаризационной комиссией проверено фактическое наличие, расхождений не выявлено.
        </div>
      </div>
    </div>
  `;

  dom.passportModalBody.innerHTML = html;
  dom.passportModal.classList.remove("hidden");
}

function renderQuickPassportChips() {
  if (!dom.passportQuickChips) return;
  const items = appState.inventoryItems.slice(0, 6);
  dom.passportQuickChips.innerHTML = items.map(i => `
    <button class="btn btn-secondary" style="font-size:0.7rem; padding:3px 7px;" onclick="window.viewPassportBySerial('${i.serialNumber || i.code}')">
      ${i.code} (${i.serialNumber || 'SN'})
    </button>
  `).join("");
}

function renderRolesMatrix() {
  if (!dom.rolesMatrixInfo) return;
  const roleObj = USER_ROLES.find(r => r.id === currentRole) || USER_ROLES[0];
  dom.rolesMatrixInfo.innerHTML = `
    <div style="background:#fff; border:1px solid var(--slate-200); border-radius:var(--radius-sm); padding:10px 12px;">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
        <strong style="color:var(--navy-primary); font-size:0.82rem;">${roleObj.nameRu}</strong>
        <span class="user-badge" style="background:${roleObj.color};">${roleObj.badge}</span>
      </div>
      <p style="margin-bottom:8px; color:var(--slate-600);">${roleObj.descRu}</p>
      <div style="font-size:0.72rem; color:var(--slate-500);">
        <strong>Права доступа:</strong> ${roleObj.permissions.join(", ")}
      </div>
    </div>
  `;
}

function renderAuditLogs() {
  if (!dom.auditLogTbody) return;
  const logs = appState.auditLogs || [];
  dom.auditLogTbody.innerHTML = logs.slice(0, 50).map((l, idx) => `
    <tr>
      <td style="text-align:center;">${idx + 1}</td>
      <td><small><code>${l.timestamp}</code></small></td>
      <td><strong>${l.userName}</strong></td>
      <td><span class="chip chip-neutral">${l.module}</span></td>
      <td><span class="chip ${l.action === 'CREATE' ? 'chip-success' : l.action === 'UPDATE' ? 'chip-info' : 'chip-warning'}">${l.action}</span></td>
      <td><code>${l.entityId}</code></td>
      <td>${l.details}</td>
      <td><small style="color:#64748b;">${l.oldValue || "-"}</small></td>
      <td><small style="color:#059669; font-weight:600;">${l.newValue || "-"}</small></td>
    </tr>
  `).join("");
}

// ====================================================================
// MODAL: UNIVERSAL ADD / EDIT ITEM
// ====================================================================
function openItemModal(targetModule = "inventory") {
  dom.modalTargetModule.value = targetModule;
  buildDynamicForm(targetModule);
  dom.itemModal.classList.remove("hidden");
}

function buildDynamicForm(moduleType) {
  let fieldsHtml = "";
  switch (moduleType) {
    case "teo":
      fieldsHtml = `
        <div class="form-group span-2"><label>Наименование оборудования *</label><input type="text" id="f-teo-name" class="form-control" required placeholder="Real-Time PCR..." /></div>
        <div class="form-group"><label>Код оборудования *</label><input type="text" id="f-teo-code" class="form-control" required placeholder="EQP-PCR-01" /></div>
        <div class="form-group"><label>Утверждённое количество *</label><input type="number" id="f-teo-approved" class="form-control" min="1" value="1" required /></div>
        <div class="form-group span-2"><label>Регион распределения *</label><select id="f-teo-region" class="form-control">${UZBEKISTAN_REGIONS.map(r => `<option value="${r}">${r}</option>`).join("")}</select></div>
        <div class="form-group span-2"><label>Получающая организация *</label><input type="text" id="f-teo-org" class="form-control" required placeholder="Sanepidqo'm boshqarmasi..." /></div>
        <div class="form-group"><label>Планируемое количество *</label><input type="number" id="f-teo-planned" class="form-control" min="1" value="1" required /></div>
        <div class="form-group"><label>Цена USD *</label><input type="number" id="f-teo-price" class="form-control" min="1" value="10000" required /></div>
      `;
      break;
    case "contract":
      fieldsHtml = `
        <div class="form-group"><label>Номер контракта *</label><input type="text" id="f-cnt-no" class="form-control" required placeholder="ADB-MED-G01/2024" /></div>
        <div class="form-group"><label>Дата контракта *</label><input type="date" id="f-cnt-date" class="form-control" required /></div>
        <div class="form-group span-2"><label>Поставщик *</label><input type="text" id="f-cnt-supplier" class="form-control" required placeholder="Bio-Rad / Agilent..." /></div>
        <div class="form-group span-2"><label>Наименование товара *</label><input type="text" id="f-cnt-desc" class="form-control" required /></div>
        <div class="form-group"><label>Количество *</label><input type="number" id="f-cnt-qty" class="form-control" min="1" value="1" required /></div>
        <div class="form-group"><label>Цена за единицу (USD) *</label><input type="number" id="f-cnt-price" class="form-control" min="1" value="5000" required /></div>
        <div class="form-group"><label>Производитель / Страна</label><input type="text" id="f-cnt-manuf" class="form-control" placeholder="Bio-Rad (USA)" /></div>
        <div class="form-group"><label>Срок гарантии</label><input type="text" id="f-cnt-warranty" class="form-control" value="36 oy / 36 months" /></div>
      `;
      break;
    case "receipt":
      fieldsHtml = `
        <div class="form-group"><label>Номер счёта-фактуры *</label><input type="text" id="f-rec-inv" class="form-control" required placeholder="INV-2024-..." /></div>
        <div class="form-group"><label>Дата поступления *</label><input type="date" id="f-rec-date" class="form-control" required /></div>
        <div class="form-group span-2"><label>Номер ГТД (Таможенная декларация) *</label><input type="text" id="f-rec-ccd" class="form-control" required placeholder="26001/..." /></div>
        <div class="form-group"><label>Номер Transfer (Внутренний) *</label><input type="text" id="f-rec-trf" class="form-control" required value="TRF-${Date.now().toString().slice(-4)}" /></div>
        <div class="form-group"><label>Поставленное количество *</label><input type="number" id="f-rec-qty" class="form-control" min="1" value="1" required /></div>
        <div class="form-group span-2"><label>Дата регистрации в УзАСБО *</label><input type="date" id="f-rec-uzasbo" class="form-control" required /></div>
      `;
      break;
    case "handover":
      fieldsHtml = `
        <div class="form-group"><label>Номер акта передачи *</label><input type="text" id="f-ho-doc" class="form-control" required value="ACT-${Date.now().toString().slice(-4)}" /></div>
        <div class="form-group"><label>Дата акта *</label><input type="date" id="f-ho-date" class="form-control" required /></div>
        <div class="form-group span-2"><label>Регион *</label><select id="f-ho-reg" class="form-control">${UZBEKISTAN_REGIONS.map(r => `<option value="${r}">${r}</option>`).join("")}</select></div>
        <div class="form-group span-2"><label>Организация-получатель *</label><input type="text" id="f-ho-org" class="form-control" required placeholder="Sanepidqo'm boshqarmasi..." /></div>
        <div class="form-group"><label>ИНН (СТИР) *</label><input type="text" id="f-ho-tin" class="form-control" required placeholder="201482910" /></div>
        <div class="form-group"><label>МОЛ (F.I.Sh.) *</label><input type="text" id="f-ho-resp" class="form-control" required placeholder="F.I.Sh." /></div>
        <div class="form-group span-2"><label>Наименование оборудования *</label><input type="text" id="f-ho-desc" class="form-control" required /></div>
        <div class="form-group"><label>Серийный номер</label><input type="text" id="f-ho-serial" class="form-control" placeholder="SN-..." /></div>
        <div class="form-group"><label>Количество *</label><input type="number" id="f-ho-qty" class="form-control" min="1" value="1" required /></div>
      `;
      break;
    case "installation":
      fieldsHtml = `
        <div class="form-group span-2"><label>Оборудование *</label><input type="text" id="f-ins-name" class="form-control" required /></div>
        <div class="form-group"><label>Серийный номер *</label><input type="text" id="f-ins-serial" class="form-control" required /></div>
        <div class="form-group"><label>Статус установки *</label><select id="f-ins-status" class="form-control">${INSTALLATION_STATUSES.map(s => `<option value="${s}">${s}</option>`).join("")}</select></div>
        <div class="form-group span-2"><label>Организация *</label><input type="text" id="f-ins-org" class="form-control" required /></div>
        <div class="form-group"><label>Дата монтажа</label><input type="date" id="f-ins-date" class="form-control" /></div>
        <div class="form-group"><label>Дата ввода в эксплуатацию</label><input type="date" id="f-ins-comm" class="form-control" /></div>
        <div class="form-group span-2"><label>Специалист поставщика</label><input type="text" id="f-ins-spec" class="form-control" placeholder="Инженер..." /></div>
      `;
      break;
    case "monitoring":
      fieldsHtml = `
        <div class="form-group span-2"><label>Организация проверки *</label><input type="text" id="f-mon-org" class="form-control" required /></div>
        <div class="form-group"><label>Дата мониторинга *</label><input type="date" id="f-mon-date" class="form-control" required /></div>
        <div class="form-group"><label>Серийный номер *</label><input type="text" id="f-mon-serial" class="form-control" required /></div>
        <div class="form-group span-2"><label>Оборудование *</label><input type="text" id="f-mon-name" class="form-control" required /></div>
        <div class="form-group span-2"><label>Техническое состояние (ТЗ) *</label><select id="f-mon-cond" class="form-control">${ITEM_TECHNICAL_CONDITIONS.map(c => `<option value="${c}">${c}</option>`).join("")}</select></div>
        <div class="form-group span-2"><label>Описание конкретных дефектов</label><input type="text" id="f-mon-defects" class="form-control" placeholder="Экран singan, kabel shikastlangan..." /></div>
        <div class="form-group span-2"><label>Рекомендации ГРП и сроки</label><input type="text" id="f-mon-recom" class="form-control" placeholder="Yetkazib beruvchiga talabnoma yuborilsin..." /></div>
      `;
      break;
    case "inventory":
    default:
      fieldsHtml = `
        <div class="form-group span-2"><label>Наименование оборудования *</label><input type="text" id="f-inv-desc" class="form-control" required placeholder="PCR / Avtoklav..." /></div>
        <div class="form-group"><label>Код оборудования *</label><input type="text" id="f-inv-code" class="form-control" required placeholder="EQP-..." /></div>
        <div class="form-group"><label>Серийный номер *</label><input type="text" id="f-inv-serial" class="form-control" required placeholder="SN-..." /></div>
        <div class="form-group"><label>Инвентарный номер</label><input type="text" id="f-inv-num" class="form-control" placeholder="INV-..." /></div>
        <div class="form-group"><label>Единица измерения</label><select id="f-inv-unit" class="form-control">${UNITS.map(u => `<option value="${u}">${u}</option>`).join("")}</select></div>
        <div class="form-group span-2"><label>Регион *</label><select id="f-inv-reg" class="form-control">${UZBEKISTAN_REGIONS.map(r => `<option value="${r}">${r}</option>`).join("")}</select></div>
        <div class="form-group span-2"><label>Организация / Подразделение *</label><input type="text" id="f-inv-inst" class="form-control" required placeholder="Sanepidqo'm..." /></div>
        <div class="form-group"><label>МОЛ (F.I.Sh.) *</label><input type="text" id="f-inv-resp" class="form-control" required placeholder="F.I.Sh." /></div>
        <div class="form-group"><label>Учётное количество (книга) *</label><input type="number" id="f-inv-bqty" class="form-control" min="0" value="1" required /></div>
        <div class="form-group"><label>Фактическое количество *</label><input type="number" id="f-inv-aqty" class="form-control" min="0" value="1" required /></div>
        <div class="form-group"><label>Цена USD *</label><input type="number" id="f-inv-price" class="form-control" min="0" value="1000" required /></div>
        <div class="form-group span-2"><label>Техсостояние</label><select id="f-inv-cond" class="form-control">${ITEM_TECHNICAL_CONDITIONS.map(c => `<option value="${c}">${c}</option>`).join("")}</select></div>
      `;
      break;
  }
  dom.dynamicFormFields.innerHTML = fieldsHtml;
}

function handleItemFormSubmit(e) {
  e.preventDefault();
  const mod = dom.modalTargetModule.value;

  if (mod === "teo") {
    const newTeo = {
      id: "teo-" + Date.now(),
      presidentialDoc: appState.projectMetadata.presidentialResolution,
      teoDoc: appState.projectMetadata.teoNumber,
      itemCode: document.getElementById("f-teo-code").value,
      itemName: document.getElementById("f-teo-name").value,
      approvedQty: Number(document.getElementById("f-teo-approved").value) || 1,
      region: document.getElementById("f-teo-region").value,
      receivingOrg: document.getElementById("f-teo-org").value,
      plannedQty: Number(document.getElementById("f-teo-planned").value) || 1,
      actualDeliveredQty: 0,
      actualTransferredQty: 0,
      remainingQty: Number(document.getElementById("f-teo-approved").value) || 1,
      unitPriceUsd: Number(document.getElementById("f-teo-price").value) || 0,
      currency: "USD",
      status: "Yangi kiritilgan"
    };
    appState.teoItems.push(newTeo);
    logAudit("ТЭО", "CREATE", newTeo.itemCode, `Yangi TEO pozitsiyasi qo'shildi: ${newTeo.itemName}`);
  } else if (mod === "contract") {
    const newCnt = {
      id: "cnt-" + Date.now(),
      contractNo: document.getElementById("f-cnt-no").value,
      contractDate: document.getElementById("f-cnt-date").value,
      supplier: document.getElementById("f-cnt-supplier").value,
      goodsDescription: document.getElementById("f-cnt-desc").value,
      quantity: Number(document.getElementById("f-cnt-qty").value) || 1,
      unitPriceUsd: Number(document.getElementById("f-cnt-price").value) || 0,
      totalGoodsValueUsd: (Number(document.getElementById("f-cnt-qty").value) || 1) * (Number(document.getElementById("f-cnt-price").value) || 0),
      totalContractAmountUsd: (Number(document.getElementById("f-cnt-qty").value) || 1) * (Number(document.getElementById("f-cnt-price").value) || 0),
      manufacturer: document.getElementById("f-cnt-manuf").value || "-",
      countryOfOrigin: "-",
      warrantyPeriod: document.getElementById("f-cnt-warranty").value || "36 oy",
      status: "Faol / Active"
    };
    appState.contracts.push(newCnt);
    logAudit("Контракты", "CREATE", newCnt.contractNo, `Yangi shartnoma ro'yxatga olindi: ${newCnt.contractNo}`);
  } else if (mod === "receipt") {
    const newRec = {
      id: "rec-" + Date.now(),
      contractNo: "ADB-MED-G01/2024",
      invoiceNo: document.getElementById("f-rec-inv").value,
      invoiceDate: document.getElementById("f-rec-date").value,
      receiptDate: document.getElementById("f-rec-date").value,
      customsDeclarationNo: document.getElementById("f-rec-ccd").value,
      customsClearanceDate: document.getElementById("f-rec-date").value,
      uzasboRegistrationDate: document.getElementById("f-rec-uzasbo").value,
      transferNo: document.getElementById("f-rec-trf").value,
      transferDate: document.getElementById("f-rec-date").value,
      deliveredQuantity: Number(document.getElementById("f-rec-qty").value) || 1,
      goodsDescription: "Uskunalar to'plami",
      uzasboAccountCode: "013-schyot",
      uzasboBookValueUzs: 100000000,
      status: "УзАСБОда кирим қилинди"
    };
    appState.receipts.push(newRec);
    logAudit("Поступление", "CREATE", newRec.transferNo, `Kirim va Transfer ro'yxatdan o'tdi: ${newRec.transferNo}`);
  } else if (mod === "handover") {
    const newHo = {
      id: "ho-" + Date.now(),
      transferNo: "TRF-2024-001",
      transferDate: document.getElementById("f-ho-date").value,
      docNo: document.getElementById("f-ho-doc").value,
      date: document.getElementById("f-ho-date").value,
      region: document.getElementById("f-ho-reg").value,
      district: "-",
      institution: document.getElementById("f-ho-org").value,
      tin: document.getElementById("f-ho-tin").value,
      department: "Markaziy laboratoriya",
      address: "-",
      itemCode: "EQP-GEN-01",
      itemDescription: document.getElementById("f-ho-desc").value,
      serialNumber: document.getElementById("f-ho-serial").value || "SN-" + Date.now(),
      inventoryNumber: "INV-" + Date.now().toString().slice(-5),
      quantity: Number(document.getElementById("f-ho-qty").value) || 1,
      unit: "komplekt",
      totalValueUsd: 15000,
      totalValueUzs: 15000 * 12850,
      responsiblePerson: document.getElementById("f-ho-resp").value,
      transferor: "GRP uskunalar bo'limi",
      recipientName: document.getElementById("f-ho-resp").value,
      condition: "Yangi",
      status: "Imzolangan"
    };
    appState.handovers.push(newHo);
    logAudit("Передача", "CREATE", newHo.docNo, `Topshirish-qabul qilish akti tuzildi: ${newHo.docNo}`);
  } else if (mod === "installation") {
    const newIns = {
      id: "ins-" + Date.now(),
      handoverDocNo: "ACT-01",
      itemCode: "EQP-INS",
      equipmentName: document.getElementById("f-ins-name").value,
      serialNumber: document.getElementById("f-ins-serial").value,
      organization: document.getElementById("f-ins-org").value,
      region: "Toshkent",
      status: document.getElementById("f-ins-status").value,
      installationDate: document.getElementById("f-ins-date").value || "-",
      commissioningDate: document.getElementById("f-ins-comm").value || "-",
      commissioningCertNo: "COM-" + Date.now().toString().slice(-4),
      supplierSpecialist: document.getElementById("f-ins-spec").value || "-",
      trainingCompleted: "Ha / Yes",
      remarks: "Muntazam ravishda ishlatilmoqda"
    };
    appState.installations.push(newIns);
    logAudit("Установка", "CREATE", newIns.serialNumber, `O'rnatish qaydi yaratildi: ${newIns.equipmentName}`);
  } else if (mod === "monitoring") {
    const newMon = {
      id: "mon-" + Date.now(),
      monitoringDate: document.getElementById("f-mon-date").value,
      monitoredOrg: document.getElementById("f-mon-org").value,
      region: "Samarqand viloyati",
      equipmentName: document.getElementById("f-mon-name").value,
      serialNumber: document.getElementById("f-mon-serial").value,
      inventoryNumber: "INV-MON-01",
      actualLocation: "Laboratoriya xonasi",
      physicalAvailability: "Mavjud",
      commissioningStatus: "Foydalanishda",
      technicalCondition: document.getElementById("f-mon-cond").value,
      identifiedDefects: document.getElementById("f-mon-defects").value || "Nuqsonlar yo'q",
      photoEvidence: "Mavjud",
      recommendations: document.getElementById("f-mon-recom").value || "Kafolat muddati nazorat qilinsin",
      monitoringResult: "Ijobiy",
      monitoringOfficer: "GRP Monitoring mutaxassisi"
    };
    appState.monitoringRecords.push(newMon);
    logAudit("Мониторинг", "CREATE", newMon.serialNumber, `Monitoring tashrifi kiritildi: ${newMon.monitoredOrg}`);
  } else {
    // Inventory
    const newInv = {
      id: "inv-" + Date.now(),
      code: document.getElementById("f-inv-code").value,
      description: document.getElementById("f-inv-desc").value,
      serialNumber: document.getElementById("f-inv-serial").value,
      inventoryNumber: document.getElementById("f-inv-num").value || "INV-" + Date.now().toString().slice(-4),
      unit: document.getElementById("f-inv-unit").value,
      region: document.getElementById("f-inv-reg").value,
      institution: document.getElementById("f-inv-inst").value,
      location: "Bosh bino",
      responsiblePerson: document.getElementById("f-inv-resp").value,
      bookQty: Number(document.getElementById("f-inv-bqty").value) || 0,
      actualQty: Number(document.getElementById("f-inv-aqty").value) || 0,
      unitPrice: Number(document.getElementById("f-inv-price").value) || 0,
      currency: "USD",
      condition: document.getElementById("f-inv-cond").value,
      handoverDoc: "ACT-01",
      transferNo: "TRF-01",
      remarks: "Auditdan o'tkazildi"
    };
    appState.inventoryItems.push(newInv);
    logAudit("Инвентаризация", "CREATE", newInv.code, `Inventarizatsiya obyekti kiritildi: ${newInv.description}`);
  }

  saveState();
  dom.itemModal.classList.add("hidden");
  renderAll();
  showToast("Yangi yozuv muvaffaqiyatli saqlandi!");
}

// ====================================================================
// PRINT PREVIEW GENERATORS (A4 OFFICIAL FORMS)
// ====================================================================
function openPrintPreview(title, contentHtml) {
  dom.printModalTitle.textContent = title;
  dom.printableSheetContent.innerHTML = contentHtml;
  dom.printModal.classList.remove("hidden");
}

window.printSingleHandover = function(id) {
  const h = appState.handovers.find(x => x.id === id) || appState.handovers[0];
  const meta = appState.projectMetadata;
  const html = `
    <div class="print-paper">
      <div class="doc-letterhead">
        <div style="font-size:8pt; font-weight:bold; letter-spacing:1px; color:#163a5f; margin-bottom:4px;">
          PEMS-UZ &bull; PROJECT EQUIPMENT MANAGEMENT SYSTEM &bull; UZBEKISTAN
        </div>
        <h2>ASOSIY VOSITALAR VA TOVAR-MODDIY BOYLIKLARNI TOPSHIRISH-QABUL QILISH DALOLATNOMASI</h2>
        <h3>CERTIFICATE OF HANDOVER AND ACCEPTANCE № ${h.docNo} (ADB / AIIB)</h3>
        <p style="font-size:8pt; color:#64748b; margin-top:2px;">
          Система учёта оборудования (АБР / АБИИ) &bull; Dev: Abdiyev Salimkhon Sanjar ugli
        </p>
      </div>

      <table class="doc-meta-table">
        <tr>
          <td><strong>Dalolatnoma tuzilgan sana:</strong> ${h.date}</td>
          <td><strong>Hudud:</strong> ${h.region}</td>
        </tr>
        <tr>
          <td><strong>Loyiha nomi:</strong> ${meta.projectName}</td>
          <td><strong>Ichki Transfer №:</strong> ${h.transferNo}</td>
        </tr>
        <tr>
          <td><strong>Asos hujjati:</strong> ${meta.presidentialResolution}</td>
          <td><strong>Qabul qiluvchi STIR (ИНН):</strong> ${h.tin}</td>
        </tr>
      </table>

      <p style="margin: 12px 0; font-size: 10pt; line-height: 1.5; text-align: justify;">
        Ushbu dalolatnoma quyidagilar o'rtasida tuzildi: <strong>Группа реализации проекта (ГРП)</strong> nomidan topshiruvchi 
        <strong>${h.transferor || "GRP mutaxassisi"}</strong> bir tomondan, hamda <strong>${h.institution}</strong> nomidan qabul qilib oluvchi 
        <strong>${h.responsiblePerson}</strong> ikkinchi tomondan, quyidagi uskunalarni topshirdi va qabul qilib oldi:
      </p>

      <table class="doc-table">
        <thead>
          <tr>
            <th>№</th><th>Uskuna kodi</th><th>Uskuna nomi</th><th>Seriya raqami</th><th>O'lchov birligi</th><th>Miqdori</th><th>Qiymati USD</th><th>Qiymati UZS</th><th>Holati</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="text-align:center;">1</td>
            <td><code>${h.itemCode}</code></td>
            <td><strong>${h.itemDescription}</strong></td>
            <td><code>${h.serialNumber}</code></td>
            <td style="text-align:center;">${h.unit}</td>
            <td style="text-align:center; font-weight:bold;">${h.quantity}</td>
            <td style="text-align:right;">$${formatNumber(h.totalValueUsd)}</td>
            <td style="text-align:right;">${formatNumber(h.totalValueUzs)}</td>
            <td>${h.condition}</td>
          </tr>
        </tbody>
      </table>

      <p style="margin: 14px 0; font-size: 10pt; line-height: 1.5;">
        Qabul qilib oluvchi tomon uskunalarning texnik butunligini, pasport va texnik hujjatlarini to'liq tekshirib oldi. Topshirilgan mulk bo'yicha e'tirozlar mavjud emas.
      </p>

      <div class="signature-grid" style="margin-top: 40px;">
        <div class="sig-block">
          <strong>TOPSHIRDI (ГРП):</strong><br/><br/>
          F.I.Sh: <strong>${h.transferor || "GRP vakili"}</strong><br/>
          Imzo: _____________________<br/>
          M.O'. (Muhr o'rni)
        </div>
        <div class="sig-block">
          <strong>QABUL QILIB OLDI (Tashkilot):</strong><br/><br/>
          F.I.Sh: <strong>${h.responsiblePerson}</strong><br/>
          Imzo: _____________________<br/>
          M.O'. (Muhr o'rni)
        </div>
      </div>
    </div>
  `;
  openPrintPreview(`Акт передачи № ${h.docNo}`, html);
};

function printTeoReport() {
  openPrintPreview("Перечень оборудования по ТЭО", generateTeoReportHtml());
}

function printContractsReport() {
  openPrintPreview("Реестр контрактов и закупок", generateContractsReportHtml());
}

function printReceiptsReport() {
  openPrintPreview("Реестр поступлений и УзАСБО", generateReceiptsReportHtml());
}

function printHandoversReport() {
  openPrintPreview("Реестр актов передачи", generateHandoverReportHtml());
}

function printInstallationsReport() {
  openPrintPreview("Отчёт о монтаже и вводе в эксплуатацию", generateInstallationsReportHtml());
}

function printMonitoringReport() {
  openPrintPreview("Отчёт мониторинга ГРП", generateMonitoringReportHtml());
}

function printAlertsReport() {
  openPrintPreview("Перечень дефектов и рекламаций", generateDiscrepanciesReportHtml());
}

function printInventorySheet() {
  openPrintPreview("Инвентаризационная опись (А4)", generateInventoryReportHtml());
}

function printCommitteeProtocol() {
  openPrintPreview("Итоговый протокол инвентаризационной комиссии", generateFinalProjectClosureReportHtml());
}

// Toast
function showToast(msg) {
  if (!dom.toast) return;
  dom.toast.textContent = msg;
  dom.toast.classList.remove("hidden");
  setTimeout(() => {
    dom.toast.classList.add("hidden");
  }, 3500);
}
