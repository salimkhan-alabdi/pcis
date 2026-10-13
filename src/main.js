import {
  UZBEKISTAN_REGIONS,
  ITEM_CONDITIONS,
  UNITS,
  DEFAULT_PROJECT_METADATA,
  DEFAULT_INVENTORY_ITEMS,
  DEFAULT_DISCREPANCIES_EXTRAS,
  DEFAULT_COMMITTEE_DATA,
  DEFAULT_HANDOVER_RECORDS
} from "./data.js";
import { exportToExcel, parseExcelFile } from "./excelExporter.js";

// Storage Keys
const STORAGE_KEY = "uz_project_inventory_state_v1";

// Application State
let appState = loadState();

function loadState() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      const parsed = JSON.parse(saved);
      return {
        projectMetadata: { ...DEFAULT_PROJECT_METADATA, ...(parsed.projectMetadata || {}) },
        items: parsed.items || DEFAULT_INVENTORY_ITEMS,
        discrepanciesExtras: { ...DEFAULT_DISCREPANCIES_EXTRAS, ...(parsed.discrepanciesExtras || {}) },
        committee: { ...DEFAULT_COMMITTEE_DATA, ...(parsed.committee || {}) },
        handoverRecords: parsed.handoverRecords || DEFAULT_HANDOVER_RECORDS
      };
    }
  } catch (e) {
    console.error("Failed to load state from localStorage:", e);
  }
  return {
    projectMetadata: { ...DEFAULT_PROJECT_METADATA },
    items: [...DEFAULT_INVENTORY_ITEMS],
    discrepanciesExtras: { ...DEFAULT_DISCREPANCIES_EXTRAS },
    committee: JSON.parse(JSON.stringify(DEFAULT_COMMITTEE_DATA)),
    handoverRecords: [...DEFAULT_HANDOVER_RECORDS]
  };
}

function saveState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(appState));
  } catch (e) {
    console.error("Failed to save state:", e);
  }
}

// UI Elements
const dom = {
  // Tabs
  tabBtns: document.querySelectorAll(".tab-btn"),
  tabContents: document.querySelectorAll(".tab-content"),
  badgeTotalItems: document.getElementById("badge-total-items"),
  badgeDiscCount: document.getElementById("badge-discrepancies-count"),
  badgeHoCount: document.getElementById("badge-handover-count"),

  // Meta Inputs
  metaProjectName: document.getElementById("meta-project-name"),
  metaOrderNo: document.getElementById("meta-order-no"),
  metaInvDate: document.getElementById("meta-inv-date"),
  metaUsdRate: document.getElementById("meta-usd-rate"),

  // Top Actions
  btnExportExcel: document.getElementById("btn-export-excel"),
  excelFileInput: document.getElementById("excel-file-input"),
  btnAddItemModal: document.getElementById("btn-add-item-modal"),
  btnResetData: document.getElementById("btn-reset-data"),

  // KPI
  kpiTotalItems: document.getElementById("kpi-total-items"),
  kpiTotalUnits: document.getElementById("kpi-total-units"),
  kpiBookValUzs: document.getElementById("kpi-book-val-uzs"),
  kpiBookValUsd: document.getElementById("kpi-book-val-usd"),
  kpiActualValUzs: document.getElementById("kpi-actual-val-uzs"),
  kpiActualValUsd: document.getElementById("kpi-actual-val-usd"),
  kpiDiffVal: document.getElementById("kpi-diff-val"),
  kpiDiffSub: document.getElementById("kpi-diff-sub"),
  kpiDiffIconWrap: document.getElementById("kpi-diff-icon-wrap"),
  kpiSurplusCount: document.getElementById("kpi-surplus-count"),
  kpiShortageCount: document.getElementById("kpi-shortage-count"),

  // Tab 1: Inventory
  searchInventory: document.getElementById("search-inventory"),
  btnClearSearch: document.getElementById("btn-clear-search"),
  filterRegion: document.getElementById("filter-region"),
  filterDiscrepancy: document.getElementById("filter-discrepancy"),
  filterCurrency: document.getElementById("filter-currency"),
  filterCondition: document.getElementById("filter-condition"),
  inventoryTbody: document.getElementById("inventory-tbody"),
  inventoryShownCount: document.getElementById("inventory-shown-count"),
  btnPrintInventoryTable: document.getElementById("btn-print-inventory-table"),

  // Tab 2: Regional
  regionalTbody: document.getElementById("regional-tbody"),
  regionalTfoot: document.getElementById("regional-tfoot"),
  regionalVisualBars: document.getElementById("regional-visual-bars"),
  btnPrintRegional: document.getElementById("btn-print-regional"),

  // Tab 3: Discrepancies
  discSummarySurplus: document.getElementById("disc-summary-surplus"),
  discSummaryShortage: document.getElementById("disc-summary-shortage"),
  discrepancyCountBadge: document.getElementById("discrepancy-count-badge"),
  discrepanciesTbody: document.getElementById("discrepancies-tbody"),

  // Tab 4: Handover
  btnCreateHandoverModal: document.getElementById("btn-create-handover-modal"),
  handoverTbody: document.getElementById("handover-tbody"),

  // Tab 5: Committee
  commOrg: document.getElementById("comm-org"),
  commProject: document.getElementById("comm-project"),
  commOrderNo: document.getElementById("comm-order-no"),
  commStartDate: document.getElementById("comm-start-date"),
  commEndDate: document.getElementById("comm-end-date"),
  commMembersTbody: document.getElementById("comm-members-tbody"),
  commConclusion: document.getElementById("comm-conclusion"),
  commDiscrepancies: document.getElementById("comm-discrepancies"),
  commRecommended: document.getElementById("comm-recommended"),
  commDecision: document.getElementById("comm-decision"),
  btnSaveCommittee: document.getElementById("btn-save-committee"),
  btnPrintCommitteeProtocol: document.getElementById("btn-print-committee-protocol"),

  // Modals
  itemModal: document.getElementById("item-modal"),
  modalItemTitle: document.getElementById("modal-item-title"),
  itemForm: document.getElementById("item-form"),
  btnCloseItemModal: document.getElementById("btn-close-item-modal"),
  btnCancelItem: document.getElementById("btn-cancel-item"),

  // Item Form Fields
  fItemId: document.getElementById("item-id"),
  fCode: document.getElementById("f-code"),
  fUnit: document.getElementById("f-unit"),
  fDesc: document.getElementById("f-desc"),
  fRegion: document.getElementById("f-region"),
  fInstitution: document.getElementById("f-institution"),
  fLocation: document.getElementById("f-location"),
  fResponsible: document.getElementById("f-responsible"),
  fCondition: document.getElementById("f-condition"),
  fBookQty: document.getElementById("f-book-qty"),
  fActualQty: document.getElementById("f-actual-qty"),
  fUnitPrice: document.getElementById("f-unit-price"),
  fCurrency: document.getElementById("f-currency"),
  fHandoverDoc: document.getElementById("f-handover-doc"),
  fRemarks: document.getElementById("f-remarks"),
  prevSurplus: document.getElementById("prev-surplus"),
  prevShortage: document.getElementById("prev-shortage"),
  prevBookVal: document.getElementById("prev-book-val"),
  prevActualVal: document.getElementById("prev-actual-val"),
  prevDiff: document.getElementById("prev-diff"),

  // Handover Modal
  handoverModal: document.getElementById("handover-modal"),
  btnCloseHandoverModal: document.getElementById("btn-close-handover-modal"),
  btnCancelHandover: document.getElementById("btn-cancel-handover"),
  handoverForm: document.getElementById("handover-form"),
  hoFDate: document.getElementById("ho-f-date"),
  hoFDocNo: document.getElementById("ho-f-docno"),
  hoFSelectItem: document.getElementById("ho-f-select-item"),
  hoFCode: document.getElementById("ho-f-code"),
  hoFUnit: document.getElementById("ho-f-unit"),
  hoFDesc: document.getElementById("ho-f-desc"),
  hoFQty: document.getElementById("ho-f-qty"),
  hoFCondition: document.getElementById("ho-f-condition"),
  hoFRegion: document.getElementById("ho-f-region"),
  hoFInstitution: document.getElementById("ho-f-institution"),
  hoFRecipient: document.getElementById("ho-f-recipient"),
  hoFTransferor: document.getElementById("ho-f-transferor"),
  hoFSupporting: document.getElementById("ho-f-supporting"),
  hoFRemarks: document.getElementById("ho-f-remarks"),

  // Print Modal
  printModal: document.getElementById("print-modal"),
  printModalTitle: document.getElementById("print-modal-title"),
  printableContent: document.getElementById("printable-sheet-content"),
  btnClosePrintModal: document.getElementById("btn-close-print-modal"),
  btnExecutePrint: document.getElementById("btn-execute-print"),

  // Toast
  toast: document.getElementById("toast")
};

// Utilities
function showToast(message, duration = 3000) {
  dom.toast.textContent = message;
  dom.toast.classList.remove("hidden");
  setTimeout(() => {
    dom.toast.classList.add("hidden");
  }, duration);
}

function formatNumber(num) {
  if (num === null || num === undefined || isNaN(num)) return "0";
  return Number(num).toLocaleString();
}

function formatCurrency(amount, currency = "UZS") {
  const formatted = formatNumber(amount);
  return currency === "USD" ? `$${formatted}` : `${formatted} UZS`;
}

// ====================================================================
// INITIALIZATION
// ====================================================================
function init() {
  populateSelectOptions();
  bindMetaInputs();
  bindTabs();
  bindModals();
  bindFilters();
  bindActions();

  renderAll();
}

function populateSelectOptions() {
  // Region filter & form dropdowns
  UZBEKISTAN_REGIONS.forEach((reg) => {
    const optFilter = document.createElement("option");
    optFilter.value = reg;
    optFilter.textContent = reg;
    dom.filterRegion.appendChild(optFilter);

    const optItem = document.createElement("option");
    optItem.value = reg;
    optItem.textContent = reg;
    dom.fRegion.appendChild(optItem);

    const optHo = document.createElement("option");
    optHo.value = reg;
    optHo.textContent = reg;
    dom.hoFRegion.appendChild(optHo);
  });

  // Units
  UNITS.forEach((unit) => {
    const opt = document.createElement("option");
    opt.value = unit;
    opt.textContent = unit;
    dom.fUnit.appendChild(opt);
  });

  // Conditions
  ITEM_CONDITIONS.forEach((cond) => {
    const opt = document.createElement("option");
    opt.value = cond;
    opt.textContent = cond;
    dom.fCondition.appendChild(opt);
  });
}

function bindMetaInputs() {
  dom.metaProjectName.value = appState.projectMetadata.projectName || "";
  dom.metaOrderNo.value = appState.projectMetadata.orderNo || "";
  dom.metaInvDate.value = appState.projectMetadata.inventoryDate || "";
  dom.metaUsdRate.value = appState.projectMetadata.usdRate || 12850;

  dom.metaProjectName.addEventListener("input", (e) => {
    appState.projectMetadata.projectName = e.target.value;
    saveState();
  });

  dom.metaOrderNo.addEventListener("input", (e) => {
    appState.projectMetadata.orderNo = e.target.value;
    saveState();
  });

  dom.metaInvDate.addEventListener("change", (e) => {
    appState.projectMetadata.inventoryDate = e.target.value;
    saveState();
  });

  dom.metaUsdRate.addEventListener("input", (e) => {
    const rate = Number(e.target.value) || 12850;
    appState.projectMetadata.usdRate = rate;
    saveState();
    renderKPIs();
    renderRegionalTable();
  });
}

function bindTabs() {
  dom.tabBtns.forEach((btn) => {
    btn.addEventListener("click", () => {
      const targetTabId = btn.dataset.tab;
      dom.tabBtns.forEach((b) => b.classList.remove("active"));
      dom.tabContents.forEach((c) => c.classList.remove("active"));

      btn.classList.add("active");
      const targetContent = document.getElementById(targetTabId);
      if (targetContent) targetContent.classList.add("active");

      window.scrollTo({ top: 0, behavior: "smooth" });

      if (targetTabId === "tab-regional") renderRegionalTable();
      if (targetTabId === "tab-discrepancies") renderDiscrepanciesTable();
      if (targetTabId === "tab-handover") renderHandoverTable();
      if (targetTabId === "tab-committee") renderCommitteeTab();
    });
  });
}

function bindFilters() {
  dom.searchInventory.addEventListener("input", renderInventoryTable);
  dom.btnClearSearch.addEventListener("click", () => {
    dom.searchInventory.value = "";
    renderInventoryTable();
  });
  dom.filterRegion.addEventListener("change", renderInventoryTable);
  dom.filterDiscrepancy.addEventListener("change", renderInventoryTable);
  dom.filterCurrency.addEventListener("change", renderInventoryTable);
  dom.filterCondition.addEventListener("change", renderInventoryTable);
}

function bindActions() {
  // Export Excel
  dom.btnExportExcel.addEventListener("click", () => {
    try {
      exportToExcel(appState);
      showToast("Excel fayli muvaffaqiyatli yaratildi! (Project_Inventory_UZ_EN.xlsx)");
    } catch (err) {
      console.error(err);
      alert("Eksportda xatolik yuz berdi: " + err.message);
    }
  });

  // Import Excel
  dom.excelFileInput.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const parsed = await parseExcelFile(file);
      if (parsed.items && parsed.items.length > 0) {
        if (confirm(`Exceldan ${parsed.items.length} ta pozitsiya topildi. Mavjud ro'yxatni yangilamoqchimisiz?`)) {
          appState.items = parsed.items;
          saveState();
          renderAll();
          showToast(`${parsed.items.length} ta tovar muvaffaqiyatli yuklandi!`);
        }
      } else {
        alert("Excel faylidan yaroqli tovarlar topilmadi. Ustunlar formatini tekshiring.");
      }
    } catch (err) {
      console.error(err);
      alert("Faylni o'qishda xatolik: " + err.message);
    }
    e.target.value = "";
  });

  // Reset to Sample Data
  dom.btnResetData.addEventListener("click", () => {
    if (confirm("Namunaviy ma'lumotlarni qayta tiklashni istaysizmi? Barcha kiritilgan o'zgarishlar asl holiga qaytadi.")) {
      localStorage.removeItem(STORAGE_KEY);
      appState = loadState();
      dom.metaProjectName.value = appState.projectMetadata.projectName;
      dom.metaOrderNo.value = appState.projectMetadata.orderNo;
      dom.metaInvDate.value = appState.projectMetadata.inventoryDate;
      dom.metaUsdRate.value = appState.projectMetadata.usdRate;
      renderAll();
      showToast("Namunaviy ma'lumotlar qayta tiklandi!");
    }
  });

  // Print Inventory Table
  dom.btnPrintInventoryTable.addEventListener("click", () => {
    openPrintModal("LOYIHA YAKUNIY INVENTARIZATSIYASI REYESTRI", generateInventoryPrintHtml());
  });

  // Print Regional Summary
  dom.btnPrintRegional.addEventListener("click", () => {
    openPrintModal("HUDUDLAR BO'YICHA JAMLANMA REYESTRI", generateRegionalPrintHtml());
  });

  // Print Committee Protocol
  dom.btnPrintCommitteeProtocol.addEventListener("click", () => {
    openPrintModal("INVENTARIZATSIYA KOMISSIYASI RASMIY BAYONNOMASI", generateCommitteePrintHtml());
  });

  // Save Committee
  dom.btnSaveCommittee.addEventListener("click", () => {
    appState.committee.organization = dom.commOrg.value;
    appState.committee.projectName = dom.commProject.value;
    appState.committee.orderNo = dom.commOrderNo.value;
    appState.committee.startDate = dom.commStartDate.value;
    appState.committee.endDate = dom.commEndDate.value;
    appState.committee.conclusion = dom.commConclusion.value;
    appState.committee.discrepanciesIdentified = dom.commDiscrepancies.value;
    appState.committee.recommendedActions = dom.commRecommended.value;
    appState.committee.managementDecision = dom.commDecision.value;

    saveState();
    showToast("Komissiya ma'lumotlari muvaffaqiyatli saqlandi!");
  });
}

function bindModals() {
  // Add item modal open
  dom.btnAddItemModal.addEventListener("click", () => {
    openItemModal(null);
  });

  dom.btnCloseItemModal.addEventListener("click", closeItemModal);
  dom.btnCancelItem.addEventListener("click", closeItemModal);

  // Live calculation in item modal
  const updateModalCalc = () => {
    const bookQty = Number(dom.fBookQty.value) || 0;
    const actualQty = Number(dom.fActualQty.value) || 0;
    const price = Number(dom.fUnitPrice.value) || 0;
    const curr = dom.fCurrency.value;

    const surplus = Math.max(actualQty - bookQty, 0);
    const shortage = Math.max(bookQty - actualQty, 0);
    const bookVal = bookQty * price;
    const actualVal = actualQty * price;
    const diff = actualVal - bookVal;

    dom.prevSurplus.textContent = `+${surplus}`;
    dom.prevShortage.textContent = `-${shortage}`;
    dom.prevBookVal.textContent = formatCurrency(bookVal, curr);
    dom.prevActualVal.textContent = formatCurrency(actualVal, curr);
    dom.prevDiff.textContent = `${diff >= 0 ? "+" : ""}${formatCurrency(diff, curr)}`;
    dom.prevDiff.className = diff < 0 ? "text-danger" : (diff > 0 ? "text-success" : "");
  };

  [dom.fBookQty, dom.fActualQty, dom.fUnitPrice, dom.fCurrency].forEach((inp) => {
    inp.addEventListener("input", updateModalCalc);
  });

  // Save Item Form
  dom.itemForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const id = dom.fItemId.value;
    const isEdit = Boolean(id);

    const itemData = {
      id: isEdit ? id : "item-" + Date.now(),
      code: dom.fCode.value.trim(),
      description: dom.fDesc.value.trim(),
      unit: dom.fUnit.value,
      region: dom.fRegion.value,
      institution: dom.fInstitution.value.trim(),
      location: dom.fLocation.value.trim(),
      responsiblePerson: dom.fResponsible.value.trim(),
      bookQty: Number(dom.fBookQty.value) || 0,
      actualQty: Number(dom.fActualQty.value) || 0,
      unitPrice: Number(dom.fUnitPrice.value) || 0,
      currency: dom.fCurrency.value,
      condition: dom.fCondition.value,
      handoverDoc: dom.fHandoverDoc.value.trim(),
      remarks: dom.fRemarks.value.trim()
    };

    if (isEdit) {
      const idx = appState.items.findIndex((it) => it.id === id);
      if (idx !== -1) appState.items[idx] = itemData;
    } else {
      appState.items.unshift(itemData);
    }

    saveState();
    closeItemModal();
    renderAll();
    showToast(isEdit ? "Tovar ma'lumotlari yangilandi!" : "Yangi tovar qo'shildi!");
  });

  // Handover Modal
  dom.btnCreateHandoverModal.addEventListener("click", () => {
    openHandoverModal(null);
  });

  dom.btnCloseHandoverModal.addEventListener("click", closeHandoverModal);
  dom.btnCancelHandover.addEventListener("click", closeHandoverModal);

  dom.hoFSelectItem.addEventListener("change", (e) => {
    const selectedId = e.target.value;
    if (!selectedId) return;
    const it = appState.items.find((item) => item.id === selectedId);
    if (it) {
      dom.hoFCode.value = it.code;
      dom.hoFDesc.value = it.description;
      dom.hoFUnit.value = it.unit;
      dom.hoFQty.value = it.actualQty;
      dom.hoFRegion.value = it.region;
      dom.hoFInstitution.value = it.institution;
      dom.hoFTransferor.value = it.responsiblePerson;
      dom.hoFCondition.value = it.condition;
    }
  });

  dom.handoverForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const newAct = {
      id: "ho-" + Date.now(),
      date: dom.hoFDate.value,
      docNo: dom.hoFDocNo.value.trim(),
      itemCode: dom.hoFCode.value.trim(),
      itemDescription: dom.hoFDesc.value.trim(),
      quantity: Number(dom.hoFQty.value) || 0,
      unit: dom.hoFUnit.value.trim(),
      region: dom.hoFRegion.value,
      institution: dom.hoFInstitution.value.trim(),
      recipientName: dom.hoFRecipient.value.trim(),
      transferor: dom.hoFTransferor.value.trim(),
      condition: dom.hoFCondition.value,
      supportingDoc: dom.hoFSupporting.value.trim(),
      signatures: "Imzolangan / Signed",
      remarks: dom.hoFRemarks.value.trim()
    };

    appState.handoverRecords.unshift(newAct);
    saveState();
    closeHandoverModal();
    renderHandoverTable();
    showToast("Topshirish-qabul qilish akti muvaffaqiyatli yaratildi!");
  });

  // Print Modal
  dom.btnClosePrintModal.addEventListener("click", closePrintModal);
  dom.btnExecutePrint.addEventListener("click", () => {
    window.print();
  });
}

// Modal open/close helpers
function openItemModal(itemToEdit = null) {
  if (itemToEdit) {
    dom.modalItemTitle.textContent = "Tovarni tahrirlash / Edit Item";
    dom.fItemId.value = itemToEdit.id;
    dom.fCode.value = itemToEdit.code;
    dom.fDesc.value = itemToEdit.description;
    dom.fUnit.value = itemToEdit.unit || "dona / pcs";
    dom.fRegion.value = itemToEdit.region;
    dom.fInstitution.value = itemToEdit.institution;
    dom.fLocation.value = itemToEdit.location || "";
    dom.fResponsible.value = itemToEdit.responsiblePerson;
    dom.fCondition.value = itemToEdit.condition || "Yaxshi / Good";
    dom.fBookQty.value = itemToEdit.bookQty;
    dom.fActualQty.value = itemToEdit.actualQty;
    dom.fUnitPrice.value = itemToEdit.unitPrice;
    dom.fCurrency.value = itemToEdit.currency || "UZS";
    dom.fHandoverDoc.value = itemToEdit.handoverDoc || "";
    dom.fRemarks.value = itemToEdit.remarks || "";
  } else {
    dom.modalItemTitle.textContent = "Yangi tovar qo'shish / Add Inventory Item";
    dom.fItemId.value = "";
    dom.itemForm.reset();
    dom.fCode.value = `ITM-2026-${String(appState.items.length + 1).padStart(3, "0")}`;
    dom.fBookQty.value = "1";
    dom.fActualQty.value = "1";
    dom.fUnitPrice.value = "1000000";
  }

  // trigger preview calculation
  dom.fBookQty.dispatchEvent(new Event("input"));
  dom.itemModal.classList.remove("hidden");
}

function closeItemModal() {
  dom.itemModal.classList.add("hidden");
}

function openHandoverModal(prefillItem = null) {
  dom.hoFSelectItem.innerHTML = '<option value="">-- Tovarni ro\'yxatdan tanlash --</option>';
  appState.items.forEach((it) => {
    const opt = document.createElement("option");
    opt.value = it.id;
    opt.textContent = `${it.code} - ${it.description} (${it.region})`;
    dom.hoFSelectItem.appendChild(opt);
  });

  const today = new Date().toISOString().split("T")[0];
  dom.hoFDate.value = today;
  dom.hoFDocNo.value = `TT-${String(appState.handoverRecords.length + 1).padStart(2, "0")}/2026`;

  if (prefillItem) {
    dom.hoFSelectItem.value = prefillItem.id;
    dom.hoFSelectItem.dispatchEvent(new Event("change"));
  } else {
    dom.hoFCode.value = "";
    dom.hoFDesc.value = "";
    dom.hoFQty.value = "1";
    dom.hoFInstitution.value = "";
    dom.hoFRecipient.value = "";
    dom.hoFTransferor.value = "";
    dom.hoFSupporting.value = appState.projectMetadata.orderNo || "";
    dom.hoFRemarks.value = "";
  }

  dom.handoverModal.classList.remove("hidden");
}

function closeHandoverModal() {
  dom.handoverModal.classList.add("hidden");
}

function openPrintModal(title, htmlContent) {
  dom.printModalTitle.textContent = title;
  dom.printableContent.innerHTML = htmlContent;
  dom.printModal.classList.remove("hidden");
}

function closePrintModal() {
  dom.printModal.classList.add("hidden");
}

// ====================================================================
// RENDERERS
// ====================================================================
function renderAll() {
  renderKPIs();
  renderInventoryTable();
  renderRegionalTable();
  renderDiscrepanciesTable();
  renderHandoverTable();
  renderCommitteeTab();
}

function renderKPIs() {
  const usdRate = Number(appState.projectMetadata.usdRate) || 12850;
  let totalItemsCount = appState.items.length;
  let totalUnits = 0;
  let bookValUzs = 0;
  let bookValUsd = 0;
  let actualValUzs = 0;
  let actualValUsd = 0;
  let surplusCount = 0;
  let shortageCount = 0;

  appState.items.forEach((it) => {
    const b = Number(it.bookQty) || 0;
    const a = Number(it.actualQty) || 0;
    const price = Number(it.unitPrice) || 0;

    totalUnits += a;
    if (a > b) surplusCount += (a - b);
    if (b > a) shortageCount += (b - a);

    if (it.currency === "USD") {
      bookValUsd += (b * price);
      actualValUsd += (a * price);
    } else {
      bookValUzs += (b * price);
      actualValUzs += (a * price);
    }
  });

  // Calculate equivalents
  const grandBookEqUzs = bookValUzs + (bookValUsd * usdRate);
  const grandActualEqUzs = actualValUzs + (actualValUsd * usdRate);
  const diffEqUzs = grandActualEqUzs - grandBookEqUzs;

  dom.badgeTotalItems.textContent = totalItemsCount;
  dom.kpiTotalItems.textContent = totalItemsCount;
  dom.kpiTotalUnits.textContent = `${formatNumber(totalUnits)} dona / birlik`;

  dom.kpiBookValUzs.textContent = `${formatNumber(bookValUzs)} UZS`;
  dom.kpiBookValUsd.textContent = `$${formatNumber(bookValUsd)} USD`;

  dom.kpiActualValUzs.textContent = `${formatNumber(actualValUzs)} UZS`;
  dom.kpiActualValUsd.textContent = `$${formatNumber(actualValUsd)} USD`;

  // Net difference
  dom.kpiDiffVal.textContent = `${diffEqUzs >= 0 ? "+" : ""}${formatNumber(diffEqUzs)} UZS`;
  if (diffEqUzs > 0) {
    dom.kpiDiffVal.className = "kpi-value text-success";
    dom.kpiDiffSub.textContent = "Ortiqcha qoldiq";
    dom.kpiDiffIconWrap.className = "kpi-icon-wrap bg-indigo";
  } else if (diffEqUzs < 0) {
    dom.kpiDiffVal.className = "kpi-value text-danger";
    dom.kpiDiffSub.textContent = "Kamomad qoldiq";
    dom.kpiDiffIconWrap.className = "kpi-icon-wrap bg-amber";
  } else {
    dom.kpiDiffVal.className = "kpi-value";
    dom.kpiDiffSub.textContent = "To'liq muvofiq (Farqsiz)";
    dom.kpiDiffIconWrap.className = "kpi-icon-wrap bg-navy";
  }

  dom.kpiSurplusCount.textContent = `+${surplusCount}`;
  dom.kpiShortageCount.textContent = `-${shortageCount}`;
}

// --------------------------------------------------------------------
// TAB 1: INVENTORY TABLE
// --------------------------------------------------------------------
function renderInventoryTable() {
  const searchQuery = (dom.searchInventory.value || "").toLowerCase().trim();
  const regionFilter = dom.filterRegion.value;
  const discrepancyFilter = dom.filterDiscrepancy.value;
  const currencyFilter = dom.filterCurrency.value;
  const conditionFilter = dom.filterCondition.value;

  const filtered = appState.items.filter((it) => {
    // Search
    if (searchQuery) {
      const match =
        (it.code && it.code.toLowerCase().includes(searchQuery)) ||
        (it.description && it.description.toLowerCase().includes(searchQuery)) ||
        (it.responsiblePerson && it.responsiblePerson.toLowerCase().includes(searchQuery)) ||
        (it.institution && it.institution.toLowerCase().includes(searchQuery)) ||
        (it.location && it.location.toLowerCase().includes(searchQuery));
      if (!match) return false;
    }

    // Region
    if (regionFilter && it.region !== regionFilter) return false;

    // Currency
    if (currencyFilter !== "all" && it.currency !== currencyFilter) return false;

    // Condition
    if (conditionFilter !== "all" && it.condition !== conditionFilter) return false;

    // Discrepancy filter
    const b = Number(it.bookQty) || 0;
    const a = Number(it.actualQty) || 0;
    const surplus = Math.max(a - b, 0);
    const shortage = Math.max(b - a, 0);

    if (discrepancyFilter === "exact" && (surplus > 0 || shortage > 0)) return false;
    if (discrepancyFilter === "discrepancy" && surplus === 0 && shortage === 0) return false;
    if (discrepancyFilter === "surplus" && surplus === 0) return false;
    if (discrepancyFilter === "shortage" && shortage === 0) return false;

    return true;
  });

  dom.inventoryShownCount.textContent = `${filtered.length} ta pozitsiya ko'rsatilmoqda`;
  dom.inventoryTbody.innerHTML = "";

  if (filtered.length === 0) {
    dom.inventoryTbody.innerHTML = `
      <tr>
        <td colspan="21" style="text-align: center; padding: 40px; color: var(--slate-500);">
          Hech qanday tovar topilmadi. Qidiruv yoki filtrlarni tekshiring.
        </td>
      </tr>
    `;
    return;
  }

  filtered.forEach((it, index) => {
    const tr = document.createElement("tr");

    const b = Number(it.bookQty) || 0;
    const a = Number(it.actualQty) || 0;
    const price = Number(it.unitPrice) || 0;
    const surplus = Math.max(a - b, 0);
    const shortage = Math.max(b - a, 0);
    const bookVal = b * price;
    const actualVal = a * price;
    const diff = actualVal - bookVal;

    if (surplus > 0) tr.classList.add("row-surplus");
    if (shortage > 0) tr.classList.add("row-shortage");

    const conditionBadgeClass = getConditionBadgeClass(it.condition);

    tr.innerHTML = `
      <td class="col-num">${index + 1}</td>
      <td style="font-family: var(--font-mono); font-weight: 700; color: var(--navy-primary);">${it.code || ""}</td>
      <td class="col-desc"><strong>${it.description || ""}</strong></td>
      <td><span class="badge" style="background: #e2e8f0;">${it.unit || "dona"}</span></td>
      <td>${it.region || ""}</td>
      <td>${it.institution || ""}</td>
      <td>${it.location || ""}</td>
      <td>${it.responsiblePerson || ""}</td>
      <td class="col-qty">${formatNumber(b)}</td>
      <td class="col-qty font-bold">${formatNumber(a)}</td>
      <td class="col-qty">${surplus > 0 ? `<span class="badge badge-surplus">+${surplus}</span>` : '<span style="color:#94a3b8;">0</span>'}</td>
      <td class="col-qty">${shortage > 0 ? `<span class="badge badge-shortage">-${shortage}</span>` : '<span style="color:#94a3b8;">0</span>'}</td>
      <td class="col-currency">${formatNumber(price)}</td>
      <td><span class="badge" style="background: ${it.currency === 'USD' ? '#dbeafe' : '#f1f5f9'}; font-weight:700;">${it.currency}</span></td>
      <td class="col-currency">${formatNumber(bookVal)}</td>
      <td class="col-currency font-bold">${formatNumber(actualVal)}</td>
      <td class="col-currency ${diff < 0 ? 'text-danger font-bold' : (diff > 0 ? 'text-success font-bold' : '')}">
        ${diff !== 0 ? (diff > 0 ? '+' : '') + formatNumber(diff) : '0'}
      </td>
      <td><span class="badge ${conditionBadgeClass}">${it.condition || "Yaxshi"}</span></td>
      <td style="font-family: var(--font-mono); font-size: 0.78rem;">${it.handoverDoc || "-"}</td>
      <td style="max-width: 180px; overflow: hidden; text-overflow: ellipsis;" title="${it.remarks || ''}">${it.remarks || "-"}</td>
      <td class="col-actions">
        <div class="action-buttons">
          <button class="btn-icon btn-edit-item" title="Tahrirlash / Edit" data-id="${it.id}">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
              <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
            </svg>
          </button>
          <button class="btn-icon" title="Topshirish akti yaratish" data-handover-id="${it.id}">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"></path>
              <rect x="8" y="2" width="8" height="4" rx="1" ry="1"></rect>
            </svg>
          </button>
          <button class="btn-icon delete btn-del-item" title="O'chirish / Delete" data-id="${it.id}">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="3 6 5 6 21 6"></polyline>
              <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
            </svg>
          </button>
        </div>
      </td>
    `;

    dom.inventoryTbody.appendChild(tr);
  });

  // Attach row action handlers
  dom.inventoryTbody.querySelectorAll(".btn-edit-item").forEach((btn) => {
    btn.addEventListener("click", () => {
      const it = appState.items.find((item) => item.id === btn.dataset.id);
      if (it) openItemModal(it);
    });
  });

  dom.inventoryTbody.querySelectorAll("[data-handover-id]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const it = appState.items.find((item) => item.id === btn.dataset.handoverId);
      if (it) openHandoverModal(it);
    });
  });

  dom.inventoryTbody.querySelectorAll(".btn-del-item").forEach((btn) => {
    btn.addEventListener("click", () => {
      const it = appState.items.find((item) => item.id === btn.dataset.id);
      if (it && confirm(`"${it.code} - ${it.description}" tovarini ro'yxatdan o'chirishni tasdiqlaysizmi?`)) {
        appState.items = appState.items.filter((item) => item.id !== it.id);
        saveState();
        renderAll();
        showToast("Tovar o'chirildi!");
      }
    });
  });
}

function getConditionBadgeClass(condition) {
  if (!condition) return "badge-good";
  if (condition.includes("Yangi")) return "badge-new";
  if (condition.includes("Yaxshi")) return "badge-good";
  if (condition.includes("Qoniqarli")) return "badge-satisfactory";
  if (condition.includes("Ta'mir")) return "badge-repair";
  if (condition.includes("Yaroqsiz")) return "badge-defective";
  return "badge-good";
}

// --------------------------------------------------------------------
// TAB 2: REGIONAL SUMMARY
// --------------------------------------------------------------------
function renderRegionalTable() {
  const usdRate = Number(appState.projectMetadata.usdRate) || 12850;
  dom.regionalTbody.innerHTML = "";
  dom.regionalVisualBars.innerHTML = "";

  let grandCount = 0;
  let grandBookQty = 0;
  let grandActualQty = 0;
  let grandSurplus = 0;
  let grandShortage = 0;
  let grandUzs = 0;
  let grandUsd = 0;

  const regionalStats = [];

  UZBEKISTAN_REGIONS.forEach((region, idx) => {
    const regItems = appState.items.filter((it) => it.region === region);
    const count = regItems.length;
    let bQty = 0;
    let aQty = 0;
    let surplus = 0;
    let shortage = 0;
    let uzsVal = 0;
    let usdVal = 0;

    regItems.forEach((it) => {
      const b = Number(it.bookQty) || 0;
      const a = Number(it.actualQty) || 0;
      const price = Number(it.unitPrice) || 0;

      bQty += b;
      aQty += a;
      surplus += Math.max(a - b, 0);
      shortage += Math.max(b - a, 0);

      const actVal = a * price;
      if (it.currency === "USD") {
        usdVal += actVal;
      } else {
        uzsVal += actVal;
      }
    });

    const totalEqUzs = uzsVal + (usdVal * usdRate);

    grandCount += count;
    grandBookQty += bQty;
    grandActualQty += aQty;
    grandSurplus += surplus;
    grandShortage += shortage;
    grandUzs += uzsVal;
    grandUsd += usdVal;

    regionalStats.push({
      region,
      count,
      bQty,
      aQty,
      surplus,
      shortage,
      uzsVal,
      usdVal,
      totalEqUzs
    });

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="col-num">${idx + 1}</td>
      <td style="font-weight: 700; color: var(--navy-dark);">${region}</td>
      <td class="col-qty"><span class="badge ${count > 0 ? 'badge-new' : 'badge-match'}">${count}</span></td>
      <td class="col-qty">${formatNumber(bQty)}</td>
      <td class="col-qty font-bold">${formatNumber(aQty)}</td>
      <td class="col-qty">${surplus > 0 ? `<span class="badge badge-surplus">+${surplus}</span>` : '0'}</td>
      <td class="col-qty">${shortage > 0 ? `<span class="badge badge-shortage">-${shortage}</span>` : '0'}</td>
      <td class="col-currency">${formatNumber(uzsVal)}</td>
      <td class="col-currency">${usdVal > 0 ? `$${formatNumber(usdVal)}` : '$0'}</td>
      <td class="col-actions">
        <button class="btn btn-sm btn-outline btn-filter-region-jump" data-region="${region}" title="Ushbu hudud tovarlarini ko'rish">
          Filtrlash
        </button>
      </td>
    `;
    dom.regionalTbody.appendChild(tr);
  });

  // Grand Total Row
  dom.regionalTfoot.innerHTML = `
    <tr>
      <td colspan="2" style="font-size: 0.95rem; text-transform: uppercase;">JAMI / GRAND TOTAL</td>
      <td class="col-qty">${grandCount}</td>
      <td class="col-qty">${formatNumber(grandBookQty)}</td>
      <td class="col-qty font-bold">${formatNumber(grandActualQty)}</td>
      <td class="col-qty"><span class="badge badge-surplus">+${grandSurplus}</span></td>
      <td class="col-qty"><span class="badge badge-shortage">-${grandShortage}</span></td>
      <td class="col-currency font-bold">${formatNumber(grandUzs)} UZS</td>
      <td class="col-currency font-bold">$${formatNumber(grandUsd)} USD</td>
      <td></td>
    </tr>
  `;

  // Render Visual Bars
  const grandTotalEq = grandUzs + (grandUsd * usdRate);
  const sortedByVal = [...regionalStats].sort((a, b) => b.totalEqUzs - a.totalEqUzs);

  sortedByVal.slice(0, 7).forEach((st) => {
    const percent = grandTotalEq > 0 ? Math.round((st.totalEqUzs / grandTotalEq) * 100) : 0;
    const barEl = document.createElement("div");
    barEl.className = "reg-bar-item";
    barEl.innerHTML = `
      <div class="reg-bar-label" title="${st.region}">${st.region.split("/")[0]}</div>
      <div class="reg-bar-track">
        <div class="reg-bar-fill" style="width: ${percent}%;"></div>
      </div>
      <div class="reg-bar-val">${percent}% (${formatNumber(Math.round(st.totalEqUzs / 1e6))} mln)</div>
    `;
    dom.regionalVisualBars.appendChild(barEl);
  });

  // Attach Jump filter
  dom.regionalTbody.querySelectorAll(".btn-filter-region-jump").forEach((btn) => {
    btn.addEventListener("click", () => {
      const reg = btn.dataset.region;
      dom.filterRegion.value = reg;
      // Switch to Inventory tab
      document.getElementById("nav-inventory").click();
      renderInventoryTable();
    });
  });
}

// --------------------------------------------------------------------
// TAB 3: DISCREPANCIES TABLE
// --------------------------------------------------------------------
function renderDiscrepanciesTable() {
  const discrepancyItems = appState.items.filter((it) => {
    const b = Number(it.bookQty) || 0;
    const a = Number(it.actualQty) || 0;
    return (a - b) !== 0;
  });

  dom.badgeDiscCount.textContent = discrepancyItems.length;
  dom.discrepancyCountBadge.textContent = `${discrepancyItems.length} ta holat`;

  let totalSurplusItems = 0;
  let totalShortageItems = 0;

  discrepancyItems.forEach((it) => {
    const b = Number(it.bookQty) || 0;
    const a = Number(it.actualQty) || 0;
    if (a > b) totalSurplusItems += (a - b);
    if (b > a) totalShortageItems += (b - a);
  });

  dom.discSummarySurplus.textContent = `${totalSurplusItems} dona / birlik`;
  dom.discSummaryShortage.textContent = `${totalShortageItems} dona / birlik`;

  dom.discrepanciesTbody.innerHTML = "";

  if (discrepancyItems.length === 0) {
    dom.discrepanciesTbody.innerHTML = `
      <tr>
        <td colspan="13" style="text-align: center; padding: 40px; color: var(--success); font-weight: 700;">
          Inventarizatsiya bo'yicha hech qanday tafovut aniqlanmagan! Barcha tovarlar hisob ma'lumotlariga to'liq mos keladi.
        </td>
      </tr>
    `;
    return;
  }

  discrepancyItems.forEach((it) => {
    const b = Number(it.bookQty) || 0;
    const a = Number(it.actualQty) || 0;
    const price = Number(it.unitPrice) || 0;
    const surplus = Math.max(a - b, 0);
    const shortage = Math.max(b - a, 0);
    const diffVal = (a - b) * price;

    const extra = appState.discrepanciesExtras[it.id] || {
      explanation: it.remarks || "",
      correctiveAction: surplus > 0 ? "Ortiqcha kirim qilinsin" : "Servis markazidan qabul qilinib topshirilsin",
      responsiblePerson: it.responsiblePerson || ""
    };

    const tr = document.createElement("tr");
    tr.className = surplus > 0 ? "row-surplus" : "row-shortage";

    tr.innerHTML = `
      <td style="font-family: var(--font-mono); font-weight: 700; color: var(--navy-primary);">${it.code || ""}</td>
      <td class="col-desc"><strong>${it.description || ""}</strong></td>
      <td>${it.region || ""}</td>
      <td>${it.institution || ""}</td>
      <td class="col-qty">${b}</td>
      <td class="col-qty font-bold">${a}</td>
      <td class="col-qty">${surplus > 0 ? `<span class="badge badge-surplus">+${surplus}</span>` : '0'}</td>
      <td class="col-qty">${shortage > 0 ? `<span class="badge badge-shortage">-${shortage}</span>` : '0'}</td>
      <td class="col-currency ${diffVal < 0 ? 'text-danger font-bold' : 'text-success font-bold'}">
        ${diffVal > 0 ? '+' : ''}${formatNumber(diffVal)} ${it.currency}
      </td>
      <td><span class="badge ${getConditionBadgeClass(it.condition)}">${it.condition || ""}</span></td>
      <td>
        <input type="text" class="cell-input disc-explanation" data-id="${it.id}" value="${extra.explanation || ''}" placeholder="Tushuntirish xati / Izoh..." />
      </td>
      <td>
        <input type="text" class="cell-input disc-action" data-id="${it.id}" value="${extra.correctiveAction || ''}" placeholder="Ko'riladigan chora..." />
      </td>
      <td>
        <input type="text" class="cell-input disc-resp" data-id="${it.id}" value="${extra.responsiblePerson || it.responsiblePerson || ''}" placeholder="Mas'ul shaxs..." />
      </td>
    `;
    dom.discrepanciesTbody.appendChild(tr);
  });

  // Attach live auto-save on change
  dom.discrepanciesTbody.querySelectorAll(".disc-explanation").forEach((input) => {
    input.addEventListener("change", (e) => {
      const id = e.target.dataset.id;
      if (!appState.discrepanciesExtras[id]) appState.discrepanciesExtras[id] = {};
      appState.discrepanciesExtras[id].explanation = e.target.value;
      saveState();
      showToast("Tushuntirish saqlandi!");
    });
  });

  dom.discrepanciesTbody.querySelectorAll(".disc-action").forEach((input) => {
    input.addEventListener("change", (e) => {
      const id = e.target.dataset.id;
      if (!appState.discrepanciesExtras[id]) appState.discrepanciesExtras[id] = {};
      appState.discrepanciesExtras[id].correctiveAction = e.target.value;
      saveState();
      showToast("Chora choralari saqlandi!");
    });
  });

  dom.discrepanciesTbody.querySelectorAll(".disc-resp").forEach((input) => {
    input.addEventListener("change", (e) => {
      const id = e.target.dataset.id;
      if (!appState.discrepanciesExtras[id]) appState.discrepanciesExtras[id] = {};
      appState.discrepanciesExtras[id].responsiblePerson = e.target.value;
      saveState();
      showToast("Mas'ul shaxs yangilandi!");
    });
  });
}

// --------------------------------------------------------------------
// TAB 4: HANDOVER REGISTER
// --------------------------------------------------------------------
function renderHandoverTable() {
  dom.badgeHoCount.textContent = appState.handoverRecords.length;
  dom.handoverTbody.innerHTML = "";

  if (appState.handoverRecords.length === 0) {
    dom.handoverTbody.innerHTML = `
      <tr>
        <td colspan="15" style="text-align: center; padding: 40px; color: var(--slate-500);">
          Hozircha topshirish-qabul qilish aktlari mavjud emas. Yuqoridagi "+ Yangi topshirish akti" tugmasi orqali yangi akt tuzing.
        </td>
      </tr>
    `;
    return;
  }

  appState.handoverRecords.forEach((rec, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${rec.date || ""}</td>
      <td style="font-family: var(--font-mono); font-weight: 700; color: var(--navy-primary);">${rec.docNo || ""}</td>
      <td style="font-family: var(--font-mono);">${rec.itemCode || ""}</td>
      <td><strong>${rec.itemDescription || ""}</strong></td>
      <td class="col-qty font-bold">${rec.quantity || 0}</td>
      <td>${rec.unit || "dona"}</td>
      <td>${rec.region || ""}</td>
      <td>${rec.institution || ""}</td>
      <td>${rec.recipientName || ""}</td>
      <td>${rec.transferor || ""}</td>
      <td><span class="badge ${getConditionBadgeClass(rec.condition)}">${rec.condition || "Yaxshi"}</span></td>
      <td style="font-size: 0.8rem;">${rec.supportingDoc || "-"}</td>
      <td><span class="badge" style="background: #d1fae5; color: #065f46;">${rec.signatures || "Imzolangan"}</span></td>
      <td>${rec.remarks || "-"}</td>
      <td class="col-actions">
        <button class="btn btn-sm btn-outline btn-print-single-act" data-index="${idx}" title="Ushbu aktni ko'rish va chop etish">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="6 9 6 2 18 2 18 9"></polyline>
            <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path>
            <rect x="6" y="14" width="12" height="8"></rect>
          </svg>
          Chop etish
        </button>
      </td>
    `;
    dom.handoverTbody.appendChild(tr);
  });

  dom.handoverTbody.querySelectorAll(".btn-print-single-act").forEach((btn) => {
    btn.addEventListener("click", () => {
      const rec = appState.handoverRecords[btn.dataset.index];
      if (rec) {
        openPrintModal(`TOPSHIRISH-QABUL QILISH DALOLATNOMASI (${rec.docNo})`, generateSingleHandoverPrintHtml(rec));
      }
    });
  });
}

// --------------------------------------------------------------------
// TAB 5: COMMITTEE & PROTOCOL
// --------------------------------------------------------------------
function renderCommitteeTab() {
  const c = appState.committee;
  dom.commOrg.value = c.organization || appState.projectMetadata.organization || "";
  dom.commProject.value = c.projectName || appState.projectMetadata.projectName || "";
  dom.commOrderNo.value = c.orderNo || appState.projectMetadata.orderNo || "";
  dom.commStartDate.value = c.startDate || "";
  dom.commEndDate.value = c.endDate || appState.projectMetadata.inventoryDate || "";

  dom.commConclusion.value = c.conclusion || "";
  dom.commDiscrepancies.value = c.discrepanciesIdentified || "";
  dom.commRecommended.value = c.recommendedActions || "";
  dom.commDecision.value = c.managementDecision || "";

  dom.commMembersTbody.innerHTML = "";
  c.members.forEach((m, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td style="font-weight: 600;">${m.role}</td>
      <td>
        <input type="text" class="cell-input comm-member-name" data-idx="${idx}" value="${m.name || ''}" />
      </td>
      <td>
        <input type="text" class="cell-input comm-member-sig" data-idx="${idx}" value="${m.signature || ''}" placeholder="Imzo kodi / matni..." />
      </td>
      <td>
        <input type="date" class="cell-input comm-member-date" data-idx="${idx}" value="${m.date || ''}" />
      </td>
    `;
    dom.commMembersTbody.appendChild(tr);
  });

  // Listen to member input edits
  dom.commMembersTbody.querySelectorAll(".comm-member-name").forEach((inp) => {
    inp.addEventListener("change", (e) => {
      appState.committee.members[e.target.dataset.idx].name = e.target.value;
      saveState();
    });
  });
  dom.commMembersTbody.querySelectorAll(".comm-member-sig").forEach((inp) => {
    inp.addEventListener("change", (e) => {
      appState.committee.members[e.target.dataset.idx].signature = e.target.value;
      saveState();
    });
  });
  dom.commMembersTbody.querySelectorAll(".comm-member-date").forEach((inp) => {
    inp.addEventListener("change", (e) => {
      appState.committee.members[e.target.dataset.idx].date = e.target.value;
      saveState();
    });
  });
}

// ====================================================================
// PRINTABLE DOCUMENT GENERATORS (A4 FORMAL FORMATTING)
// ====================================================================

function generateInventoryPrintHtml() {
  const meta = appState.projectMetadata;
  const rowsHtml = appState.items
    .map(
      (it, idx) => `
    <tr>
      <td style="text-align: center;">${idx + 1}</td>
      <td>${it.code}</td>
      <td>${it.description}</td>
      <td style="text-align: center;">${it.unit}</td>
      <td>${it.region.split("/")[0]}</td>
      <td>${it.institution}</td>
      <td style="text-align: right;">${it.bookQty}</td>
      <td style="text-align: right; font-weight: bold;">${it.actualQty}</td>
      <td style="text-align: right;">${Math.max(it.actualQty - it.bookQty, 0)}</td>
      <td style="text-align: right;">${Math.max(it.bookQty - it.actualQty, 0)}</td>
      <td style="text-align: right;">${formatNumber(it.unitPrice)} ${it.currency}</td>
      <td style="text-align: right;">${formatNumber(it.actualQty * it.unitPrice)}</td>
      <td>${it.condition}</td>
    </tr>
  `
    )
    .join("");

  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>LOYIHA YAKUNIY INVENTARIZATSIYASI REYESTRI</h2>
        <h3>PROJECT CLOSURE INVENTORY REGISTER</h3>
      </div>
      <table class="doc-meta-table">
        <tr>
          <td><strong>Loyiha / Project:</strong> ${meta.projectName}</td>
          <td><strong>Inventarizatsiya sanasi / Date:</strong> ${meta.inventoryDate}</td>
        </tr>
        <tr>
          <td><strong>Buyruq raqami va sanasi / Order:</strong> ${meta.orderNo}</td>
          <td><strong>USD kursi / Rate:</strong> 1 USD = ${meta.usdRate} UZS</td>
        </tr>
      </table>

      <table class="doc-table">
        <thead>
          <tr>
            <th>№</th>
            <th>Tovar kodi</th>
            <th>Tovar nomi</th>
            <th>Birlik</th>
            <th>Hudud</th>
            <th>Tashkilot</th>
            <th>Hisob</th>
            <th>Haqiqiy</th>
            <th>Ortiqcha</th>
            <th>Kamomad</th>
            <th>Narxi</th>
            <th>Qiymati</th>
            <th>Holati</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml}
        </tbody>
      </table>

      <div class="signature-grid">
        <div class="sig-block">
          <strong>Inventarizatsiya komissiyasi raisi:</strong><br/>
          F.I.Sh: ${appState.committee.members[0]?.name || "____________________"}<br/>
          Imzo: _________________ Sana: ${meta.inventoryDate}
        </div>
        <div class="sig-block">
          <strong>Bosh buxgalter:</strong><br/>
          F.I.Sh: ${appState.committee.members[3]?.name || "____________________"}<br/>
          Imzo: _________________ Sana: ${meta.inventoryDate}
        </div>
      </div>
    </div>
  `;
}

function generateRegionalPrintHtml() {
  const meta = appState.projectMetadata;
  const usdRate = Number(meta.usdRate) || 12850;

  const rows = UZBEKISTAN_REGIONS.map((region, idx) => {
    const rItems = appState.items.filter((it) => it.region === region);
    const count = rItems.length;
    let bQty = 0, aQty = 0, surplus = 0, shortage = 0, uzsVal = 0, usdVal = 0;
    rItems.forEach((it) => {
      const b = Number(it.bookQty) || 0;
      const a = Number(it.actualQty) || 0;
      const p = Number(it.unitPrice) || 0;
      bQty += b;
      aQty += a;
      surplus += Math.max(a - b, 0);
      shortage += Math.max(b - a, 0);
      if (it.currency === "USD") usdVal += (a * p);
      else uzsVal += (a * p);
    });
    return `
      <tr>
        <td style="text-align: center;">${idx + 1}</td>
        <td><strong>${region}</strong></td>
        <td style="text-align: center;">${count}</td>
        <td style="text-align: right;">${formatNumber(bQty)}</td>
        <td style="text-align: right; font-weight: bold;">${formatNumber(aQty)}</td>
        <td style="text-align: right;">${surplus}</td>
        <td style="text-align: right;">${shortage}</td>
        <td style="text-align: right;">${formatNumber(uzsVal)}</td>
        <td style="text-align: right;">$${formatNumber(usdVal)}</td>
      </tr>
    `;
  }).join("");

  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>HUDUDLAR BO'YICHA JAMLANMA REYESTRI</h2>
        <h3>REGIONAL ASSET SUMMARY REPORT</h3>
      </div>
      <table class="doc-meta-table">
        <tr>
          <td><strong>Loyiha:</strong> ${meta.projectName}</td>
          <td><strong>Sana:</strong> ${meta.inventoryDate}</td>
        </tr>
      </table>

      <table class="doc-table">
        <thead>
          <tr>
            <th>№</th>
            <th>Hudud / Region</th>
            <th>Pozitsiyalar</th>
            <th>Hisob miqdori</th>
            <th>Haqiqiy miqdor</th>
            <th>Ortiqcha</th>
            <th>Kamomad</th>
            <th>UZS qiymati</th>
            <th>USD qiymati</th>
          </tr>
        </thead>
        <tbody>
          ${rows}
        </tbody>
      </table>

      <div class="signature-grid">
        <div class="sig-block">
          <strong>Loyiha rahbari:</strong> _____________________<br/>
          Imzo: ________________ Sana: ${meta.inventoryDate}
        </div>
        <div class="sig-block">
          <strong>Bosh buxgalter:</strong> _____________________<br/>
          Imzo: ________________ Sana: ${meta.inventoryDate}
        </div>
      </div>
    </div>
  `;
}

function generateSingleHandoverPrintHtml(rec) {
  const meta = appState.projectMetadata;
  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>ASOSIY VOSITALAR VA TOVAR-MODDIY BOYLIKLARNI TOPSHIRISH-QABUL QILISH DALOLATNOMASI</h2>
        <h3>CERTIFICATE OF HANDOVER AND ACCEPTANCE № ${rec.docNo}</h3>
      </div>

      <table class="doc-meta-table">
        <tr>
          <td><strong>Tuzilgan sana:</strong> ${rec.date}</td>
          <td><strong>Hudud:</strong> ${rec.region}</td>
        </tr>
        <tr>
          <td><strong>Loyiha nomi:</strong> ${meta.projectName}</td>
          <td><strong>Asos hujjati:</strong> ${rec.supportingDoc || meta.orderNo}</td>
        </tr>
      </table>

      <p style="margin: 14px 0; text-indent: 20px;">
        Bizlar, quyida imzo chekuvchilar — topshiruvchi <strong>${rec.transferor}</strong> hamda qabul qilib oluvchi 
        <strong>${rec.institution}</strong> vakili <strong>${rec.recipientName}</strong> ushbu dalolatnomani tuzdik shul haqidakim, 
        loyiha doirasida quyidagi tovar-moddiy boyliklar topshirildi va qabul qilindi:
      </p>

      <table class="doc-table">
        <thead>
          <tr>
            <th>№</th>
            <th>Tovar kodi</th>
            <th>Tovar nomi va modeli</th>
            <th>O'lchov birligi</th>
            <th>Miqdori</th>
            <th>Texnik holati</th>
            <th>Izoh</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="text-align: center;">1</td>
            <td><strong>${rec.itemCode}</strong></td>
            <td>${rec.itemDescription}</td>
            <td style="text-align: center;">${rec.unit}</td>
            <td style="text-align: center; font-weight: bold; font-size: 11pt;">${rec.quantity}</td>
            <td>${rec.condition}</td>
            <td>${rec.remarks || "To'liq komplektatsiyada"}</td>
          </tr>
        </tbody>
      </table>

      <p style="margin: 14px 0; font-size: 10pt;">
        Qabul qilib oluvchi tomon jihozlarning texnik holatini, to'liqligini va ishchi parametrlarini joyida ko'zdan kechirdi. 
        Topshirilayotgan mulk bo'yicha e'tirozlar mavjud emas.
      </p>

      <div class="signature-grid" style="margin-top: 40px;">
        <div class="sig-block">
          <strong>TOPSHIRDI:</strong><br/>
          (Topshiruvchi mas'ul shaxs)<br/><br/>
          F.I.Sh: <strong>${rec.transferor}</strong><br/>
          Imzo: _____________________<br/>
          M.O'. (Muhr o'rni)
        </div>
        <div class="sig-block">
          <strong>QABUL QILIB OLDI:</strong><br/>
          (${rec.institution} nomidan)<br/><br/>
          F.I.Sh: <strong>${rec.recipientName}</strong><br/>
          Imzo: _____________________<br/>
          M.O'. (Muhr o'rni)
        </div>
      </div>
    </div>
  `;
}

function generateCommitteePrintHtml() {
  const c = appState.committee;
  const meta = appState.projectMetadata;

  const membersList = c.members
    .map(
      (m) => `
    <tr>
      <td><strong>${m.role}</strong></td>
      <td>${m.name}</td>
      <td style="text-align: center;">${m.signature || "______________"}</td>
      <td style="text-align: center;">${m.date || meta.inventoryDate}</td>
    </tr>
  `
    )
    .join("");

  return `
    <div class="print-paper">
      <div class="doc-letterhead">
        <h2>INVENTARIZATSIYA KOMISSIYASI YAKUNIY BAYONNOMASI</h2>
        <h3>FINAL PROTOCOL OF THE INVENTORY COMMISSION</h3>
      </div>

      <table class="doc-meta-table">
        <tr>
          <td><strong>Tashkilot:</strong> ${c.organization}</td>
          <td><strong>Loyiha:</strong> ${c.projectName}</td>
        </tr>
        <tr>
          <td><strong>Buyruq:</strong> ${c.orderNo}</td>
          <td><strong>Davr:</strong> ${c.startDate} dan ${c.endDate} gacha</td>
        </tr>
      </table>

      <h4 style="margin: 14px 0 6px 0; font-size: 11pt; text-transform: uppercase;">1. Inventarizatsiya komissiyasi tarkibi:</h4>
      <table class="doc-table">
        <thead>
          <tr>
            <th>Lavozim</th>
            <th>F.I.Sh.</th>
            <th>Imzo</th>
            <th>Sana</th>
          </tr>
        </thead>
        <tbody>
          ${membersList}
        </tbody>
      </table>

      <h4 style="margin: 16px 0 6px 0; font-size: 11pt; text-transform: uppercase;">2. Komissiya xulosasi:</h4>
      <p style="text-align: justify; margin-bottom: 12px;">${c.conclusion}</p>

      <h4 style="margin: 16px 0 6px 0; font-size: 11pt; text-transform: uppercase;">3. Aniqlangan tafovutlar tahlili:</h4>
      <p style="text-align: justify; margin-bottom: 12px;">${c.discrepanciesIdentified}</p>

      <h4 style="margin: 16px 0 6px 0; font-size: 11pt; text-transform: uppercase;">4. Tavsiya etilgan choralar:</h4>
      <p style="text-align: justify; margin-bottom: 12px; white-space: pre-line;">${c.recommendedActions}</p>

      <h4 style="margin: 16px 0 6px 0; font-size: 11pt; text-transform: uppercase;">5. Rahbar qarori:</h4>
      <p style="text-align: justify; margin-bottom: 24px;">${c.managementDecision}</p>

      <div class="signature-grid">
        <div class="sig-block">
          <strong>Komissiya raisi:</strong><br/>
          ${c.members[0]?.name || "___________________"}<br/>
          Imzo: _________________ Sana: ${c.endDate}
        </div>
        <div class="sig-block">
          <strong>Bosh buxgalter:</strong><br/>
          ${c.members[3]?.name || "___________________"}<br/>
          Imzo: _________________ Sana: ${c.endDate}
        </div>
      </div>
    </div>
  `;
}

// Kickoff
document.addEventListener("DOMContentLoaded", init);
