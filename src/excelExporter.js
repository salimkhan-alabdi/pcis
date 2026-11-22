import * as XLSX from "xlsx";

/**
 * Exports current PEMS-UZ application state to an official multi-sheet Excel file (.xlsx)
 * covering all modules defined in the Technical Specification (ТЗ).
 */
export function exportToExcel(state) {
  const wb = XLSX.utils.book_new();
  const meta = state.projectMetadata;

  // 1. DASHBOARD & SUMMARY SHEET
  const summaryRows = [
    ["PEMS-UZ: PROJECT EQUIPMENT MANAGEMENT SYSTEM (UZBEKISTAN)"],
    ["Система учёта, распределения, поступления, передачи, мониторинга и инвентаризации оборудования (АБР / АБИИ)"],
    ["Разработчик / Developed by:", meta.developer || "Abdiyev Salimkhon Sanjar ugli"],
    ["Loyiha / Project:", meta.projectName],
    ["Asos hujjati / Basis:", meta.presidentialResolution || "PQ-3818 (16.10.2023)"],
    ["TEO hujjati / Feasibility Study:", meta.teoNumber || "TEO-ADB/AIIB-2023/V4"],
    ["Buyruq / Order:", meta.orderNo],
    ["Sana / Date:", meta.inventoryDate],
    ["Valyuta kursi / Exchange Rate:", `1 USD = ${meta.usdRate || 12850} UZS`],
    [],
    ["ASOSIY KO'RSATKICHLAR / KEY PROJECT INDICATORS (PIPELINE RECONCILIATION)"],
    ["№", "Ko'rsatkich nomi / Indicator", "Miqdori / Quantity", "Birlik / Unit", "Holati / Status"]
  ];

  const teoTotal = (state.teoItems || []).reduce((acc, i) => acc + (Number(i.approvedQty) || 0), 0);
  const contractedTotal = (state.contracts || []).reduce((acc, c) => acc + (Number(c.quantity) || 0), 0);
  const deliveredTotal = (state.receipts || []).reduce((acc, r) => acc + (Number(r.deliveredQuantity) || 0), 0);
  const transferredTotal = (state.handovers || []).reduce((acc, h) => acc + (Number(h.quantity) || 0), 0);
  const installedTotal = (state.installations || []).filter(i => (i.status || "").includes("O‘rnatilgan") || (i.status || "").includes("Foydalanishga")).length;
  const commissionedTotal = (state.installations || []).filter(i => (i.status || "").includes("Foydalanishga topshirilgan")).length;
  const monitoredTotal = (state.monitoringRecords || []).length;
  const defectsTotal = (state.monitoringRecords || []).filter(m => (m.technicalCondition || "").includes("Nuqsonli") || (m.technicalCondition || "").includes("Ta'mir") || (m.technicalCondition || "").includes("Ishlamaydi")).length;
  const inventoriedTotal = (state.inventoryItems || state.items || []).reduce((acc, i) => acc + (Number(i.actualQty) || 0), 0);
  const remainingTotal = Math.max(teoTotal - transferredTotal, 0);

  const kpis = [
    ["1", "TEO bo'yicha tasdiqlangan uskunalar / Approved in Feasibility Study", teoTotal, "komplekt", "Reja asosida"],
    ["2", "Shartnomalar bo'yicha xarid qilingan / Contracted Equipment", contractedTotal, "komplekt", "Shartnomalar imzolangan"],
    ["3", "Fakt bo'yicha yetkazib berilgan (Ombor/Bojxona) / Actual Delivered", deliveredTotal, "komplekt", "УзАСБОда кирим қилинган"],
    ["4", "Yakuniy tashkilotlarga topshirilgan / Transferred to End-Users", transferredTotal, "komplekt", "Aktlar imzolangan"],
    ["5", "O'rnatilgan va montaj qilingan / Installed Equipment", installedTotal, "dona", "Ishchi holatda"],
    ["6", "Foydalanishga to'liq topshirilgan / Commissioned", commissionedTotal, "dona", "Ekspluatatsiyada"],
    ["7", "GRP tomonidan monitoring qilingan / Monitored by PIU", monitoredTotal, "birlik", "Dala tekshiruvi o'tkazilgan"],
    ["8", "Aniqlangan nuqsonli / ta'mir talab uskunalar / Defective / Repair", defectsTotal, "birlik", "Kafolat talabnomasi yuborilgan"],
    ["9", "Inventarizatsiya qilingan uskunalar soni / Inventoried", inventoriedTotal, "birlik", "Auditdan o'tkazildi"],
    ["10", "Qoldiq (Taqsimlanmagan yoki zaxirada) / Remaining Balance", remainingTotal, "komplekt", "Zaxirada mavjud"]
  ];

  kpis.forEach(k => summaryRows.push(k));
  const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
  XLSX.utils.book_append_sheet(wb, wsSummary, "Dashboard_Summary");

  // 2. MODULE 1: TEO DISTRIBUTION SHEET
  const teoHeaders = [
    "№", "Prezident qarori / Presidential Resolution", "TEO hujjati / Feasibility Study",
    "Uskuna kodi / Item code", "Uskuna nomi / Description", "Tasdiqlangan miqdor / Approved Qty",
    "Hudud / Region", "Qabul qiluvchi tashkilot / Receiving Org", "Rejalashtirilgan / Planned",
    "Yetkazilgan / Delivered", "Topshirilgan / Transferred", "Qoldiq / Remaining",
    "Narxi (USD) / Unit Price", "Holati / Status"
  ];
  const teoRows = [
    ["PEMS-UZ: 1. TEO TAQSIMOTI VA REJALASHTIRISH / FEASIBILITY STUDY ALLOCATION"],
    ["Loyiha:", meta.projectName],
    teoHeaders
  ];
  (state.teoItems || []).forEach((item, idx) => {
    teoRows.push([
      idx + 1,
      item.presidentialDoc || "",
      item.teoDoc || "",
      item.itemCode || "",
      item.itemName || "",
      item.approvedQty || 0,
      item.region || "",
      item.receivingOrg || "",
      item.plannedQty || 0,
      item.actualDeliveredQty || 0,
      item.actualTransferredQty || 0,
      item.remainingQty || 0,
      item.unitPriceUsd || 0,
      item.status || ""
    ]);
  });
  const wsTeo = XLSX.utils.aoa_to_sheet(teoRows);
  XLSX.utils.book_append_sheet(wb, wsTeo, "1_TEO_Distribution");

  // 3. MODULE 2: CONTRACTS & PROCUREMENT SHEET
  const cntHeaders = [
    "№", "Shartnoma № / Contract No", "Sana / Date", "Yetkazib beruvchi / Supplier",
    "Tovar nomi / Description", "Bir komplekt tarkibi / Components", "Ishlab chiqarilgan yili / Year",
    "Seriya raqamlari / Serial Range", "Ishlab chiqaruvchi / Manufacturer", "Davlat / Origin",
    "Model", "Birlik / Unit", "Miqdor / Qty", "Birlik narxi (USD)", "Jami tovar qiymati (USD)",
    "Birlik narxi (UZS)", "Jami tovar qiymati (UZS)", "Qo'shimcha xizmatlar (USD)",
    "Logistika (USD)", "Havo transporti (USD)", "Sug'urta (USD)", "Qadoqlash (USD)",
    "Umumiy shartnoma summasi (USD)", "Kafolat muddati / Warranty", "Holati / Status"
  ];
  const cntRows = [
    ["PEMS-UZ: 2. XARIDLAR VA SHARTNOMALAR / PROCUREMENT & CONTRACTS"],
    ["Loyiha:", meta.projectName],
    cntHeaders
  ];
  (state.contracts || []).forEach((c, idx) => {
    cntRows.push([
      idx + 1,
      c.contractNo || "",
      c.contractDate || "",
      c.supplier || "",
      c.goodsDescription || "",
      c.componentsIncluded || "",
      c.yearOfManufacture || "",
      c.serialNumberRange || "",
      c.manufacturer || "",
      c.countryOfOrigin || "",
      c.model || "",
      c.unit || "komplekt",
      c.quantity || 0,
      c.unitPriceUsd || 0,
      c.totalGoodsValueUsd || 0,
      c.unitPriceUzs || 0,
      c.totalGoodsValueUzs || 0,
      c.ancillaryServicesCostUsd || 0,
      c.domesticLogisticsCostUsd || 0,
      c.airFreightCostUsd || 0,
      c.insuranceCostUsd || 0,
      c.packagingCostUsd || 0,
      c.totalContractAmountUsd || 0,
      c.warrantyPeriod || "",
      c.status || ""
    ]);
  });
  const wsContracts = XLSX.utils.aoa_to_sheet(cntRows);
  XLSX.utils.book_append_sheet(wb, wsContracts, "2_Contracts_Procurement");

  // 4. MODULE 3: RECEIPTS & UZASBO SHEET
  const recHeaders = [
    "№", "Shartnoma №", "Hisob-faktura № / Invoice No", "Hisob-faktura sanasi",
    "Kelib tushgan sana / Receipt Date", "Bojxona deklaratsiyasi № (ГТД / БЮД)",
    "Bojxona sanasi", "UzASBO ro'yxatga olish sanasi", "Transfer №", "Transfer sanasi",
    "Yetkazilgan miqdor", "Tovar tavsifi", "UzASBO schyoti / Account", "UzASBO qiymati (UZS)",
    "Izoh / Remarks", "Holati / Status"
  ];
  const recRows = [
    ["PEMS-UZ: 3. USKUNALAR KELIB TUSHISHI VA UZASBO HISOBI / RECEIPTS & UZASBO"],
    ["Loyiha:", meta.projectName],
    recHeaders
  ];
  (state.receipts || []).forEach((r, idx) => {
    recRows.push([
      idx + 1,
      r.contractNo || "",
      r.invoiceNo || "",
      r.invoiceDate || "",
      r.receiptDate || "",
      r.customsDeclarationNo || "",
      r.customsClearanceDate || "",
      r.uzasboRegistrationDate || "",
      r.transferNo || "",
      r.transferDate || "",
      r.deliveredQuantity || 0,
      r.goodsDescription || "",
      r.uzasboAccountCode || "",
      r.uzasboBookValueUzs || 0,
      r.remarks || "",
      r.status || ""
    ]);
  });
  const wsReceipts = XLSX.utils.aoa_to_sheet(recRows);
  XLSX.utils.book_append_sheet(wb, wsReceipts, "3_Receipts_UzASBO");

  // 5. MODULE 4: HANDOVERS & TRANSFERS SHEET
  const hoHeaders = [
    "№", "Transfer №", "Transfer sanasi", "Dalolatnoma № / Handover Act No", "Dalolatnoma sanasi",
    "Viloyat / Region", "Tuman (shahar) / District", "Qabul qiluvchi tashkilot / Receiving Org",
    "STIR / TIN", "Bo'linma / Laboratoriya", "Manzili / Address", "Uskuna kodi",
    "Uskuna nomi", "Seriya raqami", "Inventar raqami", "Miqdori", "Birligi",
    "Umumiy qiymat (USD)", "Umumiy qiymat (UZS)", "Moddiy javobgar shaxs (F.I.Sh.)",
    "Topshiruvchi", "Qabul qiluvchi", "Holati", "Asos hujjati", "Izoh"
  ];
  const hoRows = [
    ["PEMS-UZ: 4. USKUNALARNI TOPSHIRISH VA TRANSFER REYESTRI / HANDOVER & TRANSFER"],
    ["Loyiha:", meta.projectName],
    hoHeaders
  ];
  (state.handovers || []).forEach((h, idx) => {
    hoRows.push([
      idx + 1,
      h.transferNo || "",
      h.transferDate || "",
      h.docNo || "",
      h.date || "",
      h.region || "",
      h.district || "",
      h.institution || "",
      h.tin || "",
      h.department || "",
      h.address || "",
      h.itemCode || "",
      h.itemDescription || "",
      h.serialNumber || "",
      h.inventoryNumber || "",
      h.quantity || 0,
      h.unit || "komplekt",
      h.totalValueUsd || 0,
      h.totalValueUzs || 0,
      h.responsiblePerson || "",
      h.transferor || "",
      h.recipientName || "",
      h.condition || "",
      h.supportingDoc || "",
      h.remarks || ""
    ]);
  });
  const wsHandovers = XLSX.utils.aoa_to_sheet(hoRows);
  XLSX.utils.book_append_sheet(wb, wsHandovers, "4_Handovers_Transfers");

  // 6. MODULE 5: INSTALLATION & COMMISSIONING SHEET
  const instHeaders = [
    "№", "Dalolatnoma №", "Uskuna kodi", "Uskuna nomi", "Seriya raqami", "Tashkilot",
    "Hudud", "O'rnatish holati / Status", "O'rnatish sanasi", "Montaj tugatilgan sana",
    "Sinov sanasi", "Sinov natijasi / Test Results", "Foydalanishga topshirish sanasi",
    "Ekspluatatsiya dalolatnoma №", "Dalolatnoma sanasi", "Yetkazib beruvchi mutaxassisi",
    "O'qitish o'tkazilganligi / Training", "Izoh / Remarks"
  ];
  const instRows = [
    ["PEMS-UZ: 5. O'RNATISH VA FOYDALANISHGA TOPSHIRISH / INSTALLATION & COMMISSIONING"],
    ["Loyiha:", meta.projectName],
    instHeaders
  ];
  (state.installations || []).forEach((i, idx) => {
    instRows.push([
      idx + 1,
      i.handoverDocNo || "",
      i.itemCode || "",
      i.equipmentName || "",
      i.serialNumber || "",
      i.organization || "",
      i.region || "",
      i.status || "",
      i.installationDate || "",
      i.completionDate || "",
      i.testingDate || "",
      i.testingResults || "",
      i.commissioningDate || "",
      i.commissioningCertNo || "",
      i.commissioningCertDate || "",
      i.supplierSpecialist || "",
      i.trainingCompleted || "",
      i.remarks || ""
    ]);
  });
  const wsInst = XLSX.utils.aoa_to_sheet(instRows);
  XLSX.utils.book_append_sheet(wb, wsInst, "5_Installation_Commissioning");

  // 7. MODULE 6: PIU MONITORING SHEET
  const monHeaders = [
    "№", "Monitoring sanasi", "Tashkilot nomi", "Hudud", "Uskuna nomi",
    "Seriya raqami", "Inventar raqami", "Haqiqiy joylashgan joyi", "Haqiqiy mavjudligi",
    "Ekspluatatsiya holati", "Texnik holati (8 status)", "Aniqlangan nuqsonlar / Specific Defects",
    "Foto-ilova", "Aniqlangan kamchiliklar", "Tavsiyalar va muddatlar",
    "Monitoring natijasi", "Monitoring o'tkazgan mas'ul shaxs"
  ];
  const monRows = [
    ["PEMS-UZ: 6. GRP TOMONIDAN USKUNALAR MONITORINGI / PIU MONITORING"],
    ["Loyiha:", meta.projectName],
    monHeaders
  ];
  (state.monitoringRecords || []).forEach((m, idx) => {
    monRows.push([
      idx + 1,
      m.monitoringDate || "",
      m.monitoredOrg || "",
      m.region || "",
      m.equipmentName || "",
      m.serialNumber || "",
      m.inventoryNumber || "",
      m.actualLocation || "",
      m.physicalAvailability || "",
      m.commissioningStatus || "",
      m.technicalCondition || "",
      m.identifiedDefects || "",
      m.photoEvidence || "",
      m.identifiedIssues || "",
      m.recommendations || "",
      m.monitoringResult || "",
      m.monitoringOfficer || ""
    ]);
  });
  const wsMon = XLSX.utils.aoa_to_sheet(monRows);
  XLSX.utils.book_append_sheet(wb, wsMon, "6_PIU_Monitoring");

  // 8. MODULE 7: INVENTORY AUDIT & DISCREPANCIES SHEET
  const invHeaders = [
    "№", "Uskuna kodi", "Uskuna nomi", "Seriya raqami", "Inventar raqami",
    "O'lchov birligi", "Hudud", "Tashkilot", "Saqlash joyi / Joylashuvi",
    "Mas'ul shaxs", "Hisob miqdori (Kitob bo'yicha)", "Haqiqiy miqdor (Mavjud)",
    "Ortiqcha", "Kamomad", "Birlik narxi", "Valyuta", "USD kursi",
    "Hisob qiymati", "Haqiqiy qiymat", "Farq (Balans)", "Texnik holati",
    "Topshirish hujjati (Akt №)", "Transfer №", "Izoh"
  ];
  const invRows = [
    ["PEMS-UZ: 7. ASOSIY VOSITALAR INVENTARIZATSIYA REYESTRI / INVENTORY AUDIT"],
    ["Loyiha / Project:", meta.projectName],
    ["Buyruq / Order:", meta.orderNo],
    ["Inventarizatsiya sanasi:", meta.inventoryDate],
    [
      "Eslatma: Система учёта в рамках проекта АБР / АБИИ. Dev: " +
        (meta.developer || "Abdiyev Salimkhon Sanjar ugli") +
        " | 1 USD = " +
        (meta.usdRate || 12850) +
        " UZS"
    ],
    invHeaders
  ];

  const items = state.inventoryItems || state.items || [];
  items.forEach((item, idx) => {
    const bookQty = Number(item.bookQty) || 0;
    const actualQty = Number(item.actualQty) || 0;
    const unitPrice = Number(item.unitPrice) || 0;
    const surplus = Math.max(actualQty - bookQty, 0);
    const shortage = Math.max(bookQty - actualQty, 0);
    const bookVal = bookQty * unitPrice;
    const actualVal = actualQty * unitPrice;
    const diff = actualVal - bookVal;

    invRows.push([
      idx + 1,
      item.code || "",
      item.description || "",
      item.serialNumber || "",
      item.inventoryNumber || "",
      item.unit || "komplekt",
      item.region || "",
      item.institution || "",
      item.location || "",
      item.responsiblePerson || "",
      bookQty,
      actualQty,
      surplus,
      shortage,
      unitPrice,
      item.currency || "USD",
      Number(meta.usdRate) || 12850,
      bookVal,
      actualVal,
      diff,
      item.condition || "Ishga yaroqli",
      item.handoverDoc || "",
      item.transferNo || "",
      item.remarks || ""
    ]);
  });
  const wsInv = XLSX.utils.aoa_to_sheet(invRows);
  XLSX.utils.book_append_sheet(wb, wsInv, "7_Inventory_Audit");

  // 9. AUDIT LOG SHEET
  const logHeaders = ["№", "Vaqti / Timestamp", "Foydalanuvchi / User", "Modul / Module", "Amal / Action", "Obyekt / Entity", "Batafsil / Details", "Eski qiymat / Old", "Yangi qiymat / New"];
  const logRows = [
    ["PEMS-UZ: 9. TIZIM OPERATSIYALARI JURNALI / SYSTEM AUDIT TRAIL"],
    ["Loyiha:", meta.projectName],
    logHeaders
  ];
  (state.auditLogs || []).forEach((l, idx) => {
    logRows.push([
      idx + 1,
      l.timestamp || "",
      l.userName || "",
      l.module || "",
      l.action || "",
      l.entityId || "",
      l.details || "",
      l.oldValue || "",
      l.newValue || ""
    ]);
  });
  const wsLogs = XLSX.utils.aoa_to_sheet(logRows);
  XLSX.utils.book_append_sheet(wb, wsLogs, "8_System_Audit_Log");

  // Auto-trigger browser download
  const filename = "PEMS_UZ_Equipment_Registry_ADB_AIIB.xlsx";
  XLSX.writeFile(wb, filename);
}

/**
 * Parses an uploaded Excel (.xlsx) file into application state
 */
export async function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });

        const parsedItems = [];
        const invSheet = wb.Sheets["7_Inventory_Audit"] || wb.Sheets["Inventory"];
        if (invSheet) {
          const rows = XLSX.utils.sheet_to_json(invSheet, { header: 1 });
          const headerIdx = rows.findIndex(
            (r) =>
              r &&
              (r[0] === "№ / No." ||
                r[0] === "№" ||
                String(r[1] || "").includes("Tovar kodi") ||
                String(r[1] || "").includes("Uskuna kodi") ||
                String(r[1] || "").includes("Item code"))
          );
          const startRow = headerIdx >= 0 ? headerIdx + 1 : 6;
          for (let r = startRow; r < rows.length; r++) {
            const row = rows[r];
            if (!row || !row[1]) continue;
            parsedItems.push({
              id: "item-" + Date.now() + "-" + r,
              code: String(row[1] || "").trim(),
              description: String(row[2] || "").trim(),
              serialNumber: String(row[3] || "").trim(),
              inventoryNumber: String(row[4] || "").trim(),
              unit: String(row[5] || "komplekt").trim(),
              region: String(row[6] || "").trim(),
              institution: String(row[7] || "").trim(),
              location: String(row[8] || "").trim(),
              responsiblePerson: String(row[9] || "").trim(),
              bookQty: Number(row[10]) || 0,
              actualQty: Number(row[11]) || 0,
              unitPrice: Number(row[14]) || 0,
              currency: String(row[15] || "USD").trim().toUpperCase(),
              condition: String(row[20] || "Ishga yaroqli").trim(),
              handoverDoc: String(row[21] || "").trim(),
              transferNo: String(row[22] || "").trim(),
              remarks: String(row[23] || "").trim()
            });
          }
        }

        resolve({
          items: parsedItems
        });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}
