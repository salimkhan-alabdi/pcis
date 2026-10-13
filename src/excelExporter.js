import * as XLSX from "xlsx";
import { UZBEKISTAN_REGIONS } from "./data.js";

/**
 * Exports current inventory state to an authentic multi-sheet Excel file (.xlsx)
 * matching the openpyxl script structure and formatting.
 */
export function exportToExcel(state) {
  const wb = XLSX.utils.book_new();

  // 1. INVENTORY SHEET
  const invHeaders = [
    "№ / No.",
    "Tovar kodi / Item code",
    "Tovar nomi / Item description",
    "O'lchov birligi / Unit",
    "Hudud / Region",
    "Qabul qiluvchi tashkilot / Receiving institution",
    "Saqlash joyi / Location",
    "Mas'ul shaxs / Responsible person",
    "Hisob miqdori / Book quantity",
    "Haqiqiy miqdor / Actual quantity",
    "Ortiqcha / Surplus",
    "Kamomad / Shortage",
    "Birlik narxi / Unit price",
    "Valyuta / Currency",
    "USD kursi / USD exchange rate",
    "Hisob qiymati / Book value",
    "Haqiqiy qiymat / Actual value",
    "Farq / Difference",
    "Holati / Condition",
    "Topshirish hujjati / Handover document",
    "Izoh / Remarks"
  ];

  const invRows = [
    ["LOYIHA YAKUNIY INVENTARIZATSIYASI / PROJECT CLOSURE INVENTORY"],
    ["Loyiha / Project:", state.projectMetadata.projectName],
    ["Buyruq raqami va sanasi / Order No. and date:", state.projectMetadata.orderNo],
    ["Inventarizatsiya sanasi / Inventory date:", state.projectMetadata.inventoryDate],
    ["Eslatma / Note: USD kursini tegishli sana uchun kiriting (" + state.projectMetadata.usdRate + ")."],
    invHeaders
  ];

  state.items.forEach((item, idx) => {
    const bookQty = Number(item.bookQty) || 0;
    const actualQty = Number(item.actualQty) || 0;
    const unitPrice = Number(item.unitPrice) || 0;
    const surplus = Math.max(actualQty - bookQty, 0);
    const shortage = Math.max(bookQty - actualQty, 0);
    const bookValue = bookQty * unitPrice;
    const actualValue = actualQty * unitPrice;
    const diff = actualValue - bookValue;

    invRows.push([
      idx + 1,
      item.code || "",
      item.description || "",
      item.unit || "dona",
      item.region || "",
      item.institution || "",
      item.location || "",
      item.responsiblePerson || "",
      bookQty,
      actualQty,
      surplus,
      shortage,
      unitPrice,
      item.currency || "UZS",
      Number(state.projectMetadata.usdRate) || 12850,
      bookValue,
      actualValue,
      diff,
      item.condition || "Yaxshi / Good",
      item.handoverDoc || "",
      item.remarks || ""
    ]);
  });

  const wsInv = XLSX.utils.aoa_to_sheet(invRows);

  // Set column widths for Inventory sheet
  wsInv["!cols"] = [
    { wch: 8 },  // A
    { wch: 18 }, // B
    { wch: 36 }, // C
    { wch: 12 }, // D
    { wch: 32 }, // E
    { wch: 32 }, // F
    { wch: 24 }, // G
    { wch: 26 }, // H
    { wch: 16 }, // I
    { wch: 16 }, // J
    { wch: 14 }, // K
    { wch: 14 }, // L
    { wch: 18 }, // M
    { wch: 12 }, // N
    { wch: 16 }, // O
    { wch: 20 }, // P
    { wch: 20 }, // Q
    { wch: 18 }, // R
    { wch: 20 }, // S
    { wch: 24 }, // T
    { wch: 30 }  // U
  ];

  XLSX.utils.book_append_sheet(wb, wsInv, "Inventory");

  // 2. REGIONAL SUMMARY SHEET
  const regHeaders = [
    "Hudud / Region",
    "Pozitsiyalar / Items",
    "Hisob miqdori / Book qty",
    "Haqiqiy miqdor / Actual qty",
    "Ortiqcha / Surplus",
    "Kamomad / Shortage",
    "UZS qiymati / UZS value",
    "USD qiymati / USD value"
  ];

  const regRows = [
    ["Hududlar bo'yicha jamlanma / Regional Summary"],
    regHeaders
  ];

  UZBEKISTAN_REGIONS.forEach((region) => {
    const regionItems = state.items.filter((it) => it.region === region);
    const count = regionItems.length;
    let bQty = 0;
    let aQty = 0;
    let surplus = 0;
    let shortage = 0;
    let uzsVal = 0;
    let usdVal = 0;

    regionItems.forEach((it) => {
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

    regRows.push([
      region,
      count,
      bQty,
      aQty,
      surplus,
      shortage,
      uzsVal,
      usdVal
    ]);
  });

  const wsReg = XLSX.utils.aoa_to_sheet(regRows);
  wsReg["!cols"] = [
    { wch: 48 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 16 },
    { wch: 16 },
    { wch: 22 },
    { wch: 20 }
  ];
  XLSX.utils.book_append_sheet(wb, wsReg, "Regional Summary");

  // 3. DISCREPANCIES SHEET
  const discHeaders = [
    "Item code",
    "Item description",
    "Region",
    "Institution",
    "Book qty",
    "Actual qty",
    "Surplus",
    "Shortage",
    "Difference value",
    "Condition",
    "Explanation / Izoh",
    "Corrective action / Chora",
    "Responsible person"
  ];

  const discRows = [
    ["Tafovutlar qaydnomasi / Discrepancy Register"],
    discHeaders
  ];

  state.items.forEach((item) => {
    const b = Number(item.bookQty) || 0;
    const a = Number(item.actualQty) || 0;
    const surplus = Math.max(a - b, 0);
    const shortage = Math.max(b - a, 0);

    if (surplus > 0 || shortage > 0) {
      const price = Number(item.unitPrice) || 0;
      const diffVal = (a - b) * price;
      const extra = state.discrepanciesExtras[item.id] || {};

      discRows.push([
        item.code || "",
        item.description || "",
        item.region || "",
        item.institution || "",
        b,
        a,
        surplus,
        shortage,
        `${diffVal > 0 ? "+" : ""}${diffVal.toLocaleString()} ${item.currency}`,
        item.condition || "",
        extra.explanation || item.remarks || "",
        extra.correctiveAction || (surplus > 0 ? "Kirim qilish" : "Hisobdan chiqarish / Servisga yuborish"),
        extra.responsiblePerson || item.responsiblePerson || ""
      ]);
    }
  });

  const wsDisc = XLSX.utils.aoa_to_sheet(discRows);
  wsDisc["!cols"] = [
    { wch: 18 },
    { wch: 36 },
    { wch: 30 },
    { wch: 30 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
    { wch: 22 },
    { wch: 20 },
    { wch: 36 },
    { wch: 36 },
    { wch: 26 }
  ];
  XLSX.utils.book_append_sheet(wb, wsDisc, "Discrepancies");

  // 4. HANDOVER REGISTER SHEET
  const hoHeaders = [
    "Date / Sana",
    "Document No. / Hujjat №",
    "Item code / Tovar kodi",
    "Item description / Tovar nomi",
    "Quantity / Miqdor",
    "Unit / Birlik",
    "Region / Hudud",
    "Receiving institution / Qabul qiluvchi",
    "Recipient name / Qabul qiluvchi vakili",
    "Transferor / Topshiruvchi",
    "Condition / Holati",
    "Supporting document / Asos hujjati",
    "Signatures / Imzolar",
    "Remarks / Izoh"
  ];

  const hoRows = [
    ["Topshirish-qabul qilish reyestri / Handover Register"],
    hoHeaders
  ];

  state.handoverRecords.forEach((rec) => {
    hoRows.push([
      rec.date || "",
      rec.docNo || "",
      rec.itemCode || "",
      rec.itemDescription || "",
      rec.quantity || 0,
      rec.unit || "dona",
      rec.region || "",
      rec.institution || "",
      rec.recipientName || "",
      rec.transferor || "",
      rec.condition || "",
      rec.supportingDoc || "",
      rec.signatures || "Tasdiqlangan",
      rec.remarks || ""
    ]);
  });

  const wsHo = XLSX.utils.aoa_to_sheet(hoRows);
  wsHo["!cols"] = [
    { wch: 16 },
    { wch: 18 },
    { wch: 18 },
    { wch: 34 },
    { wch: 14 },
    { wch: 12 },
    { wch: 28 },
    { wch: 32 },
    { wch: 26 },
    { wch: 26 },
    { wch: 20 },
    { wch: 28 },
    { wch: 20 },
    { wch: 30 }
  ];
  XLSX.utils.book_append_sheet(wb, wsHo, "Handover");

  // 5. COMMITTEE SHEET
  const commRows = [
    ["Inventarizatsiya komissiyasi / Inventory Committee"],
    ["Loyiha / Project", state.committee.projectName || state.projectMetadata.projectName],
    ["Tashkilot / Organization", state.committee.organization || state.projectMetadata.organization],
    ["Buyruq raqami / Order No.", state.committee.orderNo || state.projectMetadata.orderNo],
    ["Boshlanish sanasi / Start date", state.committee.startDate || ""],
    ["Tugash sanasi / End date", state.committee.endDate || state.projectMetadata.inventoryDate],
    [],
    ["Lavozim / Role", "F.I.Sh. / Full name", "Imzo / Signature", "Sana / Date"]
  ];

  state.committee.members.forEach((m) => {
    commRows.push([m.role, m.name, m.signature || "", m.date || ""]);
  });

  commRows.push([]);
  commRows.push(["Xulosa / Conclusion", state.committee.conclusion || ""]);
  commRows.push(["Aniqlangan tafovutlar / Discrepancies identified", state.committee.discrepanciesIdentified || ""]);
  commRows.push(["Tavsiya etilgan choralar / Recommended actions", state.committee.recommendedActions || ""]);
  commRows.push(["Rahbar qarori / Management decision", state.committee.managementDecision || ""]);

  const wsComm = XLSX.utils.aoa_to_sheet(commRows);
  wsComm["!cols"] = [
    { wch: 42 },
    { wch: 34 },
    { wch: 22 },
    { wch: 18 }
  ];
  XLSX.utils.book_append_sheet(wb, wsComm, "Committee");

  // Trigger browser download
  const filename = "Project_Inventory_UZ_EN.xlsx";
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
        const invSheet = wb.Sheets["Inventory"];
        if (invSheet) {
          const rows = XLSX.utils.sheet_to_json(invSheet, { header: 1 });
          // Header row is index 5 (row 6)
          // Data starts index 6 (row 7)
          for (let r = 6; r < rows.length; r++) {
            const row = rows[r];
            if (!row || !row[1]) continue; // item code
            parsedItems.push({
              id: "item-" + Date.now() + "-" + r,
              code: String(row[1] || "").trim(),
              description: String(row[2] || "").trim(),
              unit: String(row[3] || "dona").trim(),
              region: String(row[4] || "").trim(),
              institution: String(row[5] || "").trim(),
              location: String(row[6] || "").trim(),
              responsiblePerson: String(row[7] || "").trim(),
              bookQty: Number(row[8]) || 0,
              actualQty: Number(row[9]) || 0,
              unitPrice: Number(row[12]) || 0,
              currency: String(row[13] || "UZS").trim().toUpperCase(),
              condition: String(row[18] || "Yaxshi / Good").trim(),
              handoverDoc: String(row[19] || "").trim(),
              remarks: String(row[20] || "").trim()
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
