#!/usr/bin/env python3
"""
LOYIHA YAKUNIY INVENTARIZATSIYASI / PROJECT CLOSURE INVENTORY
Generates Project_Inventory_UZ_EN.xlsx matching openpyxl specifications.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import sys

FILE = "Project_Inventory_UZ_EN.xlsx"

def generate_inventory_workbook():
    wb = Workbook()

    headers = [
        "№ / No.", "Tovar kodi / Item code",
        "Tovar nomi / Item description",
        "O'lchov birligi / Unit", "Hudud / Region",
        "Qabul qiluvchi tashkilot / Receiving institution",
        "Saqlash joyi / Location",
        "Mas'ul shaxs / Responsible person",
        "Hisob miqdori / Book quantity",
        "Haqiqiy miqdor / Actual quantity",
        "Ortiqcha / Surplus", "Kamomad / Shortage",
        "Birlik narxi / Unit price", "Valyuta / Currency",
        "USD kursi / USD exchange rate",
        "Hisob qiymati / Book value",
        "Haqiqiy qiymat / Actual value",
        "Farq / Difference",
        "Holati / Condition",
        "Topshirish hujjati / Handover document",
        "Izoh / Remarks"
    ]

    # 1. Main inventory register
    ws = wb.active
    ws.title = "Inventory"
    ws.append(["LOYIHA YAKUNIY INVENTARIZATSIYASI / PROJECT CLOSURE INVENTORY"])
    ws.append(["Loyiha / Project:", "Qishloq xo'jaligini modernizatsiya qilish loyihasi / Agricultural Modernization Project"])
    ws.append(["Buyruq raqami va sanasi / Order No. and date:", "№ 42-Buyruq, 15.01.2026"])
    ws.append(["Inventarizatsiya sanasi / Inventory date:", "12.02.2026"])
    ws.append(["Eslatma / Note: USD kursini tegishli sana uchun kiriting (masalan: 12850)."])
    ws.append(headers)

    # Initial sample row to demonstrate formulas
    sample_data = [
        ("ITM-2026-001", "Dell Latitude 5540 noutbuki / Laptop", "dona", "Toshkent shahri / Tashkent City", "Axborot texnologiyalari markazi", "102-xona", "Karimov Alisher", 15, 15, 14200000, "UZS", 12850, "Yaxshi", "TT-01/2026", "Holati a'lo"),
        ("ITM-2026-002", "HP LaserJet Pro MFP 4103 printer", "dona", "Samarqand viloyati / Samarkand Region", "Viloyat qishloq xo'jaligi boshqarmasi", "Bosh ombor", "Normurodov Ilhom", 8, 7, 6500000, "UZS", 12850, "Yaxshi", "TT-02/2026", "1 dona texnik ko'rikda"),
        ("ITM-2026-003", "Tuproq tahlili laboratoriya to'plami", "komplekt", "Farg'ona viloyati / Fergana Region", "Agrokimyo tahlil markazi", "Laboratoriya 3", "Soliyev Bobur", 5, 6, 2200, "USD", 12850, "Yangi", "TT-03/2026", "1 to'plam zaxiradan kirim qilingan"),
        ("ITM-2026-004", "Garmin GPS navigatsiya uskunasi", "dona", "Buxoro viloyati / Bukhara Region", "Monitoring bo'limi", "2-ombor", "Saidov Kamol", 12, 12, 450, "USD", 12850, "Yaxshi", "TT-04/2026", "Joyida tekshirildi"),
        ("ITM-2026-005", "Ofis stoli va ergonomik kreslo", "komplekt", "Andijon viloyati / Andijan Region", "Agroinnovatsiya markazi", "Kabineta 14", "Ergashev Mansur", 20, 19, 3200000, "UZS", 12850, "Qoniqarli", "TT-05/2026", "Ta'mir talab"),
    ]

    for idx, item in enumerate(sample_data, 7):
        ws.cell(idx, 2, item[0])
        ws.cell(idx, 3, item[1])
        ws.cell(idx, 4, item[2])
        ws.cell(idx, 5, item[3])
        ws.cell(idx, 6, item[4])
        ws.cell(idx, 7, item[5])
        ws.cell(idx, 8, item[6])
        ws.cell(idx, 9, item[7])
        ws.cell(idx, 10, item[8])
        ws.cell(idx, 13, item[9])
        ws.cell(idx, 14, item[10])
        ws.cell(idx, 15, item[11])
        ws.cell(idx, 19, item[12])
        ws.cell(idx, 20, item[13])
        ws.cell(idx, 21, item[14])

    for r in range(7, 507):
        ws.cell(r, 1, f'=IF(B{r}="","",ROW()-6)')
        ws.cell(r, 11, f'=IF(OR(I{r}="",J{r}=""),"",MAX(J{r}-I{r},0))')
        ws.cell(r, 12, f'=IF(OR(I{r}="",J{r}=""),"",MAX(I{r}-J{r},0))')
        ws.cell(r, 16, f'=IF(OR(I{r}="",M{r}=""),"",I{r}*M{r})')
        ws.cell(r, 17, f'=IF(OR(J{r}="",M{r}=""),"",J{r}*M{r})')
        ws.cell(r, 18, f'=IF(OR(P{r}="",Q{r}=""),"",Q{r}-P{r})')

    # 2. Regional summary
    ws2 = wb.create_sheet("Regional Summary")
    ws2.append(["Hududlar bo'yicha jamlanma / Regional Summary"])
    ws2.append([
        "Hudud / Region", "Pozitsiyalar / Items",
        "Hisob miqdori / Book qty",
        "Haqiqiy miqdor / Actual qty",
        "Ortiqcha / Surplus", "Kamomad / Shortage",
        "UZS qiymati / UZS value",
        "USD qiymati / USD value"
    ])
    regions = [
        "Qoraqalpog'iston Respublikasi / Republic of Karakalpakstan",
        "Andijon viloyati / Andijan Region",
        "Buxoro viloyati / Bukhara Region",
        "Jizzax viloyati / Jizzakh Region",
        "Qashqadaryo viloyati / Kashkadarya Region",
        "Navoiy viloyati / Navoi Region",
        "Namangan viloyati / Namangan Region",
        "Samarqand viloyati / Samarkand Region",
        "Surxondaryo viloyati / Surkhandarya Region",
        "Sirdaryo viloyati / Syrdarya Region",
        "Toshkent viloyati / Tashkent Region",
        "Farg'ona viloyati / Fergana Region",
        "Xorazm viloyati / Khorezm Region",
        "Toshkent shahri / Tashkent City",
    ]
    for r, region in enumerate(regions, 3):
        ws2.cell(r, 1, region)
        ws2.cell(r, 2, f'=COUNTIF(Inventory!E$7:E$506,A{r})')
        ws2.cell(r, 3, f'=SUMIF(Inventory!E$7:E$506,A{r},Inventory!I$7:I$506)')
        ws2.cell(r, 4, f'=SUMIF(Inventory!E$7:E$506,A{r},Inventory!J$7:J$506)')
        ws2.cell(r, 5, f'=SUMIF(Inventory!E$7:E$506,A{r},Inventory!K$7:K$506)')
        ws2.cell(r, 6, f'=SUMIF(Inventory!E$7:E$506,A{r},Inventory!L$7:L$506)')
        ws2.cell(r, 7, f'=SUMIFS(Inventory!Q$7:Q$506,Inventory!E$7:E$506,A{r},Inventory!N$7:N$506,"UZS")')
        ws2.cell(r, 8, f'=SUMIFS(Inventory!Q$7:Q$506,Inventory!E$7:E$506,A{r},Inventory!N$7:N$506,"USD")')

    # 3. Discrepancies
    ws3 = wb.create_sheet("Discrepancies")
    ws3.append(["Tafovutlar qaydnomasi / Discrepancy Register"])
    ws3.append([
        "Item code", "Item description", "Region", "Institution",
        "Book qty", "Actual qty", "Surplus", "Shortage",
        "Difference value", "Condition", "Explanation / Izoh",
        "Corrective action / Chora", "Responsible person"
    ])
    for r in range(3, 503):
        inv = r + 4
        # Fix: Pull Item code from Inventory!B{inv} if discrepancy exists
        ws3.cell(r, 1, f'=IF(OR(Inventory!K{inv}>0,Inventory!L{inv}>0),Inventory!B{inv},"")')
        ws3.cell(r, 2, f'=IF(OR(Inventory!K{inv}>0,Inventory!L{inv}>0),Inventory!C{inv},"")')
        ws3.cell(r, 3, f'=IF(B{r}="","",Inventory!E{inv})')
        ws3.cell(r, 4, f'=IF(B{r}="","",Inventory!F{inv})')
        ws3.cell(r, 5, f'=IF(B{r}="","",Inventory!I{inv})')
        ws3.cell(r, 6, f'=IF(B{r}="","",Inventory!J{inv})')
        ws3.cell(r, 7, f'=IF(B{r}="","",Inventory!K{inv})')
        ws3.cell(r, 8, f'=IF(B{r}="","",Inventory!L{inv})')
        ws3.cell(r, 9, f'=IF(B{r}="","",Inventory!R{inv})')
        ws3.cell(r, 10, f'=IF(B{r}="","",Inventory!S{inv})')
        ws3.cell(r, 13, f'=IF(B{r}="","",Inventory!H{inv})')

    # 4. Handover register
    ws4 = wb.create_sheet("Handover")
    ws4.append(["Topshirish-qabul qilish reyestri / Handover Register"])
    ws4.append([
        "Date / Sana", "Document No. / Hujjat №",
        "Item code / Tovar kodi", "Item description / Tovar nomi",
        "Quantity / Miqdor", "Unit / Birlik",
        "Region / Hudud", "Receiving institution / Qabul qiluvchi",
        "Recipient name / Qabul qiluvchi vakili",
        "Transferor / Topshiruvchi", "Condition / Holati",
        "Supporting document / Asos hujjati",
        "Signatures / Imzolar", "Remarks / Izoh"
    ])
    for _ in range(3, 203):
        ws4.cell(_, 1).number_format = "dd.mm.yyyy"

    # 5. Inventory committee and sign-off
    ws5 = wb.create_sheet("Committee")
    ws5.append(["Inventarizatsiya komissiyasi / Inventory Committee"])
    for row in [
        ["Loyiha / Project", "Qishloq xo'jaligini modernizatsiya qilish loyihasi / Agricultural Modernization Project"],
        ["Tashkilot / Organization", "O'zbekiston Respublikasi Qishloq xo'jaligi vazirligi huzuridagi Loyiha boshqaruvi"],
        ["Buyruq raqami / Order No.", "№ 42-Buyruq"],
        ["Boshlanish sanasi / Start date", "01.02.2026"],
        ["Tugash sanasi / End date", "12.02.2026"],
        [],
        ["Lavozim / Role", "F.I.Sh. / Full name", "Imzo / Signature", "Sana / Date"],
        ["Komissiya raisi / Chairperson", "Rustamov Diyor Bahodirovich", "", "12.02.2026"],
        ["Komissiya a'zosi / Committee member", "Xoliqova Nargiza Anvarovna", "", "12.02.2026"],
        ["Komissiya a'zosi / Committee member", "Tursunov Jamshid Erkinovich", "", "12.02.2026"],
        ["Buxgalter / Accountant", "Mirzayeva Lola Shavkatovna", "", "12.02.2026"],
        ["Moddiy javobgar shaxs / Responsible custodian", "Karimov Alisher Zokirovich", "", "12.02.2026"],
        [],
        ["Xulosa / Conclusion", "Inventarizatsiya natijasida loyiha doirasida xarid qilingan asosiy vositalar va tovar-moddiy boyliklar ko'zdan kechirildi hamda qabul qiluvchi tashkilotlarga topshirishga tayyorlandi."],
        ["Aniqlangan tafovutlar / Discrepancies identified", "Jami 2 ta pozitsiya bo'yicha kichik tafovutlar aniqlandi (1 dona printer ta'mirda, 1 komplekt laboratoriya jihozi ortiqcha zaxira)."],
        ["Tavsiya etilgan choralar / Recommended actions", "Ortiqcha uskunani belgilangan tartibda kirim qilish, ta'mirdagi uskunani servis markazidan qabul qilib olgach topshirish tavsiya etiladi."],
        ["Rahbar qarori / Management decision", "Komissiya xulosasi tasdiqlansin va topshirish-qabul qilish dalolatnomalari rasmiylashtirilsin."],
    ]:
        ws5.append(row)

    # Formatting
    navy = "1F4E78"
    blue = "D9EAF7"
    white = "FFFFFF"
    thin = Side(style="thin", color="B7C9D6")

    for sheet in wb.worksheets:
        sheet.freeze_panes = "A7" if sheet.title == "Inventory" else "A3"
        sheet.sheet_view.showGridLines = False
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(
                    vertical="center", wrap_text=True
                )
                cell.border = Border(bottom=thin)
        for cell in sheet[1]:
            cell.fill = PatternFill("solid", fgColor=navy)
            cell.font = Font(color=white, bold=True, size=13)
        header_row = 6 if sheet.title == "Inventory" else (
            2 if sheet.title in ("Regional Summary", "Discrepancies", "Handover") else 7
        )
        if sheet.title != "Committee":
            for cell in sheet[header_row]:
                cell.fill = PatternFill("solid", fgColor=blue)
                cell.font = Font(bold=True, color=navy)
        for col in range(1, sheet.max_column + 1):
            letter = get_column_letter(col)
            sheet.column_dimensions[letter].width = 19
        sheet.row_dimensions[1].height = 30

    for col, width in {"C":32, "F":30, "G":24, "H":25,
                        "S":20, "T":27, "U":28}.items():
        ws.column_dimensions[col].width = width
    ws.auto_filter.ref = "A6:U506"
    ws2.column_dimensions["A"].width = 49
    ws3.column_dimensions["B"].width = 32
    ws4.column_dimensions["D"].width = 32
    ws5.column_dimensions["A"].width = 42
    ws5.column_dimensions["B"].width = 32

    wb.save(FILE)
    print(f"Created: {FILE}")

if __name__ == "__main__":
    generate_inventory_workbook()
