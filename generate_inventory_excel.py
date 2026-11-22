#!/usr/bin/env python3
"""
PEMS-UZ: PROJECT EQUIPMENT MANAGEMENT SYSTEM (UZBEKISTAN)
Developed by: Abdiyev Salimkhon Sanjar ugli
Система учёта, распределения, поступления, передачи, мониторинга и инвентаризации оборудования (АБР / АБИИ).
Generates Project_Inventory_UZ_EN.xlsx matching openpyxl specifications.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FILE = "Project_Inventory_UZ_EN.xlsx"

def generate_inventory_workbook():
    wb = Workbook()

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="163A5F", end_color="163A5F", fill_type="solid")
    title_font = Font(name="Calibri", size=13, bold=True, color="163A5F")
    bold_font = Font(name="Calibri", size=11, bold=True)
    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # 1. SUMMARY SHEET
    ws0 = wb.active
    ws0.title = "Dashboard_Summary"
    ws0.append(["PEMS-UZ: PROJECT EQUIPMENT MANAGEMENT SYSTEM (UZBEKISTAN)"])
    ws0.append(["Система учёта, распределения, поступления, передачи, мониторинга и инвентаризации оборудования (АБР / АБИИ)"])
    ws0.append(["Разработчик / Developed by:", "Abdiyev Salimkhon Sanjar ugli"])
    ws0.append(["Основание:", "Постановление Президента РУз № PQ-3818 (16.10.2023) | ТЭО № TEO-ADB/AIIB-2023/V4"])
    ws0.append(["Курс валюты:", "1 USD = 12,850 UZS"])
    ws0.append([])
    ws0.append(["№", "Показатель жизненного цикла оборудования", "Количество", "Ед. изм.", "Статус контроля"])

    kpis = [
      ("1", "Утверждено по ТЭО (Feasibility Study)", 125, "комплект", "Утверждено"),
      ("2", "Закуплено по контрактам (Contracted)", 125, "комплект", "Контракты подписаны"),
      ("3", "Поступило на склад / УзАСБО (Delivered)", 125, "комплект", "Оприходовано в УзАСБО"),
      ("4", "Передано лабораториям (Transferred to End-Users)", 120, "комплект", "Акты приёма-передачи"),
      ("5", "Установлено и смонтировано (Installed)", 116, "комплект", "Инженеры поставщика"),
      ("6", "Введено в эксплуатацию (Commissioned)", 114, "комплект", "Акты ввода + обучение"),
      ("7", "Прошло мониторинг ГРП (Monitored by PIU)", 114, "единиц", "Выездной контроль"),
      ("8", "Выявленные дефекты / рекламации (Defective)", 2, "единиц", "На гарантийном ремонте"),
      ("9", "Проинвентаризировано (Physical Inventory)", 120, "единиц", "Сличение с УзАСБО"),
      ("10", "Нераспределённый остаток / Резерв (Remaining)", 5, "комплект", "Центральный резерв")
    ]
    for row in kpis:
        ws0.append(list(row))

    # 2. INVENTORY AUDIT SHEET
    ws1 = wb.create_sheet("Inventory_Audit")
    headers_inv = [
        "№", "Tovar kodi / Item code", "Tovar nomi / Item description", "Seriya raqami / Serial No",
        "Inventar raqami / Inv No", "O'lchov birligi / Unit", "Hudud / Region",
        "Qabul qiluvchi tashkilot / Receiving institution", "Saqlash joyi / Location",
        "Mas'ul shaxs / Responsible person", "Hisob miqdori / Book qty", "Haqiqiy miqdor / Actual qty",
        "Ortiqcha / Surplus", "Kamomad / Shortage", "Birlik narxi (USD)", "Valyuta", "USD kursi",
        "Hisob qiymati (USD)", "Haqiqiy qiymat (USD)", "Farq / Diff", "Texnik holati",
        "Topshirish hujjati / Act", "Transfer №", "Izoh / Remarks"
    ]
    ws1.append(["PEMS-UZ: 7. ИНВЕНТАРИЗАЦИЯ ОБОРУДОВАНИЯ / INVENTORY AUDIT (ADB / AIIB)"])
    ws1.append(["Loyiha / Project:", "PEMS-UZ - АБР / АБИИ лойиҳаси | Dev: Abdiyev Salimkhon Sanjar ugli"])
    ws1.append(["Asos:", "PQ-3818 | TEO-ADB/AIIB-2023/V4 | Buyruq № 42, 15.01.2026"])
    ws1.append(["Sana / Date:", "12.02.2026"])
    ws1.append(["Eslatma: USD kursi: 12850 UZS."])
    ws1.append(headers_inv)

    sample_items = [
        ("EQP-PCR-01", "Real-Time PCR Amplifikatori (CFX96 Touch Deep Well)", "BR-PCR-98214-01", "INV-SAM-00421", "komplekt", "Samarqand viloyati", "Sanepidqo'm Samarqand", "204-xona", "Jo'rayev Olimjon", 1, 1, 38500, "USD", 12850, "Ishga yaroqli", "ACT-01/2024", "TRF-2024-001", "Ekspluatatsiyada"),
        ("EQP-PCR-01", "Real-Time PCR Amplifikatori (CFX96 Touch Deep Well)", "BR-PCR-98214-02", "INV-FAR-00389", "komplekt", "Farg'ona viloyati", "Sanepidqo'm Farg'ona", "14-xona", "Soliyev Bobur", 1, 1, 38500, "USD", 12850, "Ishga yaroqli", "ACT-02/2024", "TRF-2024-001", "Ishchi holatda"),
        ("EQP-HPLC-02", "Yuqori unumli suyuqlik xromatografi (Agilent 1260)", "AG-HPLC-44102-01", "INV-TAS-REF-0012", "komplekt", "Toshkent shahri", "Milliy referens laboratoriya", "108-xona", "Karimov Alisher", 1, 1, 64200, "USD", 12850, "Foydalanishga topshirilgan", "ACT-03/2024", "TRF-2024-002", "Yuqori unumli ish rejimida"),
        ("EQP-BSC-03", "Biosavfsizlik laminar shkafi II sinf A2", "ESCO-BSC-77821-05", "INV-BUX-00215", "dona", "Buxoro viloyati", "Sanepidqo'm Buxoro", "Bakteriologiya", "Saidov Kamol", 2, 2, 8900, "USD", 12850, "Ishga yaroqli", "ACT-04/2024", "TRF-2024-003", "HEPA filtrlar soz"),
        ("EQP-ACV-04", "Laboratoriya vertikal avtoklavi 85L (Tuttnauer)", "TT-AUT-55219-03", "INV-AND-00194", "dona", "Andijon viloyati", "Sanepidqo'm Andijon", "Sterilizatsiya", "Ergashev Mansur", 1, 1, 11400, "USD", 12850, "Ta'mirlash talab etiladi", "ACT-05/2024", "TRF-2024-004", "Bug' sizishi aniqlangan"),
        ("EQP-BSC-03", "Biosavfsizlik laminar shkafi II sinf A2 (Zaxira)", "ESCO-BSC-77821-19", "INV-XOR-00088", "dona", "Xorazm viloyati", "Sanepidqo'm Xorazm", "Bosh ombor", "Matyoqubov Jamol", 1, 2, 8900, "USD", 12850, "Yangi / New", "ACT-08/2024", "TRF-2024-003", "1 dona hisobdan ortiqcha")
    ]

    for idx, item in enumerate(sample_items, 7):
        b_qty = item[9]
        a_qty = item[10]
        price = item[11]
        ws1.cell(idx, 1, idx - 6)
        ws1.cell(idx, 2, item[0])
        ws1.cell(idx, 3, item[1])
        ws1.cell(idx, 4, item[2])
        ws1.cell(idx, 5, item[3])
        ws1.cell(idx, 6, item[4])
        ws1.cell(idx, 7, item[5])
        ws1.cell(idx, 8, item[6])
        ws1.cell(idx, 9, item[7])
        ws1.cell(idx, 10, item[8])
        ws1.cell(idx, 11, b_qty)
        ws1.cell(idx, 12, a_qty)
        ws1.cell(idx, 13, f'=MAX(L{idx}-K{idx},0)')
        ws1.cell(idx, 14, f'=MAX(K{idx}-L{idx},0)')
        ws1.cell(idx, 15, price)
        ws1.cell(idx, 16, item[12])
        ws1.cell(idx, 17, item[13])
        ws1.cell(idx, 18, f'=K{idx}*O{idx}')
        ws1.cell(idx, 19, f'=L{idx}*O{idx}')
        ws1.cell(idx, 20, f'=S{idx}-R{idx}')
        ws1.cell(idx, 21, item[14])
        ws1.cell(idx, 22, item[15])
        ws1.cell(idx, 23, item[16])
        ws1.cell(idx, 24, item[17])

    # 3. REGIONAL SUMMARY
    ws2 = wb.create_sheet("Regional_Summary")
    ws2.append(["PEMS-UZ: РАСПРЕДЕЛЕНИЕ ОБОРУДОВАНИЯ ПО РЕГИОНАМ (14 РЕГИОНОВ)"])
    ws2.append(["№", "Регион Республики Узбекистан", "Позиций", "Учётное количество", "Фактическое количество", "Излишек", "Недостача", "Общая стоимость USD"])
    regions = [
        "Toshkent shahri", "Toshkent viloyati", "Andijon viloyati", "Buxoro viloyati",
        "Farg'ona viloyati", "Jizzax viloyati", "Namangan viloyati", "Navoiy viloyati",
        "Qashqadaryo viloyati", "Qoraqalpog'iston Respublikasi", "Samarqand viloyati",
        "Sirdaryo viloyati", "Surxondaryo viloyati", "Xorazm viloyati"
    ]
    for r_idx, reg in enumerate(regions, 3):
        ws2.append([r_idx - 2, reg, 1, 1, 1, 0, 0, 38500])

    # Apply Styling
    for sheet in [ws0, ws1, ws2]:
        sheet.column_dimensions['A'].width = 6
        sheet.column_dimensions['B'].width = 24
        sheet.column_dimensions['C'].width = 38
        sheet.column_dimensions['D'].width = 22
        sheet.column_dimensions['E'].width = 22
        sheet.column_dimensions['F'].width = 16
        sheet.column_dimensions['G'].width = 24
        sheet.column_dimensions['H'].width = 30

    wb.save(FILE)
    print(f"Created: {FILE}")

if __name__ == "__main__":
    generate_inventory_workbook()
