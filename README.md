# Loyiha Yakuniy Inventarizatsiyasi / Project Closure Inventory System

Ushbu loyiha yakuniy inventarizatsiyasini boshqarish, hududlar bo'yicha tovar-moddiy boyliklar jamlanmasini shakllantirish, tafovutlar (ortiqcha va kamomad) tahlili, topshirish-qabul qilish dalolatnomalarini yaratish hamda inventarizatsiya komissiyasi xulosalarini rasmiylashtirish uchun to'liq interaktiv tizim.

---

## 🚀 Tizim imkoniyatlari (Features)

1. **Asosiy inventarizatsiya reyestri (1-varaq / Main Inventory)**:
   - 21 ta ustundan iborat to'liq audit jadvali (№, Tovar kodi, Tovar nomi, Birlik, Hudud, Qabul qiluvchi tashkilot, Saqlash joyi, Mas'ul shaxs, Hisob miqdori, Haqiqiy miqdor, Ortiqcha, Kamomad, Birlik narxi, Valyuta, USD kursi, Hisob qiymati, Haqiqiy qiymat, Farq, Holati, Hujjat №, Izoh).
   - Real vaqtda formulalarni avtomatik hisoblash (Ortiqcha = MAX(J-I, 0), Kamomad = MAX(I-J, 0), Qiymatlar va Tafovut balansi).
   - Qidiruv, hududlar bo'yicha filtrlash, holatlar va valyutalar bo'yicha saralash.
   - Yangi tovar qo'shish va tahrirlash (jonli hisob-kitoblar ko'rinishi bilan).

2. **Hududlar bo'yicha jamlanma (2-varaq / Regional Summary)**:
   - O'zbekiston Respublikasining barcha 14 ta hududi (Qoraqalpog'iston Respublikasi, 12 ta viloyat va Toshkent shahri) bo'yicha avtomatik jamlangan ma'lumotlar.
   - Pozitsiyalar soni, hisob va haqiqiy miqdorlar, ortiqcha va kamomadlar, UZS va USD qiymatlari.
   - Hududlar bo'yicha aktivlar taqsimoti infografikasi va bir bosishda hudud tovarlariga o'tish (Drill-down filtrlash).

3. **Tafovutlar qaydnomasi (3-varaq / Discrepancies Register)**:
   - Ortiqcha yoki kamomad aniqlangan har bir tovar bo'yicha avtomatik filtrlangan ro'yxat.
   - Har bir holat bo'yicha mas'ul shaxsning yozma izohi (tushuntirish xati) va ko'riladigan chora (kirim qilish, hisobdan chiqarish, servisga yo'naltirish) maydonlari (avtomatik saqlanadi).

4. **Topshirish-qabul qilish reyestri (4-varaq / Handover Register)**:
   - Benefitsiar tashkilotlarga aktivlarni topshirish dalolatnomalari reyestri.
   - Tovarni tanlaganda avtomatik to'ldiriladigan yangi akt yaratish moduli.
   - A4 formatidagi rasmiy, ikki tilli (O'zbek / Ingliz) topshirish-qabul qilish dalolatnomasini chop etish (Print / PDF).

5. **Inventarizatsiya komissiyasi va yakuniy bayonnoma (5-varaq / Committee Protocol)**:
   - Komissiya tarkibi (Rais, a'zolar, bosh buxgalter, moddiy javobgar shaxs) va imzolar.
   - Xulosa, aniqlangan tafovutlar, tavsiya etilgan choralar va rahbar qarori.
   - Rasmiy vazirlik / loyiha inventarizatsiya bayonnomasini chop etish (Print / PDF).

6. **Excel (.xlsx) integratsiyasi**:
   - Brauzerning o'zidan barcha 5 ta varaqni o'z ichiga olgan haqiqiy `Project_Inventory_UZ_EN.xlsx` faylini yuklab olish (Export).
   - Excel faylni tizimga qayta yuklash (Import) va ma'lumotlarni yangilash.
   - Python dasturchilari uchun tayyor `generate_inventory_excel.py` skripti mavjud.

---

## 💻 Ishga tushirish (Running Locally)

### 1. Veb-ilovadan foydalanish:
```bash
# Kutubxonalarni o'rnatish
npm install

# Mahalliy dev serverni ishga tushirish
npm run dev
```
Brauzerda oching: `http://127.0.0.1:5173/`

### 2. Python skripti orqali to'g'ridan-to'g'ri Excel yaratish:
```bash
# openpyxl kutubxonasi o'rnatilgan bo'lishi kerak:
pip install openpyxl

# Skriptni ishga tushirish:
python3 generate_inventory_excel.py
```
Natijada loyiha papkasida `Project_Inventory_UZ_EN.xlsx` fayli hosil bo'ladi.
# pcis
