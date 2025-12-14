const XLSX = require("xlsx");

// 1. خواندن فایل اکسل
const workbook = XLSX.readFile("C:/Users/HP/Desktop/cloutesdec/mmm.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // format: array of arrays

// data[0] = header → فرض می‌کنیم ردیف اول هدر است
// ستون A = کد کالا
// ستون B = نام کالا

let itemsMap = new Map(); // برای حذف تکراری‌ها

for (let i = 1; i < data.length; i++) {
  const code = data[i][0];
  const name = data[i][1];
  if (!code || !name) continue;

  // اگر برنامه فقط یکبار هر کالا را نیاز دارد:
  if (!itemsMap.has(name)) {
    itemsMap.set(name, code);
  }
}

// 2. تبدیل Map به آرایه و مرتب‌سازی بر اساس نوع کالا
const sorted = Array.from(itemsMap).sort((a, b) => a[0].localeCompare(b[0]));

// 3. نوشتن در ستون‌های D و E
sorted.forEach((item, index) => {
  const row = index + 2; // شروع از ردیف 2 برای اینکه هدر خراب نشه

  // ستون D = نام کالا
  sheet[`D${row}`] = { t: "s", v: item[0] };

  // ستون E = کد کالا
  sheet[`E${row}`] = { t: "s", v: item[1] };
});

// ست کردن هدرها
sheet["D1"] = { t: "s", v: "نام کالا (یونیک)" };
sheet["E1"] = { t: "s", v: "کد کالا" };

// 4. ذخیره فایل جدید
XLSX.writeFile(workbook, "result.xlsx");

console.log("تمام شد! خروجی در فایل result.xlsx ذخیره شد.");