const express = require("express");
const xlsx = require("xlsx");
const path = require("path");
const jalaali = require("jalaali-js");

const app = express();
const PORT = process.env.PORT || 3000;

// ---- مسیر فایل‌ها ----
const filteredPath = path.join(__dirname, "excel", "filteredData.xlsx");
const bottlePath = path.join(__dirname, "excel", "bottle.xlsx");

// ---- کالاهای مورد بررسی ----
const keywords = [
  "پیراهن",
  "شلوار",
  "کت و شلوار",
  "مانتو",
  "کفش",
  "کاپشن",
  "جلیقه کت و شلوار",
  "مقنعه"
];

// ---- Utility ----
function clean(str) {
  if (!str) return "";
  return String(str)
    .trim()
    .replace(/\s+/g, " ")     
    .replace(/‌/g, "")        
    .replace(/[ي]/g, "ی")
    .replace(/[ك]/g, "ک");
}

// ---- تبدیل تاریخ شمسی (yyyy/mm/dd) به میلادی ----
function shamsiToDate(str) {
  if (!str) return null;
  str = clean(str);
  const p = str.split(/[\/\-]/).map(n => parseInt(n, 10));
  if (p.length !== 3 || p.some(isNaN)) return null;
  const g = jalaali.toGregorian(p[0], p[1], p[2]);
  return new Date(g.gy, g.gm - 1, g.gd);
}

// ---- Load filteredData ----
const wb1 = xlsx.readFile(filteredPath);
const sheet1 = wb1.Sheets[wb1.SheetNames[0]];
let filteredData = xlsx.utils.sheet_to_json(sheet1).map(r => ({
  code: clean(r["کد پرسنلی"]),
  name: clean(r["نام کارمند"]),
  item: clean(r["نام کالا"]),
  date: clean(r["تاریخ تحویل"])
}));
console.log("✅ filteredData loaded");

// ---- Load bottle ----
const wb2 = xlsx.readFile(bottlePath);
const sheet2 = wb2.Sheets[wb2.SheetNames[0]];
let bottleData = xlsx.utils.sheet_to_json(sheet2, { defval: "" }).map(r => ({
  code: clean(r["A"] || r["کد پرسنلی"]),
  name: clean(r["E"] || r["نام کارمند"]),
  date: clean(r["D"] || r["تاریخ تحویل"])
}));
console.log("✅ Bottle data loaded");

app.use(express.urlencoded({ extended: true }));

// ---- صفحه اصلی ----
app.get("/", (req, res) => {
  res.send(`
  <html lang="fa">
  <head>
    <meta charset="UTF-8">
    <style>
      body { direction: rtl; font-family: sans-serif; background:#f5f5f5;
             display:flex; justify-content:center; align-items:center; height:100vh; }
      .box { text-align:center; }
      input { padding: 15px; width: 350px; font-size: 18px; }
      button { padding: 15px 25px; background:#2196f3; color:white; border:none;
               cursor:pointer; font-size:18px; border-radius:6px; }
    </style>
    <title>جستجوی کارمند</title>
  </head>
  <body>
    <div class="box">
      <h2>جستجوی آخرین تحویل کالا</h2>
      <form method="POST" action="/search">
        <input name="query" placeholder="نام یا کد پرسنلی را وارد کنید" required />
        <button type="submit">جستجو 🔍</button>
      </form>
    </div>
  </body>
  </html>
  `);
});

// ---- جستجو ----
app.post("/search", (req, res) => {
  const q = clean(req.body.query);

  // --- همه رکوردها از filteredData و Bottle
  const allRecords = [...filteredData];

  // برای افرادی که فقط در Bottle هستند، یک رکورد موقت بساز
  bottleData.forEach(b => {
    if (!allRecords.some(r => r.code === b.code)) {
      allRecords.push({
        code: b.code,
        name: b.name,
        item: null,
        date: null
      });
    }
  });

  // --- فیلتر بر اساس query
  let filtered = allRecords.filter(r =>
    r.code.includes(q) || r.name.includes(q)
  );

  if (!filtered.length) {
    return res.send(`<h3>هیچ رکوردی یافت نشد برای: ${q}</h3><a href="/">بازگشت</a>`);
  }

  // --- گروه‌بندی بر اساس کد پرسنلی
  const grouped = {};
  filtered.forEach(r => {
    if (!grouped[r.code]) grouped[r.code] = [];
    grouped[r.code].push(r);
  });

  const today = new Date();

  // --- ساخت جدول HTML
  let table = `
    <table style="border-collapse: collapse; width:100%; background:white;">
      <thead>
        <tr>
          <th>نام</th>
          <th>کد پرسنلی</th>
          ${keywords.map(k => `<th>${k}</th>`).join("")}
          <th>قمقمه</th>
        </tr>
      </thead>
      <tbody>
  `;

  for (let code in grouped) {
    const rows = grouped[code];
    const name = rows[0].name;

    // آخرین تاریخ هر کالا
    const lastDates = {};
    const lastDatesText = {};
    keywords.forEach(k => { lastDates[k] = null; lastDatesText[k] = "-"; });

    rows.forEach(r => {
      if (!r.item || !r.date) return;
      const d = shamsiToDate(r.date);
      if (!d) return;
      keywords.forEach(k => {
        if (r.item.includes(k)) {
          if (!lastDates[k] || d > lastDates[k]) {
            lastDates[k] = d;
            lastDatesText[k] = r.date;
          }
        }
      });
    });

    // ---- آخرین تاریخ قمقمه
    const relatedBottle = bottleData.filter(b =>
      b.code === code || b.name === name
    );

    let bottleDate = null;
    let bottleDateText = "-";
    relatedBottle.forEach(b => {
      const d = shamsiToDate(b.date);
      if (!d) return;
      if (!bottleDate || d > bottleDate) {
        bottleDate = d;
        bottleDateText = b.date;
      }
    });

    // ---- ساخت ردیف
    let rowHTML = `<tr>
      <td style="text-align:right; font-weight:bold;">${name}</td>
      <td>${code}</td>`;

    keywords.forEach(k => {
      const d = lastDates[k];
      if (!d) {
        rowHTML += `<td style="background:#eee;">-</td>`;
      } else {
        const diff = (today - d)/(1000*60*60*24);
        const color = diff >= 365 ? "rgba(0,255,0,0.3)" : "rgba(255,0,0,0.3)";
        rowHTML += `<td style="background:${color};">${lastDatesText[k]}</td>`;
      }
    });

    // قمقمه
    if (!bottleDate) {
      rowHTML += `<td style="background:#eee;">-</td>`;
    } else {
      const diff = (today - bottleDate)/(1000*60*60*24);
      const color = diff >= 365 ? "rgba(0,255,0,0.3)" : "rgba(255,0,0,0.3)";
      rowHTML += `<td style="background:${color};">${bottleDateText}</td>`;
    }

    rowHTML += "</tr>";
    table += rowHTML;
  }

  table += "</tbody></table>";

  res.send(`
  <html lang="fa">
    <head>
      <meta charset="UTF-8">
      <style>
        body { font-family:sans-serif; direction:rtl; padding:20px; background:#f5f5f5; }
        th, td { border:1px solid #ccc; padding:8px; text-align:center; }
        th { background:#4caf50; color:white; }
      </style>
      <title>نتایج جستجو</title>
    </head>
    <body>
      <h2>نتایج جستجو برای "${q}"</h2>
      ${table}
      <a href="/">بازگشت</a>
    </body>
  </html>
  `);
});

// ---- اجرا ----
app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));