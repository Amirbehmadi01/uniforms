const express = require("express");
const xlsx = require("xlsx");
const path = require("path");

const app = express();
const PORT = 3000;

// مسیر اکسل اصلی
const excelPath = path.resolve("C:/Users/HP/Desktop/cloutesdec/excel/Book1.xlsx");

// خواندن فایل اکسل
let workbook;
try {
  workbook = xlsx.readFile(excelPath);
  console.log("✅ Excel file loaded successfully!");
} catch (error) {
  console.error("❌ Error loading Excel file:", error.message);
  process.exit(1);
}

// فرض می‌کنیم شیت اول مد نظرته
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(sheet);

// کلمات کلیدی
const keywords = ["پیراهن", "شلوار", "کت و شلوار", "کلاه", "کفش", "کاپشن", "بادگیر"];

// فیلتر داده‌ها
const filtered = data.filter((row) => {
  const item = row["نام کالا"] || row["C"] || "";
  return keywords.some((kw) => item.includes(kw));
});

// ✅ مسیر برای دانلود اکسل فیلترشده
app.get("/download", (req, res) => {
  try {
    const newWorkbook = xlsx.utils.book_new();
    const newSheet = xlsx.utils.json_to_sheet(filtered);
    xlsx.utils.book_append_sheet(newWorkbook, newSheet, "Filtered Data");

    const downloadPath = path.join(__dirname, "filtered.xlsx");
    xlsx.writeFile(newWorkbook, downloadPath);

    res.download(downloadPath, "FilteredData.xlsx");
  } catch (err) {
    console.error("❌ Error creating Excel file:", err.message);
    res.status(500).send("Error generating Excel file");
  }
});

// ✅ صفحه‌ی اصلی
app.get("/", (req, res) => {
  res.send(`
    <html lang="fa">
      <head>
        <meta charset="UTF-8" />
        <title>فهرست کالاها</title>
        <style>
          body { font-family: sans-serif; direction: rtl; background: #f5f5f5; padding: 20px; }
          table { border-collapse: collapse; width: 100%; background: white; }
          th, td { border: 1px solid #ccc; padding: 8px; text-align: center; }
          th { background: #4caf50; color: white; }
          h2 { color: #333; }
          button {
            background-color: #2196f3;
            color: white;
            border: none;
            padding: 10px 20px;
            margin-bottom: 15px;
            cursor: pointer;
            border-radius: 8px;
            font-size: 15px;
          }
          button:hover { background-color: #1976d2; }
        </style>
      </head>
      <body>
        <h2>کارمندانی که کالا دریافت کرده‌اند</h2>
        <button onclick="window.location.href='/download'">📥 دانلود اکسل</button>
        <table>
          <thead>
            <tr>
              <th>کد پرسنلی</th>
              <th>نام کارمند</th>
              <th>نام کالا</th>
              <th>تاریخ تحویل</th>
            </tr>
          </thead>
          <tbody>
            ${filtered.map(row => `
              <tr>
                <td>${row["کد پرسنلی"] || ""}</td>
                <td>${row["نام کارمند"] || ""}</td>
                <td>${row["نام کالا"] || ""}</td>
                <td>${row["تاریخ تحویل"] || ""}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </body>
    </html>
  `);
});

// ✅ راه‌اندازی سرور
app.listen(PORT, () => {
  console.log("===========================================");
  console.log(`✅ Server is running at: http://localhost:${PORT}`);
  console.log("===========================================");
});