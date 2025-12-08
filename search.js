const express = require("express");
const xlsx = require("xlsx");
const path = require("path");
const bodyParser = require("body-parser");
const jalaali = require("jalaali-js");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

const filteredPath = path.join(__dirname, "excel", "FilteredData.xlsx");
const bottlePath = path.join(__dirname, "excel", "bottle.xlsx");

// قمقمه در لیست اصلی — فقط همین بخش نیاز داشت
const keywords = ["پیراهن","شلوار","کت و شلوار","مانتو","کفش","کاپشن","جلیقه کت و شلوار","مقنعه","قمقمه"];

function clean(str){
    if(!str) return "";
    return String(str).trim().replace(/\s+/g," ").replace(/‌/g,"")
        .replace(/[ي]/g,"ی").replace(/[ك]/g,"ک");
}

function parseShamsiToDate(sh){
    if(!sh) return null;
    const s = String(sh).trim().replace(/-/g,'/');
    const parts = s.split('/');
    if(parts.length !== 3) return null;
    let jy = parseInt(parts[0],10);
    let jm = parseInt(parts[1],10);
    let jd = parseInt(parts[2],10);
    if(isNaN(jy)||isNaN(jm)||isNaN(jd)) return null;
    const g = jalaali.toGregorian(jy,jm,jd);
    return new Date(g.gy, g.gm-1, g.gd);
}

function escapeHtml(str){
    if(!str) return "";
    return String(str).replace(/\\/g,'\\\\').replace(/'/g,"\\'").replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function dateToShamsi(d){
    if(!d || !(d instanceof Date)) return "";
    const j = jalaali.toJalaali(d);
    return j.jy+"/"+String(j.jm).padStart(2,'0')+"/"+String(j.jd).padStart(2,'0');
}

let filteredData = [];
let bottleData = [];

function loadFilteredData(){
    try{
        const wb = xlsx.readFile(filteredPath);
        const sheet = wb.Sheets[wb.SheetNames[0]];
        filteredData = xlsx.utils.sheet_to_json(sheet,{defval:""})
        .map(r=>({
            code: clean(r["کد پرسنلی"]||r["A"]),
            name: clean(r["نام کارمند"]||r["E"]),
            item: clean(r["نام کالا"]||r["C"]||""),
            date: String(r["تاریخ تحویل"]||r["D"]||"").replace(/-/g,'/')
        }));
    }catch(e){
        console.error("Could not load filtered data:",e.message);
        filteredData=[];
    }
}
loadFilteredData();

function loadBottleData(){
    try{
        const wb = xlsx.readFile(bottlePath);
        const sheet = wb.Sheets[wb.SheetNames[0]];
        bottleData = xlsx.utils.sheet_to_json(sheet,{defval:""})
        .map(r=>({
            code: clean(r["A"]||r["کد پرسنلی"]),
            name: clean(r["E"]||r["نام کارمند"]),
            date: String(r["D"]||r["تاریخ تحویل"]||"").replace(/-/g,'/')
        }));
    }catch(e){
        console.error("Could not load bottle data:",e.message);
        bottleData=[];
    }
}
loadBottleData();


// ----------------------------------------------------
// قمقمه تبدیل به ردیف عادی می‌شود
// ----------------------------------------------------
function mergeBottleIntoMain(){
    bottleData.forEach(b=>{
        filteredData.push({
            code: b.code,
            name: b.name,
            item: "قمقمه",
            date: b.date
        });
    });
}
mergeBottleIntoMain();


// ================= HOME ====================
app.get('/', (req,res)=>{
    res.send(
        '<!doctype html><html lang="fa"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
        '<title>جستجوی کالا</title>' +
        '<script src="https://cdn.jsdelivr.net/npm/jalaali-js/dist/jalaali.min.js"></script>' +
        '<style>body{font-family:sans-serif;direction:rtl;background:#f5f5f5;display:flex;justify-content:center;align-items:center;height:100vh}form{text-align:center}input,button{padding:10px;font-size:16px;margin:5px}</style>' +
        '</head><body>' +
        '<form method="POST" action="/search"><h2>جستجوی آخرین تحویل کالا</h2>' +
        '<input name="query" placeholder="نام یا کد پرسنلی" required />' +
        '<button type="submit">جستجو</button>' +
        '</form></body></html>'
    );
});


// ================= SEARCH ====================
app.post('/search', (req,res)=>{
    const q = clean(req.body.query||"");

    const allRecords = [...filteredData.map(r=>({...r}))];

    const filtered = allRecords.filter(r =>
        (r.code && r.code.includes(q)) ||
        (r.name && r.name.includes(q))
    );

    if(!filtered.length){
        return res.send('<!doctype html><html lang="fa"><head><meta charset="utf-8"><title>نتیجه</title></head><body>' +
            `<h3>هیچ رکوردی یافت نشد برای: ${escapeHtml(q)}</h3><a href="/">بازگشت</a></body></html>`);
    }

    const grouped = {};
    filtered.forEach(r=>{
        const key = r.code || ("name:"+r.name);
        if(!grouped[key]) grouped[key]=[];
        grouped[key].push(r);
    });

    let table = '<table border="1" style="border-collapse:collapse;width:100%;background:white;"><thead><tr>' +
        '<th>نام</th><th>کد پرسنلی</th>';

    keywords.forEach(k=> table+=`<th>${k}</th>`);
    table += '<th>به‌روزرسانی</th></tr></thead><tbody>';

    const today = new Date();

    for(const key in grouped){
        const rows = grouped[key];
        const name = rows.find(r=>r.name)?.name || "";
        const code = rows.find(r=>r.code)?.code || "";

        const lastText = {};
        const lastDate = {};
        keywords.forEach(k=>{ lastText[k]="-"; lastDate[k]=null; });

        rows.forEach(r=>{
            if(!r.item || !r.date) return;
            keywords.forEach(k=>{
                if(r.item.includes(k)){
                    let type = r.item.includes('استوک') ? 'استوک' : '';
                    lastText[k] = (type ? 'استوک - ' : '') + r.date;
                    lastDate[k] = r.date;
                }
            });
        });

        let row = `<tr><td>${escapeHtml(name)}</td><td>${escapeHtml(code)}</td>`;

        keywords.forEach(k=>{
            const txt = lastText[k];
            const dstr = lastDate[k];
            let style='';
            if(dstr && dstr!="-"){
                const d=parseShamsiToDate(dstr);
                if(d){
                    const diffDays=Math.floor((today-d)/(1000*60*60*24));
                    if(diffDays>365) style='background:rgba(144,238,144,0.5)';
                    else style='background:rgba(240,128,128,0.35)';
                }
            }
            row+=`<td style="${style}">${escapeHtml(txt)}</td>`;
        });

        row += `<td><button onclick="openUpdate('${escapeHtml(code)}','${escapeHtml(name)}')">به‌روزرسانی</button></td>`;
        row += `</tr>`;
        table += row;
    }

    table+='</tbody></table>';

    const modalOptions = keywords.map(k=> `<option value="${k}">${k}</option>`).join('');

    res.send(
        '<!doctype html><html lang="fa"><head><meta charset="utf-8">'+
        '<meta name="viewport" content="width=device-width,initial-scale=1">'+
        '<title>نتایج جستجو</title>'+
        '<script src="https://cdn.jsdelivr.net/npm/jalaali-js/dist/jalaali.min.js"></script>'+
        '<style>body{font-family:sans-serif;direction:rtl;padding:20px;background:#f5f5f5}th,td{border:1px solid #ccc;padding:8px;text-align:center}th{background:#4caf50;color:white}#modal{position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.5);display:none;justify-content:center;align-items:center}#box{background:white;padding:20px;width:380px;border-radius:8px;text-align:center}</style>'+
        '</head><body>'+
        '<h2>نتایج جستجو برای: ' + escapeHtml(q) + '</h2>' +
        table +
        '<div id="modal"><div id="box">' +
        '<h3 id="updateItem">آیتم</h3>' +
        '<label>نام کارمند:</label><br><input id="updateName" readonly style="width:90%;padding:6px;margin:6px 0"><br>' +
        '<label>آیتم:</label><br>' +
        '<select id="updateField" style="width:80%;padding:6px;margin:6px 0"><option value="">-- انتخاب آیتم --</option>' +
        modalOptions +
        '</select><br>' +
        '<label>وضعیت:</label><br>' +
        '<select id="updateType" style="width:90%;padding:6px;margin:6px 0"><option value="نو">نو</option><option value="استوک">استوک</option></select><br>' +
        '<label>تاریخ:</label><br>' +
        '<button onclick="setToday()" type="button" style="margin:6px">امروز</button>' +
        '<input type="date" id="updateDate" style="width:90%;padding:6px;margin:6px 0"><br>' +
        '<div style="margin-top:8px"><button onclick="submitUpdate()" id="saveBtn">ذخیره</button> <button onclick="closeModal()" type="button">لغو</button></div>' +
        '</div></div>' +
        '<script>' +
        'let currentCode=null;' +
        'function openUpdate(code,name){ currentCode=code; document.getElementById("updateName").value=name||""; document.getElementById("updateField").value=""; document.getElementById("updateType").value="نو"; document.getElementById("updateDate").value=""; document.getElementById("modal").style.display="flex"; }' +
        'function closeModal(){ document.getElementById("modal").style.display="none"; }' +
        'function setToday(){ const d=new Date(); const j=window.jalaali.toJalaali(d); var s=j.jy+"-"+String(j.jm).padStart(2,"0")+"-"+String(j.jd).padStart(2,"0"); document.getElementById("updateDate").value=s; }' +
        'function submitUpdate(){ const name=document.getElementById("updateName").value||""; const field=document.getElementById("updateField").value||""; const type=document.getElementById("updateType").value||""; let date=document.getElementById("updateDate").value||""; if(!field||!type||!date){ alert("تمام فیلدها باید پر شوند"); return;} date=date.replace(/-/g,"/"); let itemText=(type==="استوک")?("استوک "+field):field; fetch("/update",{method:"POST",headers:{"Content-Type":"application/json"}, body:JSON.stringify({code:currentCode,name:name,item:itemText,date:date})}).then(r=>r.text()).then(msg=>{alert(msg); closeModal(); window.location.reload();});}' +
        '</script>' +
        '<p><a href="/">بازگشت</a></p>' +
        '</body></html>'
    );
});


// ================== UPDATE ========================
app.post('/update', (req,res)=>{
    const { code,name,item,date }=req.body||{};
    if(!date||!item) return res.status(400).send('تاریخ یا آیتم وارد نشده');

    const normDate = String(date).replace(/-/g,'/');

    try{
        const wb = xlsx.readFile(filteredPath);
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        const rows = xlsx.utils.sheet_to_json(sheet,{defval:''});

        rows.push({
            "کد پرسنلی": code || "",
            "نام کارمند": name || "",
            "نام کالا": item || "",
            "تاریخ تحویل": normDate
        });

        wb.Sheets[sheetName] = xlsx.utils.json_to_sheet(rows);
        xlsx.writeFile(wb, filteredPath);

        loadFilteredData();
        mergeBottleIntoMain();

        return res.send('با موفقیت ذخیره شد');
    }catch(e){
        return res.status(500).send('خطا در ذخیره‌سازی اکسل: '+e.message);
    }
});


app.listen(PORT,()=> console.log(`Server running at http://localhost:${PORT}`));


