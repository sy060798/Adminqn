let processedData = [];

// =======================
// FLEXIBLE COLUMN
// =======================
function getColumn(row, keywords){
for(let key in row){
let col = key.toLowerCase().trim();

for(let word of keywords){
if(col.includes(word.toLowerCase())){
return row[key];
}
}
}
return "";
}

// =======================
// DISPATCH
// =======================
function getDispatchStatus(row){
return getColumn(row, ["dispatch"]);
}

// =======================
// REPORT
// =======================
function getReportInstallation(row){
return getColumn(row, ["report"]);
}

// =======================
// PRECON
// =======================
function getPrecon(row){

const preconMap = {
"kabel precon 35": "PRECON - 35 M",
"kabel precon 50": "PRECON - 50 M",
"kabel precon 75": "PRECON - 75 M",
"kabel precon 80": "PRECON - 80 M",
"kabel precon 100": "PRECON - 100 M",
"kabel precon 125": "PRECON - 125 M",
"kabel precon 150": "PRECON - 150 M",
"kabel precon 175": "PRECON - 175 M",
"kabel precon 200": "PRECON - 200 M",
"kabel precon 225": "PRECON - 225 M",
"kabel precon 250": "PRECON - 250 M"
};

let result=[];

for(let key in row){
let col = key.toLowerCase();

for(let p in preconMap){
if(col.includes(p) && (row[key]==1 || row[key]=="1")){
result.push(preconMap[p]);
}
}
}

return result.join(", ");
}

// =======================
// 🔥 CLEAN TEXT
// =======================
function cleanText(text){
return text
.replace(/\*/g,"")
.replace(/[•]/g,"")
.replace(/\t/g," ")
.replace(/\r/g,"")
.replace(/\s+/g," ")
.trim();
}

// =======================
// 🔥 HITUNG SPLICING CERDAS
// =======================
function countSplicing(text){

let count = 0;

// keyword utama
const keywords = ["join","rejoin","splice","sambung","splicing"];

// hitung keyword
keywords.forEach(k=>{
let match = text.match(new RegExp(k,"gi"));
if(match) count += match.length;
});

// ambil angka kalau ada
let num = text.match(/(\d+)/);
if(num){
count = parseInt(num[1]);
}

// NORMALISASI:
// kalau ada "dan / & / +" tetap 1
if(
text.includes(" dan ") ||
text.includes("&") ||
text.includes("+")
){
count = Math.max(1,count);
}

return count || "";
}

// =======================
// 🔥 PARSE REPORT SUPER
// =======================
function parseReport(report){

if(!report) return {newOnt:"",oldOnt:"",splacing:"",rfo:"",action:""};

let raw = report.toString();
let clean = cleanText(raw);

// pecah baris fleksibel
let lines = raw.split(/\r?\n/).map(l=>cleanText(l)).filter(l=>l);

// =======================
// 🔥 RFO & ACTION FLEXIBLE
// =======================
let rfo = "";
let action = "";

for(let i=0;i<lines.length;i++){

let l = lines[i].toLowerCase();

// RFO
if(!rfo && (l.startsWith("rfo") || l.includes("rfo"))){
rfo = lines[i].replace(/rfo\s*[:;\-]?\s*/i,"");
continue;
}

// ACTION
if(!action && (l.startsWith("act") || l.includes("act") || l.includes("action"))){
action = lines[i].replace(/(act|action)\s*[:;\-]?\s*/i,"");
continue;
}

}

// =======================
// 🔥 FORMAT KHUSUS (RFO / ACT DI BARIS BERIKUTNYA)
// =======================
for(let i=0;i<lines.length;i++){

let l = lines[i].toLowerCase();

if(l === "rfo" && lines[i+1]){
rfo = lines[i+1];
}

if((l === "act" || l === "action") && lines[i+1]){
action = lines[i+1];
}

}

// =======================
// 🔥 FALLBACK ACTION
// =======================
if(!action){
for(let line of lines){
let lower = line.toLowerCase();

if(
lower.includes("join") ||
lower.includes("splice") ||
lower.includes("sambung") ||
lower.includes("tarik") ||
lower.includes("ganti")
){
action = line;
break;
}
}
}

// =======================
// 🔥 REMAK FIX
// =======================
if(!rfo){
for(let i=0;i<lines.length;i++){
if(lines[i].toLowerCase().includes("remak")){
if(lines[i+1]) rfo = lines[i+1];
}
}
}

// =======================
// 🔥 SPLICING DARI ACTION
// =======================
let splacing = countSplicing(action);

// =======================
// 🔥 ONT SUPER STRICT
// =======================
let newOnt = "";
let oldOnt = "";

function isValidSN(sn){

if(!sn) return false;

sn = sn.toUpperCase().trim();

// ZTE
if(sn.startsWith("ZTE") && sn.length >= 10){
return true;
}

// HUAWEI (16 char)
if(/^[A-Z0-9]{16}$/.test(sn)){
return true;
}

return false;
}

// ambil semua SN
let allSN = [...clean.matchAll(/sn\s*[:\-]?\s*([a-z0-9]+)/gi)]
.map(m => m[1].toUpperCase())
.filter(sn => isValidSN(sn));

// PRIORITAS LABEL
for(let i=0;i<lines.length;i++){

let l = lines[i].toLowerCase();

if(l.includes("lama")){
let next = lines[i+1] || "";
let sn = next.match(/([a-z0-9]+)/i);
if(sn && isValidSN(sn[1])){
oldOnt = sn[1].toUpperCase();
}
}

if(l.includes("baru")){
let next = lines[i+1] || "";
let sn = next.match(/([a-z0-9]+)/i);
if(sn && isValidSN(sn[1])){
newOnt = sn[1].toUpperCase();
}
}

}

// fallback
if(!oldOnt && !newOnt){
if(allSN.length >= 2){
oldOnt = allSN[0];
newOnt = allSN[1];
}
}

if(!newOnt && allSN.length === 1){
newOnt = allSN[0];
}

// =======================
return {newOnt,oldOnt,splacing,rfo,action};

}

// =======================
// PROCESS EXCEL
// =======================
function processExcel(){

const file = document.getElementById("excelFile").files[0];

if(!file){
alert("Upload Excel dulu!");
return;
}

const reader = new FileReader();

reader.onload = function(e){

const data = new Uint8Array(e.target.result);
const workbook = XLSX.read(data, {type:"array"});
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const json = XLSX.utils.sheet_to_json(sheet,{defval:""});

const tbody = document.querySelector("#resultTable tbody");

tbody.innerHTML="";
processedData=[];

json.forEach(row=>{

let dispatch = getDispatchStatus(row).toLowerCase();

if(dispatch !== "done") return;

let report = getReportInstallation(row);
let parsed = parseReport(report);

const result = {

dispatch:"Done",
status:"Done",

id: getColumn(row, ["cust id klien"]) 
|| getColumn(row, ["customer id"]) 
|| getColumn(row, ["id"]),

wo: getColumn(row, ["no wo klien"]) 
|| getColumn(row, ["wo number"]) 
|| getColumn(row, ["wo"]),

customer: getColumn(row, ["customer name"]) 
|| getColumn(row, ["nama"]),

tanggal: getColumn(row, ["tanggal kunjungan"]) 
|| getColumn(row, ["tanggal"]) 
|| getColumn(row, ["date"]),

alamat: getColumn(row, ["alamat"]),
cabang: getColumn(row, ["cabang"]),

new_ont: parsed.newOnt,
old_ont: parsed.oldOnt,
splacing: parsed.splacing,
rfo: parsed.rfo,
action: parsed.action,

precon: getPrecon(row),
report: report
};

processedData.push(result);

// render
const tr = document.createElement("tr");

tr.innerHTML = `
<td>${result.dispatch}</td>
<td>${result.status}</td>
<td>${result.id}</td>
<td>${result.wo}</td>
<td>${result.customer}</td>
<td>${result.tanggal}</td>
<td>${result.alamat}</td>
<td>${result.cabang}</td>
<td>${result.new_ont}</td>
<td>${result.old_ont}</td>
<td>${result.splacing}</td>
<td>${result.rfo}</td>
<td>${result.action}</td>
<td>${result.precon}</td>
<td style="white-space:pre-line">${result.report}</td>
`;

tbody.appendChild(tr);

});

};

reader.readAsArrayBuffer(file);
}

// =======================
// DOWNLOAD
// =======================
function downloadExcel(){

if(processedData.length===0){
alert("Tidak ada data");
return;
}

const ws = XLSX.utils.json_to_sheet(processedData);
const wb = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(wb, ws, "Summary MT");

XLSX.writeFile(wb, "summary_mt.xlsx");
}
