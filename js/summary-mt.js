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
function getDispatchStatus(row){
return getColumn(row, ["dispatch"]);
}

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
// 🔥 PARSE REPORT FINAL FIX
// =======================
function parseReport(report){

if(!report) return {newOnt:"",oldOnt:"",splacing:"",rfo:"",action:""};

let rawText = report.toString();

// =======================
// CLEAN TEXT (AMAN)
// =======================
let clean = rawText
.replace(/\r/g,"\n")
.replace(/[\*\-\•]/g,"")
.replace(/\;/g,":")
.trim();

let lines = clean
.split("\n")
.map(l => l.trim())
.filter(l => l !== "");

// =======================
// 🔥 RFO & ACTION (FIX UTAMA)
// =======================
let rfo = "";
let action = "";

for(let i=0; i<lines.length; i++){

let l = lines[i].toLowerCase();

// ===== RFO =====
if(!rfo && /^(rfo|problem|gangguan|remak)\b/.test(l)){

if(lines[i+1]){
rfo = lines[i+1].trim();
}

if(!rfo){
rfo = l.replace(/^(rfo|problem|gangguan|remak)\s*:?/,"").trim();
}

continue;
}

// ===== ACTION =====
if(!action && /^(act|action|tindakan)\b/.test(l)){

if(lines[i+1]){
action = lines[i+1].trim();
}

if(!action){
action = l.replace(/^(act|action|tindakan)\s*:?/,"").trim();
}

continue;
}

}

// =======================
// 🔥 FALLBACK
// =======================
lines.forEach(line=>{

let l = line.toLowerCase();

if(!rfo && (
l.includes("loss") ||
l.includes("cut") ||
l.includes("putus") ||
l.includes("gigit") ||
l.includes("tikus")
)){
rfo = line;
}

if(!action && (
l.includes("join") ||
l.includes("splice") ||
l.includes("sambung") ||
l.includes("ganti") ||
l.includes("pergantian")
)){
action = line;
}

});

// =======================
// 🔥 AUTO FIX KETUKER
// =======================
if(rfo && action){

let r = rfo.toLowerCase();
let a = action.toLowerCase();

if(
r.includes("join") ||
r.includes("sambung")
){
let temp = rfo;
rfo = action;
action = temp;
}

}

// =======================
// 🔥 SPLICING (ONLY ACTION)
// =======================
let splacing = 0;

if(action){

let act = action.toLowerCase();

// PRIORITAS titik
let titik = act.match(/(\d+)\s*titik/);
if(titik){
splacing = parseInt(titik[1]);
}else{

// kalau ada "dan" → 1
if(/ dan | & | \+ /.test(act)){
if(/join|rejoin|splice|sambung/.test(act)){
splacing = 1;
}
}else{

let matches = act.match(/(\d+)?\s*(join|rejoin|splice|sambung)/gi);

if(matches){
matches.forEach(m=>{
let num = m.match(/\d+/);
splacing += num ? parseInt(num[0]) : 1;
});
}

}

}

}

// =======================
// 🔥 ONT FLEX
// =======================
let newOnt = "";
let oldOnt = "";

for(let i=0;i<lines.length;i++){

let l = lines[i].toLowerCase();

if(l.includes("lama")){
let next = lines[i+1] || "";
let sn = next.match(/([a-z0-9]{8,})/i);
if(sn) oldOnt = sn[1];
}

if(l.includes("baru")){
let next = lines[i+1] || "";
let sn = next.match(/([a-z0-9]{8,})/i);
if(sn) newOnt = sn[1];
}

}

// fallback SN global
let allSN = [...clean.matchAll(/sn\s*[:\-]?\s*([a-z0-9]+)/gi)].map(m=>m[1]);

if(!oldOnt && allSN.length >= 2){
oldOnt = allSN[0];
newOnt = allSN[1];
}

if(!newOnt && allSN.length === 1){
newOnt = allSN[0];
}

// =======================
return {
newOnt,
oldOnt,
splacing: splacing || "",
rfo,
action
};

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

// =======================
// RENDER TABLE
// =======================
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
