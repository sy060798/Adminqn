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
// PARSE REPORT (FULL FIX)
// =======================
function parseReport(report){

if(!report) return {newOnt:"",oldOnt:"",splacing:"",rfo:"",action:""};

let text = report.toString();
let lines = text.split(/\r?\n/);

// bersihin list
lines = lines.map(line =>
line.replace(/^\s*[\*\-\[\]\(\)]*/g,"")
.replace(/^\s*\d+\s*[\.\)]\s*/g,"")
.trim()
);

// =======================
// ONT
// =======================
let newOntMatch = text.match(/SN\s*(ONT|PERANGKAT)\s*BARU\s*:?\s*([A-Z0-9]+)/i);
let newOnt = newOntMatch ? newOntMatch[2] : "";

let oldOntMatch = text.match(/SN\s*(ONT|PERANGKAT)\s*LAMA\s*:?\s*([A-Z0-9]+)/i);
let oldOnt = oldOntMatch ? oldOntMatch[2] : "";

// =======================
// SPLICING
// =======================
let splacingMatch = text.match(/Sleeve\s*Protec\w*\s*:?[\s]*(\d+)/i);
let splacing = splacingMatch ? splacingMatch[1] : "";

// =======================
// RFO & ACTION
// =======================
let rfo = "";
let action = "";

for(let line of lines){
let lower = line.toLowerCase();

if(!rfo && (lower.startsWith("rfo") || lower.startsWith("problem"))){
rfo = line.replace(/(rfo|problem)\s*:/i,"").trim();
}

if(!action && (lower.startsWith("act") || lower.startsWith("action"))){
action = line.replace(/(act|action)\s*:/i,"").trim();
}
}

// =======================
// FALLBACK ACTION
// =======================
if(!action){
for(let line of lines){
let lower = line.toLowerCase();

if(
lower.includes("join") ||
lower.includes("splice") ||
lower.includes("sambung") ||
lower.includes("tarik")
){
action = line;

let num = line.match(/\d+/);
if(num) splacing = num[0];

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
break;
}
}
}

// =======================
// 🔥 SPLICING DARI RFO
// =======================
if(!splacing && rfo){
let num = rfo.match(/\d+/);
if(num) splacing = num[0];
}

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

// =======================
// 🔥 FIX ID & WO
// =======================
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
