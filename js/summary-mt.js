let processedData = [];


// ==========================
// PRECON
// ==========================

function getPrecon(row){

const preconMap = {

"Kabel Precon 35 Old": "PRECON - 35 M",
"Kabel Precon 50 Old": "PRECON - 50 M",
"Kabel Precon 75 Old": "PRECON - 75 M",
"Kabel Precon 100 Old": "PRECON - 100 M",
"Kabel Precon 125 Old": "PRECON - 125 M",
"Kabel Precon 150 Old": "PRECON - 150 M",
"Kabel Precon 175 Old": "PRECON - 175 M",
"Kabel Precon 200 Old": "PRECON - 200 M",
"Kabel Precon 225 Old": "PRECON - 225 M",
"Kabel Precon 250 Old": "PRECON - 250 M"

};

let result = [];

for(let key in preconMap){

if(row[key] == 1 || row[key] == "1"){
result.push(preconMap[key]);
}

}

return result.join(", ");

}


// ==========================
// AMBIL KOLOM FLEKSIBEL
// ==========================

function getColumn(row,name){

for(let key in row){

if(key.toLowerCase().trim() === name.toLowerCase()){
return row[key];
}

}

return "";

}


// ==========================
// STATUS DISPATCH
// ==========================

function getDispatchStatus(row){

for(let key in row){

let col = key.toLowerCase();

if(col.includes("dispatch")){
return row[key];
}

}

return "";

}


// ==========================
// AMBIL REPORT
// ==========================

function getReportInstallation(row){

for(let key in row){

let col = key.toLowerCase();

if(col.includes("report")){
return row[key];
}

}

return "";

}


// ==========================
// PARSE REPORT TEKNISI
// ==========================

function parseReport(report){

if(!report) return {newOnt:"",splacing:"",rfo:"",action:""};

let text = report.toString();


// bersihkan simbol aneh
text = text.replace(/[\*\[\]\(\)#]/g,"");

let lines = text.split(/\r?\n/);

let action = "";
let rfo = "";


// ======================
// FORMAT RFO / ACT
// ======================

let rfoMatch = text.match(/rfo\s*[:\-]\s*(.*)/i);
if(rfoMatch){
rfo = rfoMatch[1].trim();
}

let actMatch = text.match(/act\s*[:\-]\s*(.*)/i);
if(actMatch){
action = actMatch[1].trim();
}


// ======================
// FORMAT REMARK
// ======================

if(!rfo){

for(let i=0;i<lines.length;i++){

let lower = lines[i].toLowerCase();

if(lower.includes("remak") || lower.includes("remark")){

if(lines[i+1]){
rfo = lines[i+1].trim();
}

break;

}

}

}


// ======================
// FORMAT KEYWORD ACTION
// ======================

if(!action){

const actionKeywords = [
"joint",
"join",
"splice",
"splicing",
"tarik",
"ganti",
"repair",
"perbaikan",
"sambung"
];

for(let line of lines){

let lower = line.toLowerCase();

for(let key of actionKeywords){

if(lower.includes(key)){
action = line.trim();
break;
}

}

if(action) break;

}

}


// ======================
// HITUNG SPLICING
// ======================

let splacing = "";

if(action){

let numMatch = action.match(/\d+/);

if(numMatch){

splacing = numMatch[0];

}else{

let lower = action.toLowerCase();

if(lower.includes("satu")) splacing="1";
if(lower.includes("dua")) splacing="2";
if(lower.includes("tiga")) splacing="3";
if(lower.includes("empat")) splacing="4";

}

}


// NEW ONT DIKOSONGKAN
let newOnt = "";

return {newOnt,splacing,rfo,action};

}


// ==========================
// PROCESS EXCEL
// ==========================

function processExcel(){

const file = document.getElementById("excelFile").files[0];

if(!file){
alert("Upload Excel terlebih dahulu");
return;
}

const reader = new FileReader();

reader.onload = function(e){

const data = new Uint8Array(e.target.result);

const workbook = XLSX.read(data,{type:"array"});

const sheet = workbook.Sheets[workbook.SheetNames[0]];

const jsonData = XLSX.utils.sheet_to_json(sheet,{defval:""});

const tbody = document.querySelector("#resultTable tbody");

tbody.innerHTML = "";

processedData = [];

jsonData.forEach(row=>{

let dispatchStatus = getDispatchStatus(row);

dispatchStatus = String(dispatchStatus).trim().toLowerCase();

if(dispatchStatus !== "done") return;

let report = getReportInstallation(row);

let parsed = parseReport(report);

const result = {

dispatch:"Done",
status:"Done",
wo:getColumn(row,"No Wo Klien"),
id:getColumn(row,"Cust ID Klien"),
tanggal:getColumn(row,"Tanggal Kunjungan"),
alamat:getColumn(row,"Alamat"),

new_ont:parsed.newOnt,
splacing:parsed.splacing,
rfo:parsed.rfo,
action:parsed.action,

precon:getPrecon(row),
report:report

};

processedData.push(result);


// tampilkan tabel

const tr = document.createElement("tr");

tr.innerHTML = `

<td>${result.dispatch}</td>
<td>${result.status}</td>
<td>${result.wo}</td>
<td>${result.id}</td>
<td>${result.tanggal}</td>
<td>${result.alamat}</td>
<td>${result.new_ont}</td>
<td>${result.splacing}</td>
<td>${result.rfo}</td>
<td>${result.action}</td>
<td>${result.precon}</td>
<td style="max-width:600px;word-break:break-word;white-space:pre-line;">${result.report}</td>

`;

tbody.appendChild(tr);

});

};

reader.readAsArrayBuffer(file);

}


// ==========================
// DOWNLOAD EXCEL
// ==========================

function downloadExcel(){

if(processedData.length === 0){
alert("Belum ada data untuk didownload");
return;
}

const worksheet = XLSX.utils.json_to_sheet(processedData);

const workbook = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(workbook,worksheet,"Summary MT");

XLSX.writeFile(workbook,"summary_mt_done.xlsx");

}
