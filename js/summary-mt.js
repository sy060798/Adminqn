let processedData = [];

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


// ambil kolom fleksibel
function getColumn(row,name){

for(let key in row){

if(key.toLowerCase().trim() === name.toLowerCase()){
return row[key];
}

}

return "";

}


// ambil status dispatch
function getDispatchStatus(row){

for(let key in row){

let col = key.toLowerCase();

if(col.includes("dispatch")){
return row[key];
}

}

return "";

}


// ambil report installation
function getReportInstallation(row){

for(let key in row){

let col = key.toLowerCase();

if(col.includes("report")){
return row[key];
}

}

return "";

}


// ======================
// PARSE DATA DARI REPORT
// ======================

function parseReport(report){

if(!report) return {newOnt:"",splacing:"",rfo:"",action:""};

let text = report.toString();

let lines = text.split(/\n|\r/);

// NEW ONT
let newOntMatch = text.match(/SN\s*ONT\s*BARU\s*:?\s*(.*)/i);
let newOnt = newOntMatch ? newOntMatch[1].trim() : "";

// SPLICING (Sleeve Protection)
let splacingMatch = text.match(/Sleeve\s*Protec\w*\s*:?[\s]*(\d+)/i);
let splacing = splacingMatch ? splacingMatch[1] : "";

// RFO
let rfo = "";
for(let line of lines){
if(line.trim().toLowerCase().startsWith("rfo")){
rfo = line.trim();
break;
}
}

// ACTION
let action = "";
for(let line of lines){
if(line.trim().toLowerCase().startsWith("act")){
action = line.trim();
break;
}
}

return {newOnt,splacing,rfo,action};

}


// ======================
// PROCESS EXCEL
// ======================

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

const jsonData = XLSX.utils.sheet_to_json(sheet);

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


// ======================
// TAMPILKAN KE TABLE
// ======================

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
<td style="max-width:600px;word-break:break-word;">${result.report || ""}</td>

`;

tbody.appendChild(tr);

});

};

reader.readAsArrayBuffer(file);

}


// ======================
// DOWNLOAD EXCEL
// ======================

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
