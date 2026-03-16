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

if(row[key] == 1){
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


// ambil kolom STATUS BY DISPATCH
function getDispatchStatus(row){

for(let key in row){

let col = key.toLowerCase();

if(col.includes("dispatch")){
return row[key];
}

}

return "";

}


// ambil report
function getReportInstallation(row){

for(let key in row){

let col = key.toLowerCase();

if(col.includes("report")){
return row[key];
}

}

return "";

}


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

// NORMALISASI TEXT
dispatchStatus = String(dispatchStatus).trim().toLowerCase();

// FILTER DONE SAJA
if(dispatchStatus !== "done"){
return;
}

const result = {

dispatch: "Done",
status: "Done",
wo: getColumn(row,"No Wo Klien"),
ID: getColumn(row,"Cust ID Klien"),    
tanggal: getColumn(row,"Tanggal Kunjungan"),
alamat: getColumn(row,"Alamat"),
ont: getColumn(row,"ONT"),
stb: getColumn(row,"STB"),
router: getColumn(row,"Router"),
precon: getPrecon(row),
report: getReportInstallation(row)

};

processedData.push(result);

const tr = document.createElement("tr");

tr.innerHTML = `

<td>${result.dispatch}</td>
<td>${result.status}</td>
<td>${result.wo}</td>
<td>${result.Cust ID Klien}</td>
<td>${result.tanggal}</td>
<td>${result.alamat}</td>
<td>${result.ont}</td>
<td>${result.stb}</td>
<td>${result.router}</td>
<td>${result.precon}</td>
<td style="max-width:600px;word-break:break-word;">
${result.report || ""}
</td>

`;

tbody.appendChild(tr);

});

};

reader.readAsArrayBuffer(file);

}


function downloadExcel(){

if(processedData.length == 0){
alert("Belum ada data");
return;
}

const worksheet = XLSX.utils.json_to_sheet(processedData);

const workbook = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(workbook,worksheet,"Summary MT");

XLSX.writeFile(workbook,"summary_mt_done.xlsx");

}
