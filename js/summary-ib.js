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


// fungsi untuk mencari kolom walau ada spasi / beda huruf
function findColumn(row, columnName){

for(let key in row){

if(key.toLowerCase().trim() === columnName.toLowerCase().trim()){
return row[key];
}

}

return "";

}


function processExcel(){

const fileInput = document.getElementById("excelFile").files[0];

if(!fileInput){
alert("Silakan upload file Excel terlebih dahulu");
return;
}

const reader = new FileReader();

reader.onload = function(e){

const data = new Uint8Array(e.target.result);

const workbook = XLSX.read(data,{type:"array"});

const sheetName = workbook.SheetNames[0];

const sheet = workbook.Sheets[sheetName];

const jsonData = XLSX.utils.sheet_to_json(sheet);

const tbody = document.querySelector("#resultTable tbody");

tbody.innerHTML = "";

processedData = [];

jsonData.forEach(row => {

const status = findColumn(row,"Status");

const result = {

DispatchStatus: status,
Status: status,
WO: findColumn(row,"No Wo Klien"),
Tanggal: findColumn(row,"Tanggal Kunjungan"),
Alamat: findColumn(row,"Alamat"),
ONT: findColumn(row,"ONT"),
STB: findColumn(row,"STB"),
Router: findColumn(row,"Router"),
Precon: precon,
ReportInstallation: getReportInstallation(row)

};

processedData.push(result);

const tr = document.createElement("tr");

tr.innerHTML = `
<td>${result.Status || ""}</td>
<td>${result.WO || ""}</td>
<td>${result.Tanggal || ""}</td>
<td>${result.Alamat || ""}</td>
<td>${result.ONT || ""}</td>
<td>${result.STB || ""}</td>
<td>${result.Router || ""}</td>
<td>${result.Precon || ""}</td>
<td>${result.ReportInstallation || ""}</td>
`;

tbody.appendChild(tr);

});

};

reader.readAsArrayBuffer(fileInput);

}


function downloadExcel(){

if(processedData.length == 0){
alert("Belum ada data yang diproses");
return;
}

const worksheet = XLSX.utils.json_to_sheet(processedData);

const workbook = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(workbook, worksheet, "Summary IB");

XLSX.writeFile(workbook, "summary_ib.xlsx");

}
