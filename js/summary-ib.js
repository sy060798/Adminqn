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

function processExcel(){

const fileInput = document.getElementById("excelFile").files[0];

if(!fileInput){
alert("Upload file Excel dulu");
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

tbody.innerHTML="";

processedData = [];

jsonData.forEach(row=>{

const precon = getPrecon(row);

const result = {

Status: row["Status"],
WO: row["No Wo Klien"],
Tanggal: row["Tanggal Kunjungan"],
Alamat: row["Alamat"],
ONT: row["ONT"],
STB: row["STB"],
Router: row["Router"],
Precon: precon

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
`;

tbody.appendChild(tr);

});

};

reader.readAsArrayBuffer(fileInput);

}

function downloadExcel(){

if(processedData.length == 0){
alert("Belum ada data");
return;
}

const worksheet = XLSX.utils.json_to_sheet(processedData);

const workbook = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(workbook,worksheet,"Summary IB");

XLSX.writeFile(workbook,"summary_ib.xlsx");

}
