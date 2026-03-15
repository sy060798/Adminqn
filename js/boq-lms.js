let resultData = [];

function updateStatus(text){
document.getElementById("statusText").innerText = text;
}

function updateProgress(percent){

const bar = document.getElementById("progressBar");

bar.style.width = percent + "%";

bar.innerText = percent + "%";

}

function processExcel(){

const file = document.getElementById("excelFile").files[0];

if(!file){
alert("Upload file Excel dulu");
return;
}

resultData = [];

document.querySelector("#resultTable tbody").innerHTML="";

updateStatus("Membaca file Excel...");

const reader = new FileReader();

reader.onload = function(e){

const data = new Uint8Array(e.target.result);

const workbook = XLSX.read(data,{type:'array'});

const sheet = workbook.Sheets[workbook.SheetNames[0]];

const rows = XLSX.utils.sheet_to_json(sheet,{header:1});

let project="";
let spk="";
let tanggal="";

rows.forEach(r=>{

let text = (r[0]||"").toString();

if(text.includes("NAMA PROJECT")) project = r[2];
if(text.includes("NO. SPK")) spk = r[2];
if(text.includes("TANGGAL")) tanggal = r[2];

});

let total = rows.length;

rows.forEach((r,i)=>{

let no = r[0];
let item = r[1];
let qty = r[3];

qty = parseFloat(qty);

if(!isNaN(no) && item && qty>0){

addRow(project,spk,tanggal,no,item,qty);

}

let percent = Math.round((i/total)*100);

updateProgress(percent);

});

updateProgress(100);

updateStatus("Selesai membaca Excel");

};

reader.readAsArrayBuffer(file);

}

function addRow(project,spk,tanggal,no,item,qty){

const tbody = document.querySelector("#resultTable tbody");

const tr = document.createElement("tr");

tr.innerHTML = `
<td>${project}</td>
<td>${spk}</td>
<td>${tanggal}</td>
<td>${no}</td>
<td>${item}</td>
<td>${qty}</td>
`;

tbody.appendChild(tr);

resultData.push({

Project:project,
SPK:spk,
Tanggal:tanggal,
No:no,
Item:item,
Qty:qty

});

}

function downloadExcel(){

if(resultData.length===0){

alert("Tidak ada data");

return;

}

const ws = XLSX.utils.json_to_sheet(resultData);

const wb = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(wb,ws,"BOQ");

XLSX.writeFile(wb,"boq_result.xlsx");

}
