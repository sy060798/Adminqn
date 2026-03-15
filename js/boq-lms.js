let resultData = [];

function updateStatus(text){
document.getElementById("statusText").innerText=text;
}

function updateProgress(done,total){

let percent=Math.round((done/total)*100);

const bar=document.getElementById("progressBar");

bar.style.width=percent+"%";
bar.innerText=percent+"%";

}

function processExcel(){

const file=document.getElementById("excelFile").files[0];

if(!file){
alert("Upload Excel dulu");
return;
}

resultData=[];
document.querySelector("#resultTable tbody").innerHTML="";

updateStatus("Membaca Excel...");

const reader=new FileReader();

reader.onload=function(e){

const data=new Uint8Array(e.target.result);

const workbook=XLSX.read(data,{type:'array'});

const sheet=workbook.Sheets[workbook.SheetNames[0]];

const rows=XLSX.utils.sheet_to_json(sheet,{header:1});

let project="";
let wo="";
let tanggal="";

rows.forEach(r=>{

if(String(r[0]).includes("NAMA PROJECT")) project=r[2];
if(String(r[0]).includes("NO. SPK")) wo=r[2];
if(String(r[0]).includes("TANGGAL")) tanggal=r[2];

});

let total=rows.length;

rows.forEach((r,index)=>{

let no=r[0];
let item=r[1];
let qty=r[3];

qty=parseFloat(qty);

if(!isNaN(no) && qty>0){

addRow(project,wo,tanggal,no,item,qty);

}

updateProgress(index+1,total);

});

updateStatus("Selesai membaca Excel");

};

reader.readAsArrayBuffer(file);

}

function addRow(project,wo,tanggal,no,item,qty){

const tbody=document.querySelector("#resultTable tbody");

const tr=document.createElement("tr");

tr.innerHTML=`

<td>${project}</td>
<td>${wo}</td>
<td>${tanggal}</td>
<td>${no}</td>
<td>${item}</td>
<td>${qty}</td>

`;

tbody.appendChild(tr);

resultData.push({

Project:project,
WO:wo,
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

const ws=XLSX.utils.json_to_sheet(resultData);

const wb=XLSX.utils.book_new();

XLSX.utils.book_append_sheet(wb,ws,"BOQ");

XLSX.writeFile(wb,"boq_lms_result.xlsx");

}
