let resultData=[];

function updateStatus(text){
document.getElementById("statusText").innerText=text;
}

function updateProgress(percent){
const bar=document.getElementById("progressBar");
bar.style.width=percent+"%";
bar.innerText=percent+"%";
}

async function processExcel(){

const files=document.getElementById("excelFile").files;

if(files.length===0){
alert("Upload Excel dulu");
return;
}

resultData=[];
document.querySelector("#resultTable tbody").innerHTML="";

for(let f=0;f<files.length;f++){

await readExcel(files[f]);

let percent=Math.round(((f+1)/files.length)*100);

updateProgress(percent);

}

updateStatus("Selesai membaca "+files.length+" file");

}

function readExcel(file){

return new Promise((resolve)=>{

const reader=new FileReader();

reader.onload=function(e){

const data=new Uint8Array(e.target.result);

const workbook=XLSX.read(data,{type:'array'});

const sheet=workbook.Sheets[workbook.SheetNames[0]];

const rows=XLSX.utils.sheet_to_json(sheet,{header:1});

let project="";
let spk="";
let tanggal="";

for(let r of rows){

let col0=(r[0]||"").toString();

if(col0.includes("NAMA PROJECT")) project=r[2];
if(col0.includes("NO. SPK")) spk=r[2];
if(col0.includes("TANGGAL")) tanggal=r[2];

}

for(let r of rows){

let no=r[0];
let item=r[1];
let qty=r[3];

if(typeof no==="number" && item){

qty=parseFloat(qty);

if(!isNaN(qty) && qty>0){

addRow(project,spk,tanggal,no,item,qty);

}

}

}

resolve();

};

reader.readAsArrayBuffer(file);

});

}

function addRow(project,spk,tanggal,no,item,qty){

const tbody=document.querySelector("#resultTable tbody");

const tr=document.createElement("tr");

tr.innerHTML=`
<td>${project||""}</td>
<td>${spk||""}</td>
<td>${tanggal||""}</td>
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

const ws=XLSX.utils.json_to_sheet(resultData);

const wb=XLSX.utils.book_new();

XLSX.utils.book_append_sheet(wb,ws,"BOQ");

XLSX.writeFile(wb,"BOQ_LMS_RESULT.xlsx");

}
