let resultData=[];

function updateStatus(text){

document.getElementById("statusText").innerText=text;

}

function updateProgress(done,total){

let percent=Math.round((done/total)*100);

const bar=document.getElementById("progressBar");

bar.style.width=percent+"%";
bar.innerText=percent+"%";

}

function sleep(ms){

return new Promise(resolve=>setTimeout(resolve,ms));

}



async function processPDFs(){

const files=document.getElementById("pdfFiles").files;

if(files.length===0){

alert("Upload PDF dulu");

return;

}

const total=files.length;

for(let i=0;i<total;i++){

updateStatus("Memproses PDF "+(i+1)+" dari "+total);

await readPDF(files[i]);

updateProgress(i+1,total);

await sleep(200);

}

updateStatus("Selesai membaca PDF");

}



async function readPDF(file){

const buffer=await file.arrayBuffer();

const pdf=await pdfjsLib.getDocument({data:buffer}).promise;

let text="";

for(let pageNum=1;pageNum<=pdf.numPages;pageNum++){

const page=await pdf.getPage(pageNum);

const content=await page.getTextContent();

content.items.forEach(item=>{

text+=item.str+" ";

});

}

if(text.trim().length<50){

await runOCR(pdf);

}else{

extractData(text);

}

}



async function runOCR(pdf){

for(let pageNum=1;pageNum<=pdf.numPages;pageNum++){

updateStatus("OCR membaca halaman "+pageNum);

const page=await pdf.getPage(pageNum);

const viewport=page.getViewport({scale:2});

const canvas=document.createElement("canvas");

const context=canvas.getContext("2d");

canvas.height=viewport.height;
canvas.width=viewport.width;

await page.render({

canvasContext:context,
viewport:viewport

}).promise;

const result=await Tesseract.recognize(canvas,"eng");

extractData(result.data.text);

}

}



function processExcel(){

const file=document.getElementById("excelFile").files[0];

if(!file){

alert("Upload Excel dulu");

return;

}

const reader=new FileReader();

reader.onload=function(e){

const data=new Uint8Array(e.target.result);

const workbook=XLSX.read(data,{type:'array'});

const sheet=workbook.Sheets[workbook.SheetNames[0]];

const rows=XLSX.utils.sheet_to_json(sheet);

rows.forEach(r=>{

if(r.Qty>0){

addRow(r.Project,r["No WO"],r.Tanggal,r.No,r.Item,r.Qty);

}

});

updateStatus("Excel berhasil dibaca");

};

reader.readAsArrayBuffer(file);

}



function extractData(text){

let project="";
let wo="";
let tanggal="";

const projectMatch=text.match(/project\s*[:\-]?\s*(.*?)\n/i);
if(projectMatch) project=projectMatch[1];

const woMatch=text.match(/wo\s*[:\-]?\s*(\S+)/i);
if(woMatch) wo=woMatch[1];

const dateMatch=text.match(/tanggal\s*[:\-]?\s*(\S+)/i);
if(dateMatch) tanggal=dateMatch[1];

const itemRegex=/(\d+)\s+([A-Za-z0-9\s\(\)\-\/]+?)\s+(\d+)/g;

let match;

while((match=itemRegex.exec(text))!==null){

let no=match[1];
let item=match[2].trim();
let qty=parseInt(match[3]);

if(qty>0){

addRow(project,wo,tanggal,no,item,qty);

}

}

}



function addRow(project,wo,tanggal,no,item,qty){

const tbody=document.querySelector("#resultTable tbody");

const tr=document.createElement("tr");

tr.innerHTML=`

<td>${project||""}</td>
<td>${wo||""}</td>
<td>${tanggal||""}</td>
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

XLSX.writeFile(wb,"boq_lms.xlsx");

}
