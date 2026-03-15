let resultData=[];

const MATERIAL_LIST=[

"Kabel Drop 2 Core",
"Kabel Udara ADSS span 100 2 core",
"Kabel Udara ADSS span 100 12 core",
"Kabel Udara ADSS span 100 8 core",
"Kabel Udara ADSS span 100 48 core",
"Joint Clouser Dome 12 Core",
"Joint Closer Dome 96 Core",
"Pipa PVC 3/4 inch",
"PLC Splitter 1*2",
"SPLITER 1:16",
"Fixing Slack",
"Klem Pipa Conduit",
"Flexible Pipe",
"Pipa subduct 28/32",
"Tiang 7 meter",
"Tiang 9 meter",
"ONT",
"ODP",
"SPLICING + OTDR",
"Transport LMS"

];

function updateProgress(done,total){

let percent=Math.round((done/total)*100);

const bar=document.getElementById("progressBar");

bar.style.width=percent+"%";
bar.innerText=percent+"%";

}

function updateStatus(text){

document.getElementById("statusText").innerText=text;

}

async function processPDF(){

const files=document.getElementById("pdfFiles").files;

if(files.length===0){

alert("Upload PDF dulu");
return;

}

for(let i=0;i<files.length;i++){

updateStatus("Membaca "+files[i].name);

await readPDF(files[i]);

updateProgress(i+1,files.length);

}

updateStatus("Selesai");

}

async function readPDF(file){

const project=file.name.replace(".pdf","");

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

findMaterial(text,project);

}

function findMaterial(text,project){

MATERIAL_LIST.forEach(item=>{

const regex=new RegExp(item+"\\s+(\\d+)","i");

const match=text.match(regex);

if(match){

let qty=parseInt(match[1]);

if(qty>0){

addRow(project,item,qty);

}

}

});

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

MATERIAL_LIST.forEach(item=>{

if(r.Item && r.Item.toLowerCase().includes(item.toLowerCase())){

if(r.Qty>0){

addRow(r.Project||"Excel",item,r.Qty);

}

}

});

});

updateStatus("Excel berhasil dibaca");

};

reader.readAsArrayBuffer(file);

}

function addRow(project,item,qty){

const tbody=document.querySelector("#resultTable tbody");

const tr=document.createElement("tr");

tr.innerHTML=`
<td>${project}</td>
<td>${item}</td>
<td>${qty}</td>
`;

tbody.appendChild(tr);

resultData.push({

Project:project,
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
