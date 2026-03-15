let resultData = [];

const MATERIAL_LIST = [

"Jasa Perbaikan ODP Pedestal (INC COR)",
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
"SPLICING + OTDR"

];

function updateProgress(done,total){

let percent=Math.round((done/total)*100);

let bar=document.getElementById("progressBar");

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

resultData=[];
document.querySelector("#resultTable tbody").innerHTML="";

for(let i=0;i<files.length;i++){

updateStatus("Membaca "+files[i].name);

await readPDF(files[i]);

updateProgress(i+1,files.length);

}

updateStatus("Selesai membaca PDF");

}

async function readPDF(file){

const project=file.name.replace(".pdf","");

const buffer=await file.arrayBuffer();

const pdf=await pdfjsLib.getDocument({data:buffer}).promise;

let text="";

for(let page=1;page<=pdf.numPages;page++){

let pageData=await pdf.getPage(page);

let txt=await pageData.getTextContent();

txt.items.forEach(t=>{
text+=t.str+" ";
});

}

findMaterial(text,project);

}

function findMaterial(text,project){

let cleanText=text.replace(/\s+/g," ").toLowerCase();

MATERIAL_LIST.forEach(item=>{

let itemLower=item.toLowerCase();

if(cleanText.includes(itemLower)){

let index=cleanText.indexOf(itemLower);

let area=cleanText.substring(index,index+80);

let numbers=area.match(/\d+/g);

if(numbers){

let qty=parseInt(numbers[numbers.length-1]);

if(qty>0){

addRow(project,item,qty);

}

}

}

});

}

function addRow(project,item,qty){

let tbody=document.querySelector("#resultTable tbody");

let tr=document.createElement("tr");

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

let ws=XLSX.utils.json_to_sheet(resultData);

let wb=XLSX.utils.book_new();

XLSX.utils.book_append_sheet(wb,ws,"BOQ");

XLSX.writeFile(wb,"BOQ_LMS_RESULT.xlsx");

}
