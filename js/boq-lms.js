// simpan hasil
let resultData = [];

// status text
function updateStatus(text){
const status=document.getElementById("statusText");
if(status) status.innerText=text;
}

// progress bar
function updateProgress(done,total){

let percent=Math.round((done/total)*100);

let bar=document.getElementById("progressBar");

if(bar){
bar.style.width=percent+"%";
bar.innerText=percent+"%";
}

}

// tombol proses PDF
async function processPDF(){

try{

const input=document.getElementById("pdfFiles");

if(!input || input.files.length===0){

alert("Upload PDF dulu");
return;

}

const files=input.files;

resultData=[];

document.querySelector("#resultTable tbody").innerHTML="";

for(let i=0;i<files.length;i++){

updateStatus("Membaca "+files[i].name);

await readPDF(files[i]);

updateProgress(i+1,files.length);

}

updateStatus("Selesai membaca PDF");

}catch(err){

console.error(err);
alert("Terjadi error membaca PDF");

}

}

// membaca PDF
async function readPDF(file){

const project=file.name.replace(".pdf","");

const buffer=await file.arrayBuffer();

const pdf=await pdfjsLib.getDocument({data:buffer}).promise;

let words=[];

for(let page=1;page<=pdf.numPages;page++){

let p=await pdf.getPage(page);

let txt=await p.getTextContent();

txt.items.forEach(t=>{

words.push(t.str.trim());

});

}

parseTable(words,project);

}

// membaca tabel BOQ
function parseTable(words,project){

for(let i=0;i<words.length;i++){

let no=parseInt(words[i]);

// jika angka berarti kemungkinan nomor item
if(!isNaN(no)){

let item=words[i+1];

let qty=parseInt(words[i+3]);

if(item && !isNaN(qty) && qty>0){

addRow(project,item,qty);

}

}

}

}

// tambah row tabel
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

// download excel
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
