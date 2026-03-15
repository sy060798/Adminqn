let resultData=[];

async function processPDFs(){

const files=document.getElementById("pdfFiles").files;

if(files.length===0){

alert("Upload PDF terlebih dahulu");

return;

}

const tbody=document.querySelector("#resultTable tbody");

tbody.innerHTML="";

resultData=[];

for(let file of files){

await readPDF(file);

}

}



async function readPDF(file){

const reader=new FileReader();

reader.onload=async function(){

const typedarray=new Uint8Array(this.result);

const pdf=await pdfjsLib.getDocument(typedarray).promise;

let text="";

for(let pageNum=1;pageNum<=pdf.numPages;pageNum++){

const page=await pdf.getPage(pageNum);

const content=await page.getTextContent();

content.items.forEach(item=>{

text+=item.str+" ";

});

}

extractData(text);

};

reader.readAsArrayBuffer(file);

}



function extractData(text){

const tbody=document.querySelector("#resultTable tbody");

let project="";
let wo="";
let tanggal="";

const projectMatch=text.match(/project\s*[:\-]?\s*(.*?)\s{2,}/i);
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

const row={

project,
wo,
tanggal,
no,
item,
qty

};

resultData.push(row);

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

}

}

}



function downloadExcel(){

if(resultData.length===0){

alert("Tidak ada data");

return;

}

const ws=XLSX.utils.json_to_sheet(resultData);

const wb=XLSX.utils.book_new();

XLSX.utils.book_append_sheet(wb,ws,"BOQ");

XLSX.writeFile(wb,"summary_boq_lms.xlsx");

}
