let resultData = [];

async function processPDF(){

const file = document.getElementById("pdfFile").files[0];

if(!file){
alert("Upload PDF dulu");
return;
}

const reader = new FileReader();

reader.onload = async function(){

const typedarray = new Uint8Array(this.result);

const pdf = await pdfjsLib.getDocument(typedarray).promise;

let textContent = "";

for(let i=1;i<=pdf.numPages;i++){

const page = await pdf.getPage(i);

const text = await page.getTextContent();

text.items.forEach(item=>{
textContent += item.str + " ";
});

}

parseBOQ(textContent);

};

reader.readAsArrayBuffer(file);

}


function parseBOQ(text){

const tbody = document.querySelector("#resultTable tbody");

tbody.innerHTML="";

resultData=[];

let project="";
let wo="";
let tanggal="";

const projectMatch = text.match(/Project\s*[:\-]?\s*(.*?)\s{2,}/i);
if(projectMatch) project = projectMatch[1];

const woMatch = text.match(/WO\s*[:\-]?\s*(\S+)/i);
if(woMatch) wo = woMatch[1];

const dateMatch = text.match(/Tanggal\s*[:\-]?\s*(\S+)/i);
if(dateMatch) tanggal = dateMatch[1];

const itemRegex = /(\d+)\s+([A-Za-z0-9\s\(\)\-\/]+?)\s+(\d+)/g;

let match;

while((match = itemRegex.exec(text)) !== null){

let no = match[1];
let item = match[2].trim();
let qty = match[3];

if(qty > 0){

const row = {

project,
wo,
tanggal,
no,
item,
qty

};

resultData.push(row);

const tr = document.createElement("tr");

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

if(resultData.length==0){
alert("Tidak ada data");
return;
}

const ws = XLSX.utils.json_to_sheet(resultData);

const wb = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(wb,ws,"BOQ");

XLSX.writeFile(wb,"summary_boq.xlsx");

}
