// menyimpan hasil
let resultData = [];

// update status
function updateStatus(text){
document.getElementById("statusText").innerText = text;
}

// update progress bar
function updateProgress(done,total){

let percent = Math.round((done/total)*100);

let bar = document.getElementById("progressBar");

bar.style.width = percent + "%";
bar.innerText = percent + "%";

}

// proses semua PDF
async function processPDF(){

const files = document.getElementById("pdfFiles").files;

if(files.length === 0){
alert("Upload PDF dulu");
return;
}

// reset
resultData = [];
document.querySelector("#resultTable tbody").innerHTML = "";

for(let i=0;i<files.length;i++){

updateStatus("Membaca " + files[i].name);

await readPDF(files[i]);

updateProgress(i+1, files.length);

}

updateStatus("Selesai membaca semua PDF");

}

// membaca 1 file PDF
async function readPDF(file){

const project = file.name.replace(".pdf","");

const buffer = await file.arrayBuffer();

const pdf = await pdfjsLib.getDocument({data:buffer}).promise;

let words = [];

for(let p=1; p<=pdf.numPages; p++){

let page = await pdf.getPage(p);

let text = await page.getTextContent();

text.items.forEach(t => {
words.push(t.str.trim());
});

}

parseTable(words, project);

}

// membaca struktur tabel PDF
function parseTable(words, project){

for(let i=0;i<words.length;i++){

let no = parseInt(words[i]);

// jika angka = nomor tabel
if(!isNaN(no)){

let item = words[i+1];
let boq = parseInt(words[i+2]);
let aktual = parseInt(words[i+3]);

if(!isNaN(aktual) && aktual > 0){

addRow(project,item,aktual);

}

}

}

}

// tambah baris ke tabel
function addRow(project,item,qty){

let tbody = document.querySelector("#resultTable tbody");

let tr = document.createElement("tr");

tr.innerHTML = `
<td>${project}</td>
<td>${item}</td>
<td>${qty}</td>
`;

tbody.appendChild(tr);

resultData.push({
Project: project,
Item: item,
Qty: qty
});

}

// download hasil ke Excel
function downloadExcel(){

if(resultData.length === 0){

alert("Tidak ada data untuk didownload");
return;

}

let ws = XLSX.utils.json_to_sheet(resultData);

let wb = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(wb, ws, "BOQ");

XLSX.writeFile(wb, "BOQ_LMS_RESULT.xlsx");

}
