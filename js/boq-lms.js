let boqWorkbook;
let boqData;

let itemTotals={};

function setStatus(t){
document.getElementById("status").innerText=t;
}

function setProgress(p){

const bar=document.getElementById("progressBar");

bar.style.width=p+"%";
bar.innerText=p+"%";

}

function normalize(text){

return text
.toLowerCase()
.replace(/[^a-z0-9]/g,"")
.trim();

}

async function processFiles(){

const boqFile=document.getElementById("boqFile").files[0];
const lmsFiles=document.getElementById("lmsFiles").files;

if(!boqFile){
alert("Upload BOQ dulu");
return;
}

if(lmsFiles.length===0){
alert("Upload file LMS");
return;
}

itemTotals={};

setStatus("Membaca BOQ...");
setProgress(10);

await readBOQ(boqFile);

setStatus("Membaca file LMS...");
setProgress(20);

for(let i=0;i<lmsFiles.length;i++){

await readLMS(lmsFiles[i]);

let percent=20+Math.round((i+1)/lmsFiles.length*50);

setProgress(percent);

}

setStatus("Mencocokkan item...");
setProgress(80);

updateBOQ();

setProgress(100);

setStatus("Selesai ✔ Silakan download BOQ");

}

function readBOQ(file){

return new Promise(resolve=>{

const reader=new FileReader();

reader.onload=e=>{

const data=new Uint8Array(e.target.result);

boqWorkbook=XLSX.read(data,{type:'array'});

const sheet=boqWorkbook.Sheets[boqWorkbook.SheetNames[0]];

boqData=XLSX.utils.sheet_to_json(sheet,{header:1});

resolve();

};

reader.readAsArrayBuffer(file);

});

}

function readLMS(file){

return new Promise(resolve=>{

const reader=new FileReader();

reader.onload=e=>{

const data=new Uint8Array(e.target.result);

const wb=XLSX.read(data,{type:'array'});

const sheet=wb.Sheets[wb.SheetNames[0]];

const rows=XLSX.utils.sheet_to_json(sheet,{header:1});

rows.forEach(r=>{

let item=r[1];
let qty=r[4];

if(item && typeof qty==="number" && qty>0){

let key=normalize(item);

if(!itemTotals[key]) itemTotals[key]=0;

itemTotals[key]+=qty;

}

});

resolve();

};

reader.readAsArrayBuffer(file);

});

}

function updateBOQ(){

boqData.forEach(r=>{

let item=r[1];

if(!item) return;

let key=normalize(item);

for(let lmsItem in itemTotals){

if(lmsItem.includes(key) || key.includes(lmsItem)){

r[3]=itemTotals[lmsItem];

}

}

});

const newSheet=XLSX.utils.aoa_to_sheet(boqData);

boqWorkbook.Sheets[boqWorkbook.SheetNames[0]]=newSheet;

}

function downloadBOQ(){

if(!boqWorkbook){

alert("Belum ada data");

return;

}

XLSX.writeFile(boqWorkbook,"BOQ_UPDATED.xlsx");

}
