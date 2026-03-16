let summaryData = [];
let ossMap = {};


/* FORMAT TANGGAL */
function formatDate(value){

if(!value) return "";

if(typeof value === "number"){

const d = XLSX.SSF.parse_date_code(value);

if(!d) return value;

let day = String(d.d).padStart(2,"0");
let month = String(d.m).padStart(2,"0");
let year = d.y;

return `${day}-${month}-${year}`;

}

let date = new Date(value);

if(isNaN(date)) return value;

let day = String(date.getDate()).padStart(2,"0");
let month = String(date.getMonth()+1).padStart(2,"0");
let year = date.getFullYear();

return `${day}-${month}-${year}`;

}



/* LOAD OSS (BISA BANYAK FILE) */

async function loadOSS(){

const files=document.getElementById("ossInput").files;

if(files.length===0){
alert("Upload file OSS dulu");
return;
}

ossMap={};

for(const file of files){

const data=await file.arrayBuffer();

const workbook=XLSX.read(data);

const sheet=workbook.Sheets[workbook.SheetNames[0]];

const json=XLSX.utils.sheet_to_json(sheet);

json.forEach(row=>{

const id=row["Cust ID Klien"];
const name=row["Customer Name"];

if(id){
ossMap[id]=name || "";
}

});

}

alert("OSS data loaded : "+Object.keys(ossMap).length+" customer");

}



/* PROSES MYREP */

function processFile(){

const file=document.getElementById("fileInput").files[0];

if(!file){
alert("Upload file MyRep dulu");
return;
}

document.getElementById("progress").innerText="Membaca file...";

const reader=new FileReader();

reader.onload=function(e){

const data=new Uint8Array(e.target.result);

const workbook=XLSX.read(data,{type:"array"});

const sheet=workbook.Sheets[workbook.SheetNames[0]];

const json=XLSX.utils.sheet_to_json(sheet);

generateSummary(json);

};

reader.readAsArrayBuffer(file);

}



/* GENERATE SUMMARY */

function generateSummary(data){

const tbody=document.querySelector("#resultTable tbody");

tbody.innerHTML="";

summaryData=[];

document.getElementById("progress").innerText=
"Memproses "+data.length+" data...";

data.forEach(row=>{

const item={

"CUSTOMER ID":row.subscription_id || "",

"CUSTOMER NAME":ossMap[row.subscription_id] || "",

"CUSTOMER ADDRESS":row.address || "",

"CLUSTER":row.cluster_name || "",

"FAT NO":row.fat_code || "",

"WO ID":row.work_order_number || "",

"WO DATE SCHEDULING":formatDate(row.service_activation_date),

"BAST ID":row.work_order_number || "",

"BAST DATE":formatDate(row.service_activation_date),

"SERIAL NUMBER":row.ont_serial_number || "",

"MAC ADDRESS":row.ont_mac_address || "",

"Status":row.service_status || "",

"Ket":row.workorder_status || "",

"AREA":row.area || ""

};

summaryData.push(item);

const tr=document.createElement("tr");

Object.values(item).forEach(val=>{

const td=document.createElement("td");
td.textContent=val;
tr.appendChild(td);

});

tbody.appendChild(tr);

});

document.getElementById("progress").innerText=
"Selesai. Total "+summaryData.length+" data.";

}



/* DOWNLOAD EXCEL */

function downloadExcel(){

if(summaryData.length===0){
alert("Belum ada data");
return;
}

const worksheet=XLSX.utils.json_to_sheet(summaryData);

const workbook=XLSX.utils.book_new();

XLSX.utils.book_append_sheet(workbook,worksheet,"Summary");

XLSX.writeFile(workbook,"myrep_summary.xlsx");

}
