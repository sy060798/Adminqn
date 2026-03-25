let boqWorkbook;
let boqData;

// ================= NORMALIZE =================
function normalize(text){
return String(text)
.toLowerCase()
.replace(/[^a-z0-9 ]/g," ")
.replace(/\s+/g," ")
.trim()
}

// ================= MATCH FLEXIBLE =================
function isMatch(templateItem, lmsItems){

let temp = normalize(templateItem)

for(let key in lmsItems){

if(temp === key) return key

if(temp.includes(key) || key.includes(temp)){
return key
}

let words = temp.split(" ")

for(let w of words){
if(w.length > 3 && key.includes(w)){
return key
}
}

}

return null
}

// ================= AMBIL WO =================
function extractWO(fileName){
let match = fileName.match(/T\d{6,}-\d+/)
return match ? match[0] : fileName.replace(".xlsx","")
}

// ================= PROCESS =================
async function processFiles(){

const boqFile=document.getElementById("boqFile").files[0]
const lmsFiles=document.getElementById("lmsFiles").files

if(!boqFile){
alert("Upload BOQ Template dulu")
return
}

if(lmsFiles.length===0){
alert("Upload file LMS dulu")
return
}

document.getElementById("status").innerText="⏳ Memproses..."

await readBOQ(boqFile)

for(let i=0;i<lmsFiles.length;i++){

await new Promise(r => setTimeout(r,50))

let lmsData=await readLMS(lmsFiles[i])

fillBOQ(lmsData,lmsFiles[i].name,i)

}

document.getElementById("status").innerText="✅ Selesai ✔"
}

// ================= READ BOQ =================
function readBOQ(file){

return new Promise(resolve=>{

const reader=new FileReader()

reader.onload=e=>{

const data=new Uint8Array(e.target.result)

boqWorkbook=XLSX.read(data,{type:'array'})

const sheet=boqWorkbook.Sheets[boqWorkbook.SheetNames[0]]

boqData=XLSX.utils.sheet_to_json(sheet,{header:1})

resolve()

}

reader.readAsArrayBuffer(file)

})
}

// ================= READ LMS =================
function readLMS(file){

return new Promise(resolve=>{

const reader=new FileReader()

reader.onload=e=>{

const data=new Uint8Array(e.target.result)

const wb=XLSX.read(data,{type:'array'})

let sheet=wb.Sheets["BoQ Aktual (Mitra)"]
if(!sheet) sheet=wb.Sheets[wb.SheetNames[0]]

const rows=XLSX.utils.sheet_to_json(sheet,{header:1})

let items={}
let itemCol=-1
let qtyCol=-1
let headerRow=0

for(let r=0;r<10;r++){

if(!rows[r]) continue

for(let c=0;c<rows[r].length;c++){

let text=String(rows[r][c]).toLowerCase()

if(text.includes("item")) itemCol=c
if(text.includes("boq") || text.includes("qty")) qtyCol=c

}

if(itemCol!=-1 && qtyCol!=-1){
headerRow=r
break
}

}

// ambil data
for(let i=headerRow+1;i<rows.length;i++){

let item=rows[i]?.[itemCol]
let qty=Number(rows[i]?.[qtyCol])

if(item && !isNaN(qty)){

let key = normalize(item)

if(items[key]){
items[key] += qty
}else{
items[key] = qty
}

}

}

resolve({
items,
wo: extractWO(file.name),
project: file.name.replace(".xlsx","")
})

}

reader.readAsArrayBuffer(file)

})
}

// ================= FILL BOQ =================
function fillBOQ(lmsData,fileName,index){

const sheetName = boqWorkbook.SheetNames[0]
const sheet = boqWorkbook.Sheets[sheetName]

let lmsItems = lmsData.items
let wo = lmsData.wo

// tampilkan WO
sheet["C1"] = { v: "WO : " + wo }

// cari kolom LMS awal
let startCol=0

for(let c=0;c<boqData[0].length;c++){
if(String(boqData[0][c]).toLowerCase().includes("lms")){
startCol=c
break
}
}

// posisi kolom (QTY & TOTAL)
let col=startCol+(index*2)
let totalCol=col+1

// HEADER SESUAI CONTOH KAMU
let headerCell = XLSX.utils.encode_cell({r:1,c:col})
let totalHeaderCell = XLSX.utils.encode_cell({r:2,c:col})

sheet[headerCell] = { v: wo }      // baris 2: WO
sheet[totalHeaderCell] = { v: "QTY" }

let headerTotal = XLSX.utils.encode_cell({r:2,c:totalCol})
sheet[headerTotal] = { v: "TOTAL" }

// isi data
for(let i=5;i<boqData.length;i++){

let item=boqData[i]?.[1]
let harga=Number(boqData[i]?.[2]) || 0

if(!item) continue

let matchKey = isMatch(item, lmsItems)

if(matchKey){

let qty = lmsItems[matchKey]
let total = qty * harga

let cellQty = XLSX.utils.encode_cell({r:i,c:col})
let cellTotal = XLSX.utils.encode_cell({r:i,c:totalCol})

if(!sheet[cellQty]) sheet[cellQty] = {}
sheet[cellQty].v = qty
sheet[cellQty].t = "n"

if(!sheet[cellTotal]) sheet[cellTotal] = {}
sheet[cellTotal].v = total
sheet[cellTotal].t = "n"
sheet[cellTotal].z = '"Rp"#,##0'

}

}

// GRAND TOTAL ke baris yang SUDAH ADA
let grandTotal = 0

for(let i=5;i<boqData.length;i++){
let cellTotal = XLSX.utils.encode_cell({r:i,c:totalCol})
if(sheet[cellTotal]){
grandTotal += Number(sheet[cellTotal].v || 0)
}
}

// cari tulisan GRAND TOTAL
let totalRow = boqData.findIndex(row =>
row && String(row[1]).toLowerCase().includes("grand total")
)

if(totalRow !== -1){

let cellGT = XLSX.utils.encode_cell({r:totalRow,c:totalCol})

if(!sheet[cellGT]) sheet[cellGT] = {}
sheet[cellGT].v = grandTotal
sheet[cellGT].t = "n"
sheet[cellGT].z = '"Rp"#,##0'

}

}

// ================= DOWNLOAD =================
function downloadBOQ(){

if(!boqWorkbook){
alert("Proses dulu sebelum download")
return
}

XLSX.writeFile(boqWorkbook,"BOQ_REKAP_LMS.xlsx")

}
