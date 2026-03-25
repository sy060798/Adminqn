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

if(temp.includes(key) || key.includes(temp)) return key

let words = temp.split(" ")
for(let w of words){
if(w.length > 3 && key.includes(w)){
return key
}
}

}

return null
}

// ================= AMBIL PROJECT & WO =================
function extractInfo(rows){

let wo = ""
let project = ""

for(let r=0;r<10;r++){
if(!rows[r]) continue

for(let c=0;c<rows[r].length;c++){

let text = String(rows[r][c])

// ambil WO
if(text.match(/T\d{6,}-\d+/)){
wo = text.match(/T\d{6,}-\d+/)[0]
}

// ambil nama project
if(text.toLowerCase().includes("nama project")){
project = rows[r][c+1] || text
}

}
}

return {wo, project}
}

// ================= PROCESS =================
async function processFiles(){

const boqFile=document.getElementById("boqFile").files[0]
const lmsFiles=document.getElementById("lmsFiles").files

if(!boqFile) return alert("Upload BOQ Template dulu")
if(lmsFiles.length===0) return alert("Upload file LMS dulu")

document.getElementById("status").innerText="⏳ Memproses..."

await readBOQ(boqFile)

for(let i=0;i<lmsFiles.length;i++){

await new Promise(r => setTimeout(r,50))

let lmsData=await readLMS(lmsFiles[i])

fillBOQ(lmsData,i)

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

// cari header
for(let r=0;r<10;r++){

if(!rows[r]) continue

for(let c=0;c<rows[r].length;c++){

let text=String(rows[r][c]).toLowerCase()

if(text.includes("item")) itemCol=c
if(text.includes("boq aktual")) qtyCol=c

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

items[key] = (items[key] || 0) + qty

}

}

// ambil info project & wo
let info = extractInfo(rows)

resolve({
items,
wo: info.wo,
project: info.project
})

}

reader.readAsArrayBuffer(file)

})
}

// ================= FILL BOQ =================
function fillBOQ(lmsData,index){

const sheetName = boqWorkbook.SheetNames[0]
const sheet = boqWorkbook.Sheets[sheetName]

let lmsItems = lmsData.items

// 🔥 ISI HEADER UTAMA (CUMA SEKALI)
if(index === 0){
sheet["C2"] = { v: lmsData.project }
sheet["C3"] = { v: lmsData.wo }
}

// cari kolom LMS
let startCol=0

for(let c=0;c<boqData[0].length;c++){
if(String(boqData[0][c]).toLowerCase().includes("boq aktual")){
startCol=c
break
}
}

// posisi kolom
let col=startCol+(index*2)
let totalCol=col+1

// isi QTY & TOTAL
for(let i=5;i<boqData.length;i++){

let item=boqData[i]?.[1]
let harga=Number(boqData[i]?.[3]) || 0

if(!item) continue

let matchKey = isMatch(item, lmsItems)

if(matchKey){

let qty = lmsItems[matchKey]
let total = qty * harga

let cellQty = XLSX.utils.encode_cell({r:i,c:col})
let cellTotal = XLSX.utils.encode_cell({r:i,c:totalCol})

if(!sheet[cellQty]) sheet[cellQty] = {}
sheet[cellQty].v = qty

if(!sheet[cellTotal]) sheet[cellTotal] = {}
sheet[cellTotal].v = total
sheet[cellTotal].z = '"Rp"#,##0'

}

}

// GRAND TOTAL
let grandTotal = 0

for(let i=5;i<boqData.length;i++){
let cellTotal = XLSX.utils.encode_cell({r:i,c:totalCol})
if(sheet[cellTotal]){
grandTotal += Number(sheet[cellTotal].v || 0)
}
}

let totalRow = boqData.findIndex(r =>
r && String(r[1]).toLowerCase().includes("total")
)

if(totalRow !== -1){

let cellGT = XLSX.utils.encode_cell({r:totalRow,c:totalCol})

sheet[cellGT] = {
v: grandTotal,
t: "n",
z: '"Rp"#,##0'
}

}

}

// ================= DOWNLOAD =================
function downloadBOQ(){

if(!boqWorkbook){
alert("Proses dulu")
return
}

XLSX.writeFile(boqWorkbook,"BOQ_FINAL.xlsx")

}
