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

// ================= SIMILARITY =================
function similarity(a, b){
a = normalize(a)
b = normalize(b)

let longer = a.length > b.length ? a : b
let shorter = a.length > b.length ? b : a

let longerLength = longer.length
if(longerLength === 0) return 1

return (longerLength - editDistance(longer, shorter)) / longerLength
}

function editDistance(a, b){
let matrix = []

for(let i=0;i<=b.length;i++) matrix[i] = [i]
for(let j=0;j<=a.length;j++) matrix[0][j] = j

for(let i=1;i<=b.length;i++){
for(let j=1;j<=a.length;j++){
if(b[i-1] === a[j-1]){
matrix[i][j] = matrix[i-1][j-1]
}else{
matrix[i][j] = Math.min(
matrix[i-1][j-1] + 1,
matrix[i][j-1] + 1,
matrix[i-1][j] + 1
)
}
}
}

return matrix[b.length][a.length]
}

// ================= MATCH ITEM =================
function findBestMatch(templateItem, lmsItems){

let bestKey = null
let bestScore = 0

for(let key in lmsItems){

let score = similarity(templateItem, key)

if(score > bestScore && score > 0.5){
bestScore = score
bestKey = key
}

}

return bestKey
}

// ================= FIND COLUMN (ANTI ERROR) =================
function findColumns(){

let startCol = null
let hargaCol = null

for(let r=0;r<15;r++){
for(let c=0;c<boqData[r]?.length;c++){

let txt = normalize(boqData[r][c])

// QTY / LMS
if(
txt.includes("boq") ||
txt.includes("aktual") ||
txt.includes("qty") ||
txt.includes("lms")
){
if(startCol === null) startCol = c
}

// HARGA
if(
txt.includes("harga") ||
txt.includes("price") ||
txt.includes("satuan")
){
hargaCol = c
}

}
}

// fallback (biar gak error)
if(startCol === null) startCol = 5
if(hargaCol === null) hargaCol = 4

return {startCol, hargaCol}
}

// ================= AMBIL WO & PROJECT =================
function extractInfo(rows){

let wo = ""
let project = ""

for(let r=0;r<15;r++){
if(!rows[r]) continue

for(let c=0;c<rows[r].length;c++){

let text = String(rows[r][c])

// WO
let match = text.match(/T\d{6,}-\d+/)
if(match) wo = match[0]

// PROJECT
if(text.toLowerCase().includes("nama project")){
let parts = text.split(":")
project = parts[1] ? parts[1].trim() : text
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

// cari header fleksibel
for(let r=0;r<15;r++){

if(!rows[r]) continue

for(let c=0;c<rows[r].length;c++){

let text=normalize(rows[r][c])

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
items[key] = (items[key] || 0) + qty
}

}

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

// isi project & WO sekali
if(index === 0){
if(lmsData.project) sheet["C2"] = { v: lmsData.project }
if(lmsData.wo) sheet["C3"] = { v: lmsData.wo }
}

// ambil kolom
let cols = findColumns()
let startCol = cols.startCol
let hargaCol = cols.hargaCol

let col = startCol + (index*2)
let totalCol = col + 1

// isi data
for(let i=5;i<boqData.length;i++){

let item = boqData[i]?.[1]
let harga = Number(boqData[i]?.[hargaCol]) || 0

if(!item) continue

let matchKey = findBestMatch(item, lmsData.items)

if(matchKey){

let qty = lmsData.items[matchKey]
let total = qty * harga

let cQty = XLSX.utils.encode_cell({r:i,c:col})
let cTot = XLSX.utils.encode_cell({r:i,c:totalCol})

sheet[cQty] = { v: qty, t:"n" }
sheet[cTot] = { v: total, t:"n", z:'"Rp"#,##0' }

}

}

// GRAND TOTAL
let grand = 0

for(let i=5;i<boqData.length;i++){
let cTot = XLSX.utils.encode_cell({r:i,c:totalCol})
if(sheet[cTot]) grand += Number(sheet[cTot].v || 0)
}

// isi ke baris GRAND TOTAL
for(let i=boqData.length-1;i>=0;i--){
if(boqData[i] && String(boqData[i][1]).toLowerCase().includes("grand total")){
let cGT = XLSX.utils.encode_cell({r:i,c:totalCol})
sheet[cGT] = { v: grand, t:"n", z:'"Rp"#,##0' }
break
}
}

}

// ================= DOWNLOAD =================
function downloadBOQ(){
if(!boqWorkbook) return alert("Proses dulu")
XLSX.writeFile(boqWorkbook,"BOQ_FINAL.xlsx")
}
