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

// ================= MATCH =================
function smartMatch(templateItem, lmsItems){

let templateWords = normalize(templateItem).split(" ")

let bestKey = null
let bestScore = 0

for(let key in lmsItems){

let keyWords = key.split(" ")

let common = templateWords.filter(w => keyWords.includes(w))
if(common.length === 0) continue

let score = similarity(templateItem, key)

if(score > bestScore && score >= 0.75){
bestScore = score
bestKey = key
}

}

return bestKey
}

// ================= AMBIL WO SAJA =================
function extractInfo(rows){

let wo = ""

for(let r=0;r<20;r++){
for(let c=0;c<rows[r]?.length;c++){

let text = String(rows[r][c])

let match = text.match(/T\d{6,}-\d{6}(-\d+)?/)
if(match){
wo = match[0]
}

}
}

return {wo}
}

// ================= PROCESS =================
async function processFiles(){

const boqFile=document.getElementById("boqFile").files[0]
const lmsFiles=document.getElementById("lmsFiles").files

if(!boqFile) return alert("Upload BOQ Template dulu")
if(lmsFiles.length===0) return alert("Upload file LMS dulu")

await readBOQ(boqFile)

for(let i=0;i<lmsFiles.length;i++){

document.getElementById("status").innerText =
`⏳ Processing ${i+1}/${lmsFiles.length}...`

await new Promise(r => setTimeout(r,300))

let lmsData = await readLMS(lmsFiles[i])

fillBOQ(lmsData, i)

}

document.getElementById("status").innerText="✅ Selesai"
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

if(item && !isNaN(qty) && qty > 0){
let key = normalize(item)
items[key] = (items[key] || 0) + qty
}

}

let info = extractInfo(rows)

resolve({
items,
wo: info.wo
})

}

reader.readAsArrayBuffer(file)

})
}

// ================= FILL =================
function fillBOQ(lmsData,index){

const sheetName = boqWorkbook.SheetNames[0]
const sheet = boqWorkbook.Sheets[sheetName]

// isi WO saja
if(index === 0){
if(lmsData.wo) sheet["C3"] = { v: lmsData.wo }
}

// cari kolom
let startCol = 5
let hargaCol = 4

let col = startCol + (index*2)
let totalCol = col + 1

for(let i=5;i<boqData.length;i++){

let item = boqData[i]?.[1]
let harga = Number(boqData[i]?.[hargaCol]) || 0

if(!item) continue

let matchKey = smartMatch(item, lmsData.items)

if(matchKey){

let qty = lmsData.items[matchKey]
let total = qty * harga

let cQty = XLSX.utils.encode_cell({r:i,c:col})
let cTot = XLSX.utils.encode_cell({r:i,c:totalCol})

sheet[cQty] = { v: qty }
sheet[cTot] = { v: total }

}

}

}

// ================= DOWNLOAD =================
function downloadBOQ(){
XLSX.writeFile(boqWorkbook,"BOQ_FINAL.xlsx")
}
