let boqWorkbook;
let boqData;
let woList = []

// ================= NORMALIZE =================
function normalize(text){
return String(text || "")
.toLowerCase()
.replace(/[^a-z0-9: ]/g," ")
.replace(/\s+/g," ")
.trim()
}

// ================= RATIO =================
function extractRatio(text){
let match = text.match(/\d+:\d+/)
return match ? match[0] : null
}

// ================= KEYWORD =================
function getMainKeyword(text){
let words = normalize(text).split(" ")
return words.filter(w => 
!w.match(/\d+:\d+/) &&
w !== "pcs" &&
w !== "unit" &&
w.length > 2
)
}

// ================= SIMILARITY =================
function similarity(a, b){
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

let templateNorm = normalize(templateItem)
let templateRatio = extractRatio(templateNorm)
let templateKeywords = getMainKeyword(templateNorm)

let bestKey = null
let bestScore = 0

for(let key in lmsItems){

let keyNorm = normalize(key)
let keyRatio = extractRatio(keyNorm)
let keyKeywords = getMainKeyword(keyNorm)

// wajib sama ratio
if(templateRatio && keyRatio && templateRatio !== keyRatio){
continue
}

// wajib ada keyword sama
let common = templateKeywords.filter(w => keyKeywords.includes(w))
if(common.length === 0) continue

let score = similarity(templateNorm, keyNorm)

if(score > bestScore && score >= 0.8){
bestScore = score
bestKey = key
}

}

return bestKey
}

// ================= AMBIL WO =================
function extractInfo(rows){

let wo = ""

for(let r=0;r<20;r++){
for(let c=0;c<rows[r]?.length;c++){

let text = String(rows[r][c])
let match = text.match(/T\d{6,}-\d{6}(-\d+)?/)

if(match) wo = match[0]

}
}

return {wo}
}

// ================= PROCESS =================
async function processFiles(){

try{

const boqFile=document.getElementById("boqFile").files[0]
const lmsFiles=document.getElementById("lmsFiles").files

if(!boqFile){
alert("Upload BOQ dulu")
return
}

if(lmsFiles.length===0){
alert("Upload LMS dulu")
return
}

document.getElementById("status").innerText = "⏳ Membaca BOQ..."

await readBOQ(boqFile)

for(let i=0;i<lmsFiles.length;i++){

document.getElementById("status").innerText =
`⏳ Processing ${i+1}/${lmsFiles.length}...`

await new Promise(r => setTimeout(r,200))

let lmsData = await readLMS(lmsFiles[i])

woList[i] = lmsData.wo

fillBOQ(lmsData, i)

}

document.getElementById("status").innerText="✅ Selesai ✔"

}catch(err){
console.error(err)
alert("Error: " + err.message)
}

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

sheet[cQty] = { v: qty, t:"n" }
sheet[cTot] = { v: total, t:"n", z:'"Rp"#,##0' }

}

}

// WO di bawah
let lastRow = boqData.length + 2

let woCell = XLSX.utils.encode_cell({r:lastRow,c:col})
sheet[woCell] = { v: woList[index] || "-", t:"s" }

if(index === 0){
let labelCell = XLSX.utils.encode_cell({r:lastRow,c:col-1})
sheet[labelCell] = { v: "NO WO", t:"s" }
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

// 🔥 WAJIB (BIAR BUTTON JALAN)
window.processFiles = processFiles;
window.downloadBOQ = downloadBOQ;
