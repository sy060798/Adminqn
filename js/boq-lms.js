let boqWorkbook;
let boqData;

function normalize(text){
return String(text)
.toLowerCase()
.replace(/[^a-z0-9]/g,"")
.trim()
}

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

let lmsItems=await readLMS(lmsFiles[i])

fillBOQ(lmsItems,lmsFiles[i].name,i)

}

document.getElementById("status").innerText="✅ Selesai ✔"
}

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

function readLMS(file){

return new Promise(resolve=>{

const reader=new FileReader()

reader.onload=e=>{

const data=new Uint8Array(e.target.result)

const wb=XLSX.read(data,{type:'array'})

let sheet=wb.Sheets["BoQ Aktual (Mitra)"]

if(!sheet){
sheet=wb.Sheets[wb.SheetNames[0]]
}

const rows=XLSX.utils.sheet_to_json(sheet,{header:1})

let items={}
let itemCol=-1
let qtyCol=-1
let headerRow=0

for(let r=0;r<10;r++){

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

if(itemCol==-1 || qtyCol==-1){
resolve({})
return
}

for(let i=headerRow+1;i<rows.length;i++){

let item=rows[i][itemCol]
let qty=Number(rows[i][qtyCol])

if(item && qty){
items[normalize(item)]=qty
}

}

resolve(items)

}

reader.readAsArrayBuffer(file)

})

}

function fillBOQ(lmsItems,fileName,index){

let header=fileName.replace(".xlsx","")

let startCol=0

for(let c=0;c<boqData[0].length;c++){
if(String(boqData[0][c]).toLowerCase().includes("lms")){
startCol=c
break
}
}

let col=startCol+(index*2)
let totalCol=col+1

boqData[1][col]=header
boqData[1][totalCol]="TOTAL"

for(let i=5;i<boqData.length;i++){

let item=boqData[i][1]
let harga=Number(boqData[i][2]) || 0

if(!item) continue

let key=normalize(item)

if(lmsItems[key]){

let qty=lmsItems[key]
let total=qty*harga

boqData[i][col]=qty
boqData[i][totalCol]=total

}

}

let grandTotal=0

for(let i=5;i<boqData.length;i++){
grandTotal+=Number(boqData[i][totalCol]||0)
}

let lastRow=boqData.length

boqData[lastRow]=[]
boqData[lastRow][1]="GRAND TOTAL"
boqData[lastRow][totalCol]=grandTotal

const newSheet=XLSX.utils.aoa_to_sheet(boqData)

for(let i=5;i<=lastRow;i++){

let cell=XLSX.utils.encode_cell({r:i,c:totalCol})

if(newSheet[cell]){
newSheet[cell].z='"Rp"#,##0'
}

}

boqWorkbook.Sheets[boqWorkbook.SheetNames[0]]=newSheet

}

function downloadBOQ(){

if(!boqWorkbook){
alert("Proses dulu sebelum download")
return
}

XLSX.writeFile(boqWorkbook,"BOQ_REKAP_LMS.xlsx")

}
