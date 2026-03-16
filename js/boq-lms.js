let boqWorkbook
let boqData

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

await readBOQ(boqFile)

for(let i=0;i<lmsFiles.length;i++){

let lmsItems=await readLMS(lmsFiles[i])

fillBOQ(lmsItems,lmsFiles[i].name,i)

}

document.getElementById("status").innerText="Selesai ✔"

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

const sheet=wb.Sheets["BoQ Aktual (Mitra)"]

if(!sheet){
resolve({})
return
}

const rows=XLSX.utils.sheet_to_json(sheet,{header:1})

let items={}

for(let i=1;i<rows.length;i++){

let item=rows[i][1]   // kolom Item
let qty=Number(rows[i][3]) // kolom BoQ Aktual (Mitra)

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

// cari kolom LMS pertama
let startCol=0

for(let c=0;c<boqData[0].length;c++){

if(String(boqData[0][c]).toLowerCase().includes("lms")){
startCol=c
break
}

}

// tiap LMS punya Qty + Total
let col=startCol+(index*2)

// isi judul LMS
boqData[1][col]=header

for(let i=5;i<boqData.length;i++){

let item=boqData[i][1]

if(!item) continue

let key=normalize(item)

if(lmsItems[key]){

boqData[i][col]=lmsItems[key]

}

}

const newSheet=XLSX.utils.aoa_to_sheet(boqData)

boqWorkbook.Sheets[boqWorkbook.SheetNames[0]]=newSheet

}

function downloadBOQ(){

XLSX.writeFile(boqWorkbook,"BOQ_REKAP_LMS.xlsx")

}
