let boqWorkbook
let boqData

function normalize(text){

return String(text)
.toLowerCase()
.replace(/[^a-z0-9]/g,"")
trim()

}

async function processFiles(){

const boqFile=document.getElementById("boqFile").files[0]
const lmsFiles=document.getElementById("lmsFiles").files

if(!boqFile){
alert("Upload BOQ Template dulu")
return
}

if(lmsFiles.length===0){
alert("Upload file LMS")
return
}

await readBOQ(boqFile)

for(let i=0;i<lmsFiles.length;i++){

let lmsData=await readLMS(lmsFiles[i])

fillBOQ(lmsData,lmsFiles[i].name,i)

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

rows.forEach(r=>{

let item=r[1]
let qty=Number(r[4])

if(item && qty){

items[normalize(item)]=qty

}

})

resolve(items)

}

reader.readAsArrayBuffer(file)

})

}

function fillBOQ(lmsItems,fileName,index){

let col=index+1

let header=fileName.replace(".xlsx","")

boqData.forEach((r,i)=>{

if(i===0){

r[col]=header
return

}

let item=r[0]

if(!item) return

let key=normalize(item)

if(lmsItems[key]){

r[col]=lmsItems[key]

}

})

const newSheet=XLSX.utils.aoa_to_sheet(boqData)

boqWorkbook.Sheets[boqWorkbook.SheetNames[0]]=newSheet

}

function downloadBOQ(){

XLSX.writeFile(boqWorkbook,"BOQ_REKAP.xlsx")

}
