let boqWorkbook
let boqData

function normalize(text){

return String(text)
.toLowerCase()
.replace(/[^a-z0-9]/g,"")
.trim()

}

function setStatus(t){
document.getElementById("status").innerText=t
}

function setProgress(p){

const bar=document.getElementById("progressBar")

bar.style.width=p+"%"
bar.innerText=p+"%"

}

async function processFiles(){

const boqFile=document.getElementById("boqFile").files[0]
const lmsFiles=document.getElementById("lmsFiles").files

if(!boqFile){
alert("Upload BOQ dulu")
return
}

if(lmsFiles.length===0){
alert("Upload LMS")
return
}

setStatus("Membaca BOQ...")
setProgress(10)

await readBOQ(boqFile)

for(let i=0;i<lmsFiles.length;i++){

setStatus("Membaca "+lmsFiles[i].name)

let lmsItems=await readLMS(lmsFiles[i])

fillBOQ(lmsItems,i,lmsFiles[i].name)

let percent=10+Math.round((i+1)/lmsFiles.length*70)

setProgress(percent)

}

setStatus("Selesai ✔")
setProgress(100)

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

rows.forEach((r,i)=>{

if(i===0) return

let item=r[1]
let qty=Number(r[4])

if(item && qty){

let key=normalize(item)

items[key]=qty

}

})

resolve(items)

}

reader.readAsArrayBuffer(file)

})

}

function fillBOQ(lmsItems,index,fileName){

let col=5+index

let header=fileName.replace(".xlsx","")

boqData.forEach((r,i)=>{

if(i===10){

r[col]=header
return

}

if(i<11) return

let item=r[1]

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

if(!boqWorkbook){
alert("Belum ada data")
return
}

XLSX.writeFile(boqWorkbook,"BOQ_UPDATED.xlsx")

}
