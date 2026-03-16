let boqWorkbook
let boqData

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

// ambil sheet LMS
const sheet=wb.Sheets["BoQ Aktual (Mitra)"]

if(!sheet){
resolve([])
return
}

const rows=XLSX.utils.sheet_to_json(sheet,{header:1})

let qtyCol=-1

// cari kolom header "BoQ Aktual (Mitra)"
for(let c=0;c<rows[0].length;c++){

let header=String(rows[0][c]).toLowerCase()

if(header.includes("aktual")){
qtyCol=c
break
}

}

if(qtyCol===-1){
resolve([])
return
}

let qtyList=[]

// ambil semua qty di bawah header
for(let i=1;i<rows.length;i++){

let qty=Number(rows[i][qtyCol])

if(qty){
qtyList.push(qty)
}

}

resolve(qtyList)

}

reader.readAsArrayBuffer(file)

})

}

function fillBOQ(lmsRows,fileName,index){

let header=fileName.replace(".xlsx","")

// cari kolom LMS pertama
let startCol=0

for(let c=0;c<boqData[0].length;c++){

if(String(boqData[0][c]).toLowerCase().includes("lms")){
startCol=c
break
}

}

// setiap LMS punya 2 kolom (Qty + Total)
let col=startCol+(index*2)

// isi nama file di baris judul LMS
boqData[1][col]=header

let lmsIndex=0

// mulai isi item
for(let i=5;i<boqData.length;i++){

if(!boqData[i][1]) continue

let qty=lmsRows[lmsIndex]

if(qty){
boqData[i][col]=qty
}

lmsIndex++

}

// update sheet
const newSheet=XLSX.utils.aoa_to_sheet(boqData)

boqWorkbook.Sheets[boqWorkbook.SheetNames[0]]=newSheet

}

function downloadBOQ(){

XLSX.writeFile(boqWorkbook,"BOQ_REKAP_LMS.xlsx")

}
