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

let qtyCol=-1

// cari kolom "BoQ Aktual (Mitra)"
for(let c=0;c<rows[0].length;c++){

let header=String(rows[0][c]).toLowerCase()

if(header.includes("aktual")){
qtyCol=c
break
}

}

let items={}

rows.forEach(r=>{

let item=r[1]

let qty=Number(r[qtyCol])

if(item && qty){

items[normalize(item)]=qty

}

})

resolve(items)

}

reader.readAsArrayBuffer(file)

})

}
