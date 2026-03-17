function fillBOQ(lmsItems,fileName,index){

let header=fileName.replace(".xlsx","")

// cari kolom LMS
let startCol=0

for(let c=0;c<boqData[0].length;c++){

if(String(boqData[0][c]).toLowerCase().includes("lms")){
startCol=c
break
}

}

// 🔥 kolom qty & total
let col=startCol+(index*2)
let totalCol=col+1

// isi nama file
boqData[1][col]=header
boqData[1][totalCol]="TOTAL"

// ==========================
// ISI DATA + HITUNG TOTAL
// ==========================
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

// ==========================
// GRAND TOTAL
// ==========================
let grandTotal=0

for(let i=5;i<boqData.length;i++){
grandTotal+=Number(boqData[i][totalCol]||0)
}

let lastRow=boqData.length

boqData[lastRow]=[]
boqData[lastRow][1]="GRAND TOTAL"
boqData[lastRow][totalCol]=grandTotal

// ==========================
// BUAT SHEET BARU
// ==========================
const newSheet=XLSX.utils.aoa_to_sheet(boqData)

// ==========================
// FORMAT RUPIAH
// ==========================
for(let i=5;i<=lastRow;i++){

let cellAddress=XLSX.utils.encode_cell({r:i,c:totalCol})

if(newSheet[cellAddress]){
newSheet[cellAddress].z='"Rp"#,##0'
}

}

// replace sheet
boqWorkbook.Sheets[boqWorkbook.SheetNames[0]]=newSheet

}
