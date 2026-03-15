let resultData=[];

async function processPDF(){

const file=document.getElementById("pdfFile").files[0];

if(!file){
alert("Upload PDF dulu");
return;
}

const reader=new FileReader();

reader.onload=async function(){

const typedarray=new Uint8Array(this.result);

const pdf=await pdfjsLib.getDocument(typedarray).promise;

let lines=[];

for(let pageNum=1;pageNum<=pdf.numPages;pageNum++){

const page=await pdf.getPage(pageNum);

const textContent=await page.getTextContent();

textContent.items.forEach(item=>{
lines.push(item.str.trim());
});

}

parseBOQ(lines);

};

reader.readAsArrayBuffer(file);

}



function parseBOQ(lines){

const tbody=document.querySelector("#resultTable tbody");

tbody.innerHTML="";

resultData=[];

let project="";
let wo="";
let tanggal="";

lines.forEach(line=>{

let lower=line.toLowerCase();

if(lower.includes("project")){
project=line.replace(/project/i,"").trim();
}

if(lower.includes("wo")){
wo=line.replace(/no/i,"").trim();
}

if(lower.includes("tanggal")){
tanggal=line.replace(/tanggal/i,"").trim();
}

});

for(let i=0;i<lines.length;i++){

let no=lines[i];

if(!isNaN(no)){

let item=lines[i+1];
let qty=lines[i+2];

if(!isNaN(qty) && Number(qty)>0){

const row={
project,
wo,
tanggal,
no,
item,
qty
};

resultData.push(row);

const tr=document.createElement("tr");

tr.innerHTML=`
<td>${project}</td>
<td>${wo}</td>
<td>${tanggal}</td>
<td>${no}</td>
<td>${item}</td>
<td>${qty}</td>
`;

tbody.appendChild(tr);

}

}

}

}



function downloadExcel(){

if(resultData.length===0){

alert("Tidak ada data");

return;

}

const ws=XLSX.utils.json_to_sheet(resultData);

const wb=XLSX.utils.book_new();

XLSX.utils.book_append_sheet(wb,ws,"BOQ");

XLSX.writeFile(wb,"summary_boq_lms.xlsx");

}
