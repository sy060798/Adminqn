console.log("✅ JS KELOAD");

document.addEventListener("DOMContentLoaded", function(){

console.log("✅ DOM SIAP");

let allData = [];
let lastFiltered = [];
let currentWorkbook = null;

// ==========================
// INIT EVENT
// ==========================
const upload = document.getElementById('upload');
const btnScan = document.getElementById('btnScan');
const btnExport = document.getElementById('btnExport');

if(upload) upload.addEventListener('change', handleFile);
if(btnScan) btnScan.addEventListener('click', applyFilter);
if(btnExport) btnExport.addEventListener('click', exportExcel);

// ==========================
function setStatus(msg){
    const el = document.getElementById('status');
    if(el) el.innerText = msg;
}

// ==========================
// HANDLE FILE
// ==========================
function handleFile(e){

    console.log("📂 FILE DIPILIH");

    const file = e.target.files[0];
    if(!file) return;

    setStatus("⏳ Membaca file...");

    const reader = new FileReader();

    reader.onload = function(evt){

        console.log("📖 FILE DIBACA");

        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type:'array'});

        currentWorkbook = workbook;

        const sheetSelect = document.getElementById('sheetSelect');
        sheetSelect.innerHTML = "";

        workbook.SheetNames.forEach(name=>{
            sheetSelect.innerHTML += `<option value="${name}">${name}</option>`;
        });

        setStatus("✅ File siap di scan");
    };

    reader.readAsArrayBuffer(file);
}

// ==========================
// PARSE NUMBER
// ==========================
function parseNumber(val){
    if(!val) return 0;
    return Number(String(val).replace(/[^0-9]/g,'')) || 0;
}

// ==========================
// SCAN DATA
// ==========================
function applyFilter(){

    console.log("🔍 SCAN DIKLIK");

    if(!currentWorkbook){
        alert("Upload file dulu!");
        return;
    }

    setStatus("🔍 Scan data...");

    const sheetName = document.getElementById('sheetSelect').value;

    const sheet = currentWorkbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: ""
    });

    allData = [];

    rows.forEach(row => {

        // 🔥 SESUAI FILE KAMU
        const keterangan = row[2];   // C
        const invoice = row[3];      // D
        const dpp = row[4];          // E
        const tglBayar = row[13];    // N
        const pembayaran = row[14];  // O

        if(!keterangan || !invoice) return;

        // skip header
        if(String(keterangan).toLowerCase().includes("keterangan")) return;

        const dppNum = parseNumber(dpp);
        const bayarNum = parseNumber(pembayaran);

        allData.push({
            keterangan: String(keterangan),
            invoice: String(invoice),
            dpp: dppNum,
            tglBayar: tglBayar || "-",
            pembayaran: bayarNum,
            sisa: dppNum - bayarNum
        });

    });

    lastFiltered = allData;

    renderTable(allData);

    setStatus(`✅ Selesai. Data: ${allData.length}`);
}

// ==========================
// RENDER TABLE
// ==========================
function renderTable(data){

    let html = "";
    let total = 0;

    data.forEach(d=>{
        total += d.dpp;

        const warna = d.sisa > 0 ? 'style="color:red;"' : '';

        html += `
        <tr>
            <td>${d.keterangan}</td>
            <td>${d.invoice}</td>
            <td>${d.dpp.toLocaleString()}</td>
            <td>${d.tglBayar}</td>
            <td>${d.pembayaran.toLocaleString()}</td>
            <td ${warna}>${d.sisa.toLocaleString()}</td>
        </tr>`;
    });

    document.getElementById('result').innerHTML =
        html || `<tr><td colspan="6">Tidak ada data</td></tr>`;

    document.getElementById('total').innerText =
        "Total DPP: " + total.toLocaleString();
}

// ==========================
// EXPORT EXCEL
// ==========================
function exportExcel(){

    console.log("⬇ EXPORT DIKLIK");

    if(lastFiltered.length === 0){
        alert("Tidak ada data!");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(lastFiltered);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "HASIL");

    XLSX.writeFile(wb, "hasil_treking.xlsx");
}

});
