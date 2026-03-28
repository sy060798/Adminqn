document.addEventListener("DOMContentLoaded", function(){

let allData = [];
let lastFiltered = [];
let currentWorkbook = null;

// ==========================
const upload = document.getElementById('upload');
const btnScan = document.getElementById('btnScan');
const btnExport = document.getElementById('btnExport');

// 🔥 FIX EVENT (ANTI GAGAL KLIK)
upload.addEventListener('change', handleFile);
btnScan.addEventListener('click', applyFilter);
btnExport.addEventListener('click', exportExcel);

// ==========================
function setStatus(msg){
    document.getElementById('status').innerText = msg;
}

// ==========================
function handleFile(e){
    const file = e.target.files[0];
    if(!file) return;

    setStatus("⏳ Membaca file...");

    const reader = new FileReader();

    reader.onload = function(evt){
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type:'array'});

        currentWorkbook = workbook;

        const sheetSelect = document.getElementById('sheetSelect');
        sheetSelect.innerHTML = '<option value="">-- PILIH SHEET --</option>';

        workbook.SheetNames.forEach(name=>{
            sheetSelect.innerHTML += `<option value="${name}">${name}</option>`;
        });

        setStatus("✅ File siap di scan");
    };

    reader.readAsArrayBuffer(file);
}

// ==========================
function parseNumber(val){
    if(!val) return 0;
    return Number(String(val).replace(/[^0-9]/g,'')) || 0;
}

// ==========================
function processWorkbook(workbook, selectedSheet){

    allData = [];

    const sheets = selectedSheet 
        ? [selectedSheet] 
        : workbook.SheetNames;

    sheets.forEach(sheetName => {

        const sheet = workbook.Sheets[sheetName];

        const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: ""
        });

        rows.forEach(row => {

            // 🔥 SESUAI EXCEL KAMU
            const keterangan = row[2];   // C
            const invoice = row[3];      // D
            const dpp = row[4];          // E
            const tglBayar = row[13];    // N
            const pembayaran = row[14];  // O

            if(!keterangan || !invoice) return;
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

    });

    console.log("DATA:", allData);
}

// ==========================
function applyFilter(){

    if(!currentWorkbook){
        alert("Upload file dulu!");
        return;
    }

    setStatus("🔍 Scan data...");

    const sheet = document.getElementById('sheetSelect').value;

    processWorkbook(currentWorkbook, sheet);

    const keyword = document.getElementById('keteranganInput').value.toLowerCase();

    const filtered = allData.filter(d => {
        return (!keyword || d.keterangan.toLowerCase().includes(keyword));
    });

    lastFiltered = filtered;

    renderTable(filtered);

    setStatus(`✅ Selesai. Ditemukan ${filtered.length} data`);
}

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
function exportExcel(){

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
