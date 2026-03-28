document.addEventListener("DOMContentLoaded", function(){

let allData = [];
let lastFiltered = [];
let currentWorkbook = null;

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
            const sisaExcel = row[18];   // S

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
                sisa: dppNum - bayarNum // pakai hitung otomatis
                // kalau mau pakai excel: parseNumber(sisaExcel)
            });

        });

    });

    console.log("DATA FINAL:", allData);
}

// ==========================
function applyFilter(){

    const sheet = document.getElementById('sheetSelect').value;

    if(!currentWorkbook){
        alert("Upload file dulu!");
        return;
    }

    setStatus("🔍 Scan data...");

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

    const result = document.getElementById('result');
    if(result){
        result.innerHTML = html || `<tr><td colspan="6">Tidak ada data</td></tr>`;
    }

    const totalEl = document.getElementById('total');
    if(totalEl){
        totalEl.innerText = "Total DPP: " + total.toLocaleString();
    }
}

// ==========================
function exportExcel(){

    if(lastFiltered.length === 0){
        alert("Tidak ada data!");
        return;
    }

    const exportData = lastFiltered.map(d => ({
        Keterangan: d.keterangan,
        Invoice: d.invoice,
        DPP: d.dpp,
        Tgl_Pembayaran: d.tglBayar,
        Pembayaran: d.pembayaran,
        Sisa: d.sisa
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "HASIL");

    XLSX.writeFile(wb, "hasil_treking.xlsx");
}

});
