let allData = [];
let lastFiltered = [];
let currentWorkbook = null;

// ==========================
// INIT EVENT
// ==========================
document.getElementById('upload').addEventListener('change', handleFile);
document.getElementById('btnScan').addEventListener('click', applyFilter);
document.getElementById('btnExport').addEventListener('click', exportExcel);

// ==========================
// STATUS
// ==========================
function setStatus(msg){
    document.getElementById('status').innerText = msg;
}

// ==========================
// HANDLE FILE UPLOAD
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

        // 🔥 ISI DROPDOWN SHEET
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
// PARSE ANGKA (AMAN)
// ==========================
function parseNumber(val){
    if(!val) return 0;
    return Number(String(val).replace(/[^0-9]/g,'')) || 0;
}

// ==========================
// PROSES DATA
// ==========================
function processWorkbook(workbook, selectedSheet){

    allData = [];

    const sheets = selectedSheet ? [selectedSheet] : workbook.SheetNames;

    sheets.forEach(sheetName => {

        const sheet = workbook.Sheets[sheetName];

        const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: ""
        });

        rows.forEach(row => {

            // 🔥 POSISI KOLOM SESUAI FILE KAMU
            const kota = row[3];
            const periode = row[4];
            const invoice = row[6];
            const dpp = row[9];

            // skip baris tidak valid
            if(!kota || !periode || !invoice) return;

            // skip header
            if(String(kota).toLowerCase() === "kota") return;

            allData.push({
                sheet: sheetName,
                kota: String(kota),
                periode: String(periode),
                invoice: String(invoice),
                dpp: parseNumber(dpp)
            });

        });

    });

    console.log("DATA FINAL:", allData);
}

// ==========================
// FILTER + SCAN
// ==========================
function applyFilter(){

    const sheet = document.getElementById('sheetSelect').value;

    if(!currentWorkbook){
        alert("Upload file dulu!");
        return;
    }

    setStatus("🔍 Scan data...");

    // 🔥 ambil ulang data dari sheet
    processWorkbook(currentWorkbook, sheet);

    const kotaKey = document.getElementById('kotaInput').value.toLowerCase();
    const periodeKey = document.getElementById('periodeInput').value.toLowerCase();

    const filtered = allData.filter(d => {
        return (
            (!kotaKey || d.kota.toLowerCase().includes(kotaKey)) &&
            (!periodeKey || d.periode.toLowerCase().includes(periodeKey))
        );
    });

    lastFiltered = filtered;

    renderTable(filtered);

    setStatus(`✅ Selesai. Ditemukan ${filtered.length} data`);
}

// ==========================
// RENDER TABLE
// ==========================
function renderTable(data){

    let html = "";
    let total = 0;

    data.forEach(d=>{
        total += d.dpp;

        html += `
        <tr>
            <td>${d.kota}</td>
            <td>${d.periode}</td>
            <td>${d.invoice}</td>
            <td>${d.dpp.toLocaleString()}</td>
        </tr>`;
    });

    document.getElementById('result').innerHTML =
        html || `<tr><td colspan="4" style="text-align:center;">Tidak ada data</td></tr>`;

    document.getElementById('total').innerText =
        "Total DPP: " + total.toLocaleString();
}

// ==========================
// EXPORT EXCEL
// ==========================
function exportExcel(){

    if(lastFiltered.length === 0){
        alert("Tidak ada data untuk di export!");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(lastFiltered);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "HASIL");

    XLSX.writeFile(wb, "hasil_treking.xlsx");
}
