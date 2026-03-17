let allData = [];
let lastFiltered = [];

document.getElementById('upload').addEventListener('change', handleFile);
document.getElementById('btnScan').addEventListener('click', applyFilter);
document.getElementById('btnExport').addEventListener('click', exportExcel);

function setStatus(msg){
    document.getElementById('status').innerText = msg;
}

function handleFile(e){
    const file = e.target.files[0];
    if(!file) return;

    const reader = new FileReader();

    reader.onload = function(evt){
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type:'array'});

        processWorkbook(workbook);

        setStatus("✅ File siap di scan");
    };

    reader.readAsArrayBuffer(file);
}

// 🔥 PARSE ANGKA AMAN
function parseNumber(val){
    if(!val) return 0;
    return Number(String(val).replace(/[^0-9]/g,'')) || 0;
}

// 🔥 PROSES UTAMA (FIX SESUAI FILE KAMU)
function processWorkbook(workbook){

    allData = [];

    workbook.SheetNames.forEach(sheetName => {

        // hanya ambil sheet invoice
        if(!sheetName.toLowerCase().includes("invoice")) return;

        const sheet = workbook.Sheets[sheetName];

        const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: ""
        });

        rows.forEach(row => {

            // 🔥 ambil berdasarkan posisi kolom
            const kota = row[3];
            const periode = row[4];
            const invoice = row[6];
            const dpp = row[9];

            // skip kalau bukan data
            if(!kota || !periode || !invoice) return;

            allData.push({
                sheet: sheetName,
                kota: String(kota),
                periode: String(periode),
                invoice: String(invoice),
                dpp: parseNumber(dpp)
            });

        });

    });

    console.log("DATA FIX:", allData);

    renderTable(allData);
}

// FILTER
function applyFilter(){

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

    setStatus(`✅ ${filtered.length} data ditemukan`);
}

// RENDER
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
        html || `<tr><td colspan="4" align="center">Tidak ada data</td></tr>`;

    document.getElementById('total').innerText =
        "Total DPP: " + total.toLocaleString();
}

// EXPORT
function exportExcel(){

    if(lastFiltered.length === 0){
        alert("Tidak ada data");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(lastFiltered);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "HASIL");

    XLSX.writeFile(wb, "hasil.xlsx");
}
