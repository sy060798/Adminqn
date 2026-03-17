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

    setStatus("⏳ Membaca file...");

    const reader = new FileReader();

    reader.onload = function(evt){
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type:'array'});

        processWorkbook(workbook);
        populateSheets();

        setStatus("✅ File loaded. Silakan scan data.");
    };

    reader.readAsArrayBuffer(file);
}

// fleksibel ambil kolom
function getValue(row, names){
    for(let n of names){
        if(row[n] !== undefined) return row[n];
    }
    return "";
}

// baca semua sheet
function processWorkbook(workbook){
    allData = [];

    workbook.SheetNames.forEach(sheetName => {

        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet);

        json.forEach(row => {
            allData.push({
                sheet: sheetName,
                kota: getValue(row, ["KOTA","Kota","LOKASI"]),
                periode: getValue(row, ["PERIODE","Periode","BULAN"]),
                invoice: getValue(row, ["NOMOR INVOICE","NO INV","NO INVOICE"]),
                dpp: parseFloat(getValue(row, ["DPP","TOTAL","JUMLAH"])) || 0
            });
        });

    });
}

// isi dropdown sheet
function populateSheets(){
    const sheetSet = [...new Set(allData.map(d => d.sheet))];
    const el = document.getElementById('sheetSelect');
    el.innerHTML = '<option value="">-- PILIH SHEET --</option>';

    sheetSet.forEach(s=>{
        el.innerHTML += `<option value="${s}">${s}</option>`;
    });
}

// filter keyword
function applyFilter(){

    setStatus("🔍 Scanning data...");

    const sheet = document.getElementById('sheetSelect').value;
    const kotaKey = document.getElementById('kotaInput').value.toLowerCase();
    const periodeKey = document.getElementById('periodeInput').value.toLowerCase();

    const filtered = allData.filter(d => {

        const matchSheet = !sheet || d.sheet === sheet;

        const matchKota = !kotaKey || 
            (d.kota && d.kota.toLowerCase().includes(kotaKey));

        const matchPeriode = !periodeKey || 
            (d.periode && d.periode.toLowerCase().includes(periodeKey));

        return matchSheet && matchKota && matchPeriode;
    });

    lastFiltered = filtered;

    renderTable(filtered);

    setStatus(`✅ Selesai. Ditemukan ${filtered.length} data`);
}

// render
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

    document.getElementById('result').innerHTML = html || 
        `<tr><td colspan="4" align="center">Tidak ada data</td></tr>`;

    document.getElementById('total').innerText =
        "Total DPP: " + total.toLocaleString();
}

// export
function exportExcel(){

    if(lastFiltered.length === 0){
        alert("Tidak ada data untuk di export");
        return;
    }

    const exportData = lastFiltered.map(d => ({
        KOTA: d.kota,
        PERIODE: d.periode,
        "NOMOR INVOICE": d.invoice,
        DPP: d.dpp
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "HASIL TREKING");

    XLSX.writeFile(wb, "hasil_treking.xlsx");
}
