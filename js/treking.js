let allData = [];
let lastFiltered = [];

// EVENT
document.getElementById('upload').addEventListener('change', handleFile);
document.getElementById('btnScan').addEventListener('click', applyFilter);
document.getElementById('btnExport').addEventListener('click', exportExcel);

// STATUS
function setStatus(msg){
    document.getElementById('status').innerText = msg;
}

// HANDLE FILE
function handleFile(e){
    const file = e.target.files[0];
    if(!file) return;

    setStatus("⏳ Membaca file...");

    const reader = new FileReader();

    reader.onload = function(evt){
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type:'array'});

        processWorkbook(workbook);

        setStatus("✅ File terbaca. Klik SCAN DATA");
    };

    reader.readAsArrayBuffer(file);
}

// NORMALIZE
function clean(str){
    return String(str || "")
        .toLowerCase()
        .replace(/\s+/g,'')
        .replace(/[^a-z0-9]/g,'');
}

// PARSE ANGKA (SUPER AMAN)
function parseNumber(val){
    if(!val) return 0;

    let str = String(val);

    // hapus semua kecuali angka
    str = str.replace(/[^0-9]/g,'');

    return parseInt(str) || 0;
}

// PROSES WORKBOOK (PAKSA AMBIL DATA)
function processWorkbook(workbook){

    allData = [];

    workbook.SheetNames.forEach(sheetName => {

        const sheet = workbook.Sheets[sheetName];

        // 🔥 PAKSA BACA RAW
        const json = XLSX.utils.sheet_to_json(sheet, {defval:""});

        if(json.length === 0) return;

        console.log("SAMPLE ROW:", json[0]);

        const keys = Object.keys(json[0]);

        // 🔥 DETECT MANUAL
        let colKota = null;
        let colPeriode = null;
        let colInvoice = null;
        let colDpp = null;

        keys.forEach(k=>{
            const c = clean(k);

            if(c.includes("kota")) colKota = k;
            if(c.includes("periode")) colPeriode = k;
            if(c.includes("invoice")) colInvoice = k;
            if(c === "dpp") colDpp = k;
        });

        console.log("DETECTED:", {colKota, colPeriode, colInvoice, colDpp});

        json.forEach(row => {

            const data = {
                sheet: sheetName,
                kota: row[colKota] || "",
                periode: row[colPeriode] || "",
                invoice: row[colInvoice] || "",
                dpp: parseNumber(row[colDpp])
            };

            allData.push(data);
        });

    });

    console.log("ALL DATA FINAL:", allData);

    // 🔥 LANGSUNG TAMPILKAN TANPA FILTER
    renderTable(allData);
}

// FILTER (sementara sederhana)
function applyFilter(){

    setStatus("🔍 Scan...");

    const kotaKey = document.getElementById('kotaInput').value.toLowerCase();
    const periodeKey = document.getElementById('periodeInput').value.toLowerCase();

    const filtered = allData.filter(d => {

        const kotaVal = String(d.kota).toLowerCase();
        const periodeVal = String(d.periode).toLowerCase();

        return (
            (!kotaKey || kotaVal.includes(kotaKey)) &&
            (!periodeKey || periodeVal.includes(periodeKey))
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
        html || `<tr><td colspan="4" style="text-align:center;">Tidak ada data</td></tr>`;

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
