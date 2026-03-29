let allData = [];
let lastFiltered = [];
let currentWorkbook = null;

// ==========================
document.getElementById('upload').addEventListener('change', handleFile);
document.getElementById('btnScan').addEventListener('click', applyFilter);
document.getElementById('btnExport').addEventListener('click', exportExcel);

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
function formatTanggal(val){
    if(!val || val === 0) return "-";

    if(typeof val === "number"){
        return XLSX.SSF.format("dd-mmm-yyyy", val);
    }

    return val;
}

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

            const kota = row[3];
            const periode = row[4];
            const invoice = row[6];

            const dpp = row[9];
            const totalProforma = row[10]; // ✅ KOLOM TOTAL PROFORMA (sesuaikan jika beda)

            const tglBayar = formatTanggal(row[13]);
            const pembayaran = row[14];

            if(!kota || !periode || !invoice) return;
            if(String(kota).toLowerCase() === "kota") return;

            let dppFix = 0;

            // 🔍 Deteksi jenis sheet
            const isProforma = sheetName.toLowerCase().includes("proforma");

            if(isProforma){
                // ✅ PROFORMA → ambil dari TOTAL PROFORMA
                const totalNum = parseNumber(totalProforma);
                if(totalNum === 0) return;

                dppFix = totalNum;

            } else {
                // ✅ INVOICE → DPP + PPN 11%
                const dppNum = parseNumber(dpp);
                if(dppNum === 0) return;

                dppFix = Math.round(dppNum * 1.11);
            }

            const bayarNum = parseNumber(pembayaran);

            allData.push({
                sheet: sheetName,
                kota: String(kota),
                periode: String(periode),
                invoice: String(invoice),
                dpp: dppFix,
                totalProforma: parseNumber(totalProforma), // ✅ tambahan
                tglBayar: tglBayar,
                pembayaran: bayarNum
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
            <td>${d.totalProforma ? d.totalProforma.toLocaleString() : '-'}</td>
            <td>${d.tglBayar}</td>
            <td>${d.pembayaran.toLocaleString()}</td>
        </tr>`;
    });

    document.getElementById('result').innerHTML =
        html || `<tr><td colspan="7" style="text-align:center;">Tidak ada data</td></tr>`;

    document.getElementById('total').innerText =
        "Total DPP (Include PPN / Proforma): " + total.toLocaleString();
}

// ==========================
function exportExcel(){

    if(lastFiltered.length === 0){
        alert("Tidak ada data untuk di export!");
        return;
    }

    const exportData = lastFiltered.map(d => ({
        Kota: d.kota,
        Periode: d.periode,
        Invoice: d.invoice,
        DPP_Include_PPN: d.dpp,
        Total_Proforma: d.totalProforma,
        Tgl_Bayar: d.tglBayar,
        Pembayaran: d.pembayaran
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "HASIL");

    XLSX.writeFile(wb, "hasil_treking.xlsx");
}
