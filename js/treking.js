let allData = [];
let lastFiltered = [];
let currentWorkbook = null;

// ==========================
document.getElementById('upload').addEventListener('change', handleFile);
document.getElementById('btnScan').addEventListener('click', applyFilter);
document.getElementById('btnExport').addEventListener('click', exportExcel);

// 🔥 AUTO FILTER (REALTIME TANPA SCAN)
document.getElementById('sheetSelect').addEventListener('change', autoFilter);
document.getElementById('tahunSelect').addEventListener('change', autoFilter);
document.querySelectorAll('.tipeCheck').forEach(cb => {
    cb.addEventListener('change', autoFilter);
});

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

        // 🔥 isi tahun
        extractYears(workbook);

        setStatus("✅ File siap di scan");
    };

    reader.readAsArrayBuffer(file);
}

// ==========================
function extractYears(workbook){

    const tahunSet = new Set();

    workbook.SheetNames.forEach(sheetName=>{
        const sheet = workbook.Sheets[sheetName];

        const rows = XLSX.utils.sheet_to_json(sheet, {
            header:1,
            defval:""
        });

        rows.forEach(row=>{
            const periode = String(row[4] || "");

            const match = periode.match(/\b(20\d{2})\b/);
            if(match){
                tahunSet.add(match[1]);
            }
        });
    });

    const tahunSelect = document.getElementById('tahunSelect');
    tahunSelect.innerHTML = '<option value="">-- PILIH TAHUN --</option>';

    [...tahunSet].sort().forEach(t=>{
        tahunSelect.innerHTML += `<option value="${t}">${t}</option>`;
    });
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
            const totalProforma = row[10];
            const totalInvoice = row[11];

            const tglBayar = formatTanggal(row[13]);
            const pembayaran = row[14];

            if(!kota || !periode || !invoice) return;
            if(String(kota).toLowerCase() === "kota") return;

            const dppNum = parseNumber(dpp);
            if(dppNum === 0) return;

            const bayarNum = parseNumber(pembayaran);

            const sheetLower = sheetName.toLowerCase();
            const isProforma = sheetLower.includes("proforma");
            const isInvoice = sheetLower.includes("invoice");

            const ppnProforma = parseNumber(totalProforma);
            const ppnInvoice = parseNumber(totalInvoice);

            const totalProformaFix = ppnProforma ? dppNum + ppnProforma : 0;
            const totalInvoiceFix = ppnInvoice ? dppNum + ppnInvoice : 0;

            allData.push({
                sheet: sheetName,
                kota: String(kota),
                periode: String(periode),
                invoice: String(invoice),
                dpp: dppNum,
                totalProforma: isProforma ? totalProformaFix : 0,
                totalInvoice: isInvoice ? totalInvoiceFix : 0,
                tglBayar: tglBayar,
                pembayaran: bayarNum
            });

        });

    });
}

// ==========================
function autoFilter(){
    if(!currentWorkbook) return;

    const sheet = document.getElementById('sheetSelect').value;

    processWorkbook(currentWorkbook, sheet);

    applyFilter(true); // 🔥 tanpa scan ulang berat
}

// ==========================
function applyFilter(isAuto = false){

    if(!currentWorkbook){
        alert("Upload file dulu!");
        return;
    }

    if(!isAuto){
        setStatus("🔍 Scan data...");
    }

    const kotaKey = document.getElementById('kotaInput').value.toLowerCase();
    const periodeKey = document.getElementById('periodeInput').value.toLowerCase();
    const tahun = document.getElementById('tahunSelect').value;

    // 🔥 checkbox
    const tipeChecked = Array.from(document.querySelectorAll('.tipeCheck:checked'))
        .map(el => el.value);

    const filtered = allData.filter(d => {

        // 🔥 filter kota
        if(kotaKey && !d.kota.toLowerCase().includes(kotaKey)) return false;

        // 🔥 filter periode
        if(periodeKey && !d.periode.toLowerCase().includes(periodeKey)) return false;

        // 🔥 filter tahun (HANYA JIKA DIPILIH)
        if(tahun){
            if(!d.periode.includes(tahun)) return false;
        }

        // 🔥 filter tipe (HANYA JIKA ADA YANG DICENTANG)
        if(tipeChecked.length > 0){
            const isProforma = d.totalProforma > 0;
            const isInvoice = d.totalInvoice > 0;

            if(tipeChecked.includes("proforma") && isProforma) return true;
            if(tipeChecked.includes("invoice") && isInvoice) return true;

            return false;
        }

        return true;
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
            <td style="text-align:right">${d.dpp.toLocaleString()}</td>
            <td style="text-align:right">${d.totalProforma ? d.totalProforma.toLocaleString() : '-'}</td>
            <td style="text-align:right">${d.totalInvoice ? d.totalInvoice.toLocaleString() : '-'}</td>
            <td>${d.tglBayar}</td>
            <td style="text-align:right">${d.pembayaran.toLocaleString()}</td>
        </tr>`;
    });

    document.getElementById('result').innerHTML =
        html || `<tr><td colspan="8" style="text-align:center;">Tidak ada data</td></tr>`;

    document.getElementById('total').innerText =
        "Total DPP: " + total.toLocaleString();
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
        DPP: d.dpp,
        Total_Proforma: d.totalProforma,
        Total_Invoice: d.totalInvoice,
        Tgl_Bayar: d.tglBayar,
        Pembayaran: d.pembayaran
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "HASIL");

    XLSX.writeFile(wb, "hasil_treking.xlsx");
}
