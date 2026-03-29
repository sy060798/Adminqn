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
            const totalProforma = row[10];
            const totalInvoice = row[11];

            const tglBayar = formatTanggal(row[13]);
            const pembayaran = row[14];

            if(!kota || !periode || !invoice) return;
            if(String(kota).toLowerCase() === "kota") return;

            const dppNum = parseNumber(dpp);
            if(dppNum === 0) return;

            const bayarNum = parseNumber(pembayaran);
            const isProforma = sheetName.toLowerCase().includes("proforma");

            // 🔥 hitung total (DPP + PPN)
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
                totalInvoice: !isProforma ? totalInvoiceFix : 0,
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

    // 🔥 FILTER DARI CHECKBOX (PROFORMA / INVOICE)
    const selectedTipe = Array.from(document.querySelectorAll('.tipeCheck:checked'))
        .map(el => el.value);

    const filtered = allData.filter(d => {

        const matchTipe =
            selectedTipe.length === 0 ||
            (selectedTipe.includes("proforma") && d.totalProforma > 0) ||
            (selectedTipe.includes("invoice") && d.totalInvoice > 0);

        return (
            (!kotaKey || d.kota.toLowerCase().includes(kotaKey)) &&
            (!periodeKey || d.periode.toLowerCase().includes(periodeKey)) &&
            matchTipe
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
