let allData = [];
let lastFiltered = [];
let currentWorkbook = null;

// ==========================
document.getElementById('upload').addEventListener('change', handleFile);
document.getElementById('btnScan').addEventListener('click', applyFilter);

// reset biar tidak ngunci
document.getElementById('sheetSelect').addEventListener('change', () => {
    document.getElementById('tahunSelect').value = "";
    setStatus("⚠️ Filter berubah, klik SCAN DATA");
});

document.getElementById('tahunSelect').addEventListener('change', () => {
    document.getElementById('sheetSelect').value = "";
    setStatus("⚠️ Filter berubah, klik SCAN DATA");
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

        // sheet list
        const sheetSelect = document.getElementById('sheetSelect');
        sheetSelect.innerHTML = '<option value="">-- PILIH SHEET --</option>';

        workbook.SheetNames.forEach(name=>{
            sheetSelect.innerHTML += `<option value="${name}">${name}</option>`;
        });

        // tahun otomatis
        extractYears(workbook);

        setStatus("✅ File siap di scan");
    };

    reader.readAsArrayBuffer(file);
}

// ==========================
// 🔥 FIX ANGKA (TIDAK KURANG NOL)
function parseNumber(val){
    if(!val) return 0;

    val = String(val).trim();
    val = val.replace(/\./g, '').replace(/,/g, '');

    return Number(val) || 0;
}

// ==========================
// 💰 FORMAT RUPIAH
function formatRupiah(num){
    if(!num) return "-";
    return "Rp " + num.toLocaleString("id-ID");
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
// 🔥 AMBIL TAHUN DARI EXCEL
function extractYears(workbook){
    let years = new Set();

    workbook.SheetNames.forEach(sheetName=>{
        const sheet = workbook.Sheets[sheetName];

        const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: ""
        });

        rows.forEach(row=>{
            const periode = row[4];
            if(!periode) return;

            const match = String(periode).match(/\b(20\d{2})\b/);
            if(match){
                years.add(match[1]);
            }
        });
    });

    const tahunSelect = document.getElementById('tahunSelect');
    tahunSelect.innerHTML = '<option value="">-- PILIH TAHUN --</option>';

    [...years].sort().forEach(y=>{
        tahunSelect.innerHTML += `<option value="${y}">${y}</option>`;
    });
}

// ==========================
function processWorkbook(workbook, selectedSheet){

    allData = [];
    const seenInvoice = new Set();

    const sheets = selectedSheet ? [selectedSheet] : workbook.SheetNames;

    sheets.forEach(sheetName => {

        // 🔥 hanya ambil PROFORMA & INVOICE
        const lower = sheetName.toLowerCase();
        if(!lower.includes("proforma") && !lower.includes("invoice")) return;

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

            // 🔥 HAPUS DUPLIKAT INVOICE
            if(seenInvoice.has(invoice)) return;
            seenInvoice.add(invoice);

            const dppNum = parseNumber(dpp);
            if(dppNum === 0) return;

            const bayarNum = parseNumber(pembayaran);

            const isProforma = lower.includes("proforma");

            // 🔥 FIX PERFORMA (JANGAN TAMBAH DPP LAGI)
            const totalProformaFix = parseNumber(totalProforma);
            const totalInvoiceFix = parseNumber(totalInvoice);

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
}

// ==========================
function applyFilter(){

    if(!currentWorkbook){
        alert("Upload file dulu!");
        return;
    }

    setStatus("🔍 Scan data...");

    const sheet = document.getElementById('sheetSelect').value;
    const tahun = document.getElementById('tahunSelect').value;

    processWorkbook(currentWorkbook, sheet);

    const kotaKey = document.getElementById('kotaInput').value.toLowerCase();
    const periodeKey = document.getElementById('periodeInput').value.toLowerCase();

    const tipeChecked = [...document.querySelectorAll('.tipeCheck:checked')]
        .map(el => el.value);

    const filtered = allData.filter(d => {

        const isProforma = d.totalProforma > 0;
        const isInvoice = d.totalInvoice > 0;

        return (
            (!kotaKey || d.kota.toLowerCase().includes(kotaKey)) &&
            (!periodeKey || d.periode.toLowerCase().includes(periodeKey)) &&
            (!tahun || d.periode.includes(tahun)) &&
            (
                (tipeChecked.includes("proforma") && isProforma) ||
                (tipeChecked.includes("invoice") && isInvoice)
            )
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
            <td style="text-align:right">${formatRupiah(d.dpp)}</td>
            <td style="text-align:right">${d.totalProforma ? formatRupiah(d.totalProforma) : '-'}</td>
            <td style="text-align:right">${d.totalInvoice ? formatRupiah(d.totalInvoice) : '-'}</td>
            <td>${d.tglBayar}</td>
            <td style="text-align:right">${formatRupiah(d.pembayaran)}</td>
        </tr>`;
    });

    document.getElementById('result').innerHTML =
        html || `<tr><td colspan="8" style="text-align:center;">Tidak ada data</td></tr>`;

    document.getElementById('total').innerText =
        "Total DPP: " + formatRupiah(total);
}
