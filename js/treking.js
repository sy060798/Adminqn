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

// NORMALIZE
function normalize(str){
    return String(str || "")
        .toLowerCase()
        .replace(/\s+/g,'')
        .replace(/[^a-z0-9]/g,'');
}

// PARSE ANGKA (FIX DPP 0)
function parseNumber(val){
    if(!val) return 0;

    let str = String(val)
        .replace(/[^0-9,-]/g,'') // hapus huruf
        .replace(/,/g,'');       // hapus koma

    return parseFloat(str) || 0;
}

// AUTO DETECT KOLOM
function detectColumn(row, targets){
    const keys = Object.keys(row);

    for(let key of keys){
        const cleanKey = normalize(key);

        for(let t of targets){
            if(cleanKey.includes(normalize(t))){
                return key;
            }
        }
    }
    return null;
}

// HANDLE FILE
function handleFile(e){
    const file = e.target.files[0];
    if(!file) return;

    setStatus("⏳ Membaca file Excel...");

    const reader = new FileReader();

    reader.onload = function(evt){
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type:'array'});

        processWorkbook(workbook);
        populateSheets();
        populateFilterDropdowns();

        setStatus("✅ File siap. Silakan filter data.");
    };

    reader.readAsArrayBuffer(file);
}

// PROSES DATA
function processWorkbook(workbook){
    allData = [];

    workbook.SheetNames.forEach(sheetName => {

        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet);

        if(json.length === 0) return;

        const sample = json[0];

        const colKota = detectColumn(sample, ["kota","area","regional"]);
        const colPeriode = detectColumn(sample, ["periode","bulan"]);
        const colInvoice = detectColumn(sample, ["invoice","inv"]);
        const colDpp = detectColumn(sample, ["dpp","amount","total","nilai"]);

        console.log("DETECT:", {sheetName, colKota, colPeriode, colInvoice, colDpp});

        json.forEach(row => {
            allData.push({
                sheet: sheetName,
                kota: colKota ? row[colKota] : "",
                periode: colPeriode ? row[colPeriode] : "",
                invoice: colInvoice ? row[colInvoice] : "",
                dpp: colDpp ? parseNumber(row[colDpp]) : 0
            });
        });

    });

    console.log("ALL DATA:", allData);
}

// SHEET DROPDOWN
function populateSheets(){
    const sheetSet = [...new Set(allData.map(d => d.sheet))];
    const el = document.getElementById('sheetSelect');

    el.innerHTML = '<option value="">-- PILIH SHEET --</option>';

    sheetSet.forEach(s=>{
        el.innerHTML += `<option value="${s}">${s}</option>`;
    });
}

// 🔥 AUTO DROPDOWN KOTA & PERIODE
function populateFilterDropdowns(){

    const kotaSet = [...new Set(allData.map(d => d.kota).filter(Boolean))];
    const periodeSet = [...new Set(allData.map(d => d.periode).filter(Boolean))];

    const kotaInput = document.getElementById('kotaInput');
    const periodeInput = document.getElementById('periodeInput');

    // bikin datalist
    let kotaList = document.getElementById('kotaList');
    let periodeList = document.getElementById('periodeList');

    if(!kotaList){
        kotaList = document.createElement("datalist");
        kotaList.id = "kotaList";
        document.body.appendChild(kotaList);
        kotaInput.setAttribute("list","kotaList");
    }

    if(!periodeList){
        periodeList = document.createElement("datalist");
        periodeList.id = "periodeList";
        document.body.appendChild(periodeList);
        periodeInput.setAttribute("list","periodeList");
    }

    kotaList.innerHTML = kotaSet.map(k => `<option value="${k}">`).join("");
    periodeList.innerHTML = periodeSet.map(p => `<option value="${p}">`).join("");
}

// FILTER
function applyFilter(){

    setStatus("🔍 Scanning data...");

    const sheet = document.getElementById('sheetSelect').value;
    const kotaKey = document.getElementById('kotaInput').value.toLowerCase();
    const periodeKey = document.getElementById('periodeInput').value.toLowerCase();

    const filtered = allData.filter(d => {

        const matchSheet = !sheet || d.sheet === sheet;

        const kotaVal = String(d.kota || "").toLowerCase();
        const periodeVal = String(d.periode || "").toLowerCase();

        const matchKota = !kotaKey || kotaVal.includes(kotaKey);
        const matchPeriode = !periodeKey || periodeVal.includes(periodeKey);

        return matchSheet && matchKota && matchPeriode;
    });

    lastFiltered = filtered;

    renderTable(filtered);

    setStatus(`✅ Selesai. Ditemukan ${filtered.length} data`);
}

// RENDER
function renderTable(data){

    let html = "";
    let total = 0;

    data.forEach(d=>{
        total += d.dpp;

        html += `
        <tr>
            <td>${d.kota || '-'}</td>
            <td>${d.periode || '-'}</td>
            <td>${d.invoice || '-'}</td>
            <td>${d.dpp.toLocaleString()}</td>
        </tr>`;
    });

    document.getElementById('result').innerHTML = html || 
        `<tr><td colspan="4" style="text-align:center;">Tidak ada data</td></tr>`;

    document.getElementById('total').innerText =
        "Total DPP: " + total.toLocaleString();
}

// EXPORT
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
