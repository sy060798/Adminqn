let allData = [];

document.getElementById('upload').addEventListener('change', handleFile);

function handleFile(e){
    const file = e.target.files[0];
    if(!file) return;

    const reader = new FileReader();

    reader.onload = function(evt){
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type:'array'});

        processWorkbook(workbook);
        populateFilters();
    };

    reader.readAsArrayBuffer(file);
}

// ambil value fleksibel
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
                kota: getValue(row, ["KOTA","Kota","LOKASI","CITY"]),
                periode: getValue(row, ["PERIODE","Periode","BULAN"]),
                invoice: getValue(row, ["NOMOR INVOICE","NO INV","NO INVOICE"]),
                dpp: parseFloat(getValue(row, ["DPP","TOTAL","JUMLAH"])) || 0
            });
        });

    });

    alert("Excel berhasil dimuat: " + allData.length + " data");
}

// isi dropdown
function populateFilters(){

    const sheetSet = [...new Set(allData.map(d => d.sheet))];
    const sheetSelect = document.getElementById('sheetSelect');
    sheetSelect.innerHTML = '<option value="">-- PILIH SHEET --</option>';

    sheetSet.forEach(s => {
        sheetSelect.innerHTML += `<option value="${s}">${s}</option>`;
    });

    const kotaSet = [...new Set(allData.map(d => d.kota).filter(Boolean))];
    const kotaSelect = document.getElementById('kotaSelect');
    kotaSelect.innerHTML = "";

    kotaSet.forEach(k => {
        kotaSelect.innerHTML += `<option value="${k}">${k}</option>`;
    });

    const periodeSet = [...new Set(allData.map(d => d.periode).filter(Boolean))];
    const periodeSelect = document.getElementById('periodeSelect');
    periodeSelect.innerHTML = "";

    periodeSet.forEach(p => {
        periodeSelect.innerHTML += `<option value="${p}">${p}</option>`;
    });
}

// filter data
function applyFilter(){

    const sheet = document.getElementById('sheetSelect').value;

    const kota = Array.from(document.getElementById('kotaSelect').selectedOptions).map(o => o.value);

    const periode = Array.from(document.getElementById('periodeSelect').selectedOptions).map(o => o.value);

    const filtered = allData.filter(d => {
        return (
            (!sheet || d.sheet === sheet) &&
            (kota.length === 0 || kota.includes(d.kota)) &&
            (periode.length === 0 || periode.includes(d.periode))
        );
    });

    renderTable(filtered);
}

// render tabel
function renderTable(data){

    let html = "";
    let total = 0;

    data.forEach(d => {
        total += d.dpp;

        html += `
        <tr>
            <td>${d.kota}</td>
            <td>${d.periode}</td>
            <td>${d.invoice}</td>
            <td>${d.dpp.toLocaleString()}</td>
        </tr>
        `;
    });

    document.getElementById('result').innerHTML = html;
    document.getElementById('total').innerText =
        "Total DPP: " + total.toLocaleString() + " | Jumlah Data: " + data.length;
}

// export excel
function exportExcel(){

    const sheet = document.getElementById('sheetSelect').value;

    const kota = Array.from(document.getElementById('kotaSelect').selectedOptions).map(o => o.value);

    const periode = Array.from(document.getElementById('periodeSelect').selectedOptions).map(o => o.value);

    const filtered = allData.filter(d => {
        return (
            (!sheet || d.sheet === sheet) &&
            (kota.length === 0 || kota.includes(d.kota)) &&
            (periode.length === 0 || periode.includes(d.periode))
        );
    });

    if(filtered.length === 0){
        alert("Tidak ada data untuk di export");
        return;
    }

    const exportData = filtered.map(d => ({
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
