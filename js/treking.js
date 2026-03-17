function processWorkbook(workbook){

    allData = [];

    workbook.SheetNames.forEach(sheetName => {

        const sheet = workbook.Sheets[sheetName];

        // 🔥 AMBIL SEMUA DALAM BENTUK ARRAY (BUKAN JSON)
        const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: ""
        });

        if(rows.length < 3) return;

        console.log("RAW ROWS:", rows.slice(0,5));

        // 🔥 CARI BARIS HEADER OTOMATIS
        let headerRowIndex = -1;

        for(let i=0; i<rows.length; i++){
            const row = rows[i].join(" ").toLowerCase();

            if(
                row.includes("kota") &&
                row.includes("periode") &&
                row.includes("invoice")
            ){
                headerRowIndex = i;
                break;
            }
        }

        if(headerRowIndex === -1){
            console.log("❌ Header tidak ditemukan di sheet:", sheetName);
            return;
        }

        console.log("✅ HEADER DI BARIS:", headerRowIndex);

        const headers = rows[headerRowIndex].map(h => clean(h));

        const colIndex = {
            kota: headers.findIndex(h => h.includes("kota")),
            periode: headers.findIndex(h => h.includes("periode")),
            invoice: headers.findIndex(h => h.includes("invoice")),
            dpp: headers.findIndex(h => h === "dpp")
        };

        console.log("COLUMN INDEX:", colIndex);

        // 🔥 AMBIL DATA SETELAH HEADER
        for(let i = headerRowIndex + 1; i < rows.length; i++){

            const row = rows[i];

            // skip kosong
            if(row.join("").trim() === "") continue;

            const data = {
                sheet: sheetName,
                kota: row[colIndex.kota] || "",
                periode: row[colIndex.periode] || "",
                invoice: row[colIndex.invoice] || "",
                dpp: parseNumber(row[colIndex.dpp])
            };

            allData.push(data);
        }

    });

    console.log("FINAL DATA:", allData);

    renderTable(allData);
}
