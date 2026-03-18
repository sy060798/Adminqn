let processedData = [];

// NORMALIZE
function normalize(text){
    return String(text).toLowerCase().trim();
}

// PRECON
function getPrecon(row){
    const preconMap = {
        "kabel precon 35 old": "PRECON - 35 M",
        "kabel precon 50 old": "PRECON - 50 M",
        "kabel precon 75 old": "PRECON - 75 M",
        "kabel precon 80 old": "PRECON - 80 M",
        "kabel precon 100 old": "PRECON - 100 M",
        "kabel precon 125 old": "PRECON - 125 M",
        "kabel precon 150 old": "PRECON - 150 M",
        "kabel precon 175 old": "PRECON - 175 M",
        "kabel precon 200 old": "PRECON - 200 M",
        "kabel precon 225 old": "PRECON - 225 M",
        "kabel precon 250 old": "PRECON - 250 M"
    };

    let result = [];

    for(let key in row){
        let k = normalize(key);
        if(preconMap[k] && (row[key] == 1 || row[key] == "1")){
            result.push(preconMap[k]);
        }
    }

    return result.join(", ");
}

// AMBIL KOLOM FLEXIBLE
function getColumn(row, keyword){
    keyword = normalize(keyword);
    for(let key in row){
        if(normalize(key).includes(keyword)){
            return row[key];
        }
    }
    return "";
}

// PROCESS
function processExcel(){

    const file = document.getElementById("excelFile").files[0];

    if(!file){
        alert("Upload Excel dulu!");
        return;
    }

    const reader = new FileReader();

    reader.onload = function(e){

        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data,{type:"array"});
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet,{defval:""});

        const tbody = document.querySelector("#resultTable tbody");

        tbody.innerHTML = "";
        processedData = [];

        json.forEach(row=>{

            let dispatch = normalize(getColumn(row,"dispatch"));
            if(dispatch !== "done") return;

            const result = {

                dispatch:"Done",
                status:"Done",

                id:getColumn(row,"id"),
                wo:getColumn(row,"wo"),
                customer:getColumn(row,"customer"),

                tanggal:getColumn(row,"tanggal"),
                alamat:getColumn(row,"alamat"),
                cabang:getColumn(row,"cabang"),

                new_ont:getColumn(row,"ont baru"),
                old_ont:getColumn(row,"ont lama"),
                splicing:getColumn(row,"splicing"),
                rfo:getColumn(row,"rfo"),
                action:getColumn(row,"action"),

                precon:getPrecon(row),
                report:getColumn(row,"report")
            };

            processedData.push(result);

            const tr = document.createElement("tr");

            tr.innerHTML = `
            <td>${result.dispatch}</td>
            <td style="color:green;font-weight:bold">${result.status}</td>
            <td>${result.id}</td>
            <td>${result.wo}</td>
            <td>${result.customer}</td>
            <td>${result.tanggal}</td>
            <td>${result.alamat}</td>
            <td>${result.cabang}</td>
            <td>${result.new_ont}</td>
            <td>${result.old_ont}</td>
            <td>${result.splicing}</td>
            <td>${result.rfo}</td>
            <td>${result.action}</td>
            <td>${result.precon}</td>
            <td>${result.report}</td>
            `;

            tbody.appendChild(tr);

        });

    };

    reader.readAsArrayBuffer(file);
}

// DOWNLOAD
function downloadExcel(){

    if(processedData.length === 0){
        alert("Belum ada data!");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(processedData,{
        header:[
            "dispatch","status","id","wo","customer","tanggal",
            "alamat","cabang","new_ont","old_ont",
            "splicing","rfo","action","precon","report"
        ]
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Summary MT");

    XLSX.writeFile(wb,"summary_mt_done.xlsx");
}
