let entrega = 0.09;
let montaje = 0.14;

// Actualizar valores desde inputs
document.getElementById("entregaInput").addEventListener("input", (e) => {
    entrega = parseFloat(e.target.value);
});
document.getElementById("montajeInput").addEventListener("input", (e) => {
    montaje = parseFloat(e.target.value);
});

document.getElementById("processBtn").addEventListener("click", () => {
    const file1 = document.getElementById("file1").files[0];
    const file2 = document.getElementById("file2").files[0];

    if (!file1 || !file2) {
        alert("Por favor selecciona ambos archivos.");
        return;
    }

    // Leer los dos excels
    Promise.all([readExcel(file1), readExcel(file2)]).then(([data1, data2]) => {
        processFiles(data1, data2);
    });
});

function readExcel(file) {
    return new Promise((resolve) => {
        let reader = new FileReader();
        reader.onload = (e) => {
            let wb = XLSX.read(e.target.result, { type: "binary" });
            let sheet = wb.SheetNames[0];
            let data = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval: "" });
            resolve(data);
        };
        reader.readAsBinaryString(file);
    });
}

function processFiles(data1, data2) {

    // Crear mapa segundo excel: (Pedido + Código/Referencia) -> Importe neto
    const precioMap = {};
    data2.forEach(row => {
        const key = (row["Pedido de ventas"] + "|" + row["Código de artículo"]).trim();
        precioMap[key] = parseFloat(row["Importe neto"]) || 0;
    });

    const result = data1.map(row => {
        const key = (row["Pedido de ventas"] + "|" + row["Artículo – Referencia"]).trim();
        const importeNeto = precioMap[key] || 0;

        let tarifa = parseFloat(row["Tarifa unit."]) || 0;

        if (row["Categoría"].includes("PREM")) {
            tarifa = importeNeto * montaje;
        } else if (row["Categoría"].includes("TIMA")) {
            tarifa = importeNeto * entrega;
        }

        const cantidad = parseFloat(row["Artículo – Cantidad"]) || 0;
        const total = tarifa * cantidad;

        // Redondeo a 2 decimales
        const tarifaRedondeada = Math.round(tarifa * 100) / 100;
        const totalRedondeado = Math.round(total * 100) / 100;
        
        return {
            "Fecha": row["Fecha"],
            "Expedidor": row["Expedidor"],
            "Transportista": row["Transportista"],
            "Identificador de la tarea": row["Identificador de la tarea"],
            "Cuenta del cliente": row["Cuenta del cliente"],
            "Pedido de ventas": row["Pedido de ventas"],
            "Artículo – Nombre": row["Artículo – Nombre"],
            "Artículo – Cantidad": cantidad,
            "Artículo – Referencia": row["Artículo – Referencia"],
            "Retirada": row["Retirada"],
            "Cruce": row["Cruce"],
            "Categoría": row["Categoría"],
            "Importe neto": importeNeto,
            "Tarifa unit.": tarifaRedondeada,
            "Total": totalRedondeado
        };
    });

    // Crear y descargar archivo
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(result);
    XLSX.utils.book_append_sheet(wb, ws, "Resultado");
    XLSX.writeFile(wb, "resultado.xlsx");
}
