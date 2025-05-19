// <script src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>

function exportDynamicsQueryToExcel(queryResult, fileName = "Relatorio_Dynamics.xlsx") {
    if (!queryResult || !Array.isArray(queryResult.value) || queryResult.value.length === 0) {
        alert("Nenhum dado encontrado para exportação.");
        return;
    }

    // Extrai colunas dinamicamente
    const rows = queryResult.value.map(item => {
        const row = {};
        for (const key in item) {
            // Ignora propriedades de metadata
            if (!key.startsWith("@") && !key.includes("odata.")) {
                row[key] = item[key];
            }
        }
        return row;
    });

    // Cria uma planilha com SheetJS
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Dados");

    // Converte para um blob e força o download
    const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

Xrm.WebApi.retrieveMultipleRecords("account", "?$select=name,telephone1,revenue").then(function(result) {
    exportDynamicsQueryToExcel(result, "Contas.xlsx");
});
// Exemplo de uso