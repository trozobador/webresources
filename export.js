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



function getLoggedUserBusinessUnit() {
    return new Promise(function (resolve, reject) {
        var userId = Xrm.Utility.getGlobalContext().userSettings.userId;
        userId = userId.replace("{", "").replace("}", "");

        Xrm.WebApi.online.retrieveRecord("systemuser", userId, "?$select=fullname&$expand=businessunitid($select=name)")
            .then(function (result) {
                if (result.hasOwnProperty("businessunitid") && result.businessunitid && result.businessunitid.name) {
                    resolve(result.businessunitid.name);
                } else {
                    reject("Business Unit not found for the logged user.");
                }
            })
            .catch(function (error) {
                reject(error.message);
            });
    });
}


getLoggedUserBusinessUnit().then(function (businessUnitName) {
    console.log("Business Unit do usuário logado:", businessUnitName);
}).catch(function (error) {
    console.error("Erro ao obter Business Unit:", error);
});





function getBusinessUnitIdSync() {
    try {
        var baseUrl = Xrm.Utility.getGlobalContext().getClientUrl();
        var url = baseUrl + "/api/data/v9.2/WhoAmI";

        var request = new XMLHttpRequest();
        request.open("GET", url, false); // false = requisição síncrona
        request.setRequestHeader("Accept", "application/json");
        request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        request.setRequestHeader("OData-MaxVersion", "4.0");
        request.setRequestHeader("OData-Version", "4.0");

        request.send();

        if (request.status === 200) {
            var response = JSON.parse(request.responseText);
            return response.BusinessUnitId;
        } else {
            console.error("Erro na requisição WhoAmI: " + request.status + " - " + request.statusText);
            return null;
        }
    } catch (e) {
        console.error("Exceção ao obter BusinessUnitId: " + e.message);
        return null;
    }
}
