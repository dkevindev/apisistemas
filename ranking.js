const ExcelJS = require('exceljs');
const path = require('path');

async function consultarRankingAtestados() {
    const workbookClientes = new ExcelJS.Workbook();
    const workbookAtestados = new ExcelJS.Workbook();
    const filenameClientes = path.join(__dirname, 'bd.xlsx');
    const filenameAtestados = path.join(__dirname, 'atestados.xlsx');

    await Promise.all([
        workbookClientes.xlsx.readFile(filenameClientes),
        workbookAtestados.xlsx.readFile(filenameAtestados)
    ]);

    const worksheetClientes = workbookClientes.getWorksheet('Clientes');
    const worksheetAtestados = workbookAtestados.getWorksheet('Atestados');

    const dadosClientes = [];
    const headerRowClientes = worksheetClientes.getRow(1).values.slice(1); // Remover a primeira coluna

    worksheetClientes.eachRow({ includeEmpty: false, skipHeader: true }, row => {
        const rowData = {};
        row.eachCell((cell, colNumber) => {
            rowData[headerRowClientes[colNumber - 1]] = cell.value; // Ajustar o índice para corresponder ao header
        });
        dadosClientes.push(rowData);
    });

    const dadosAtestados = [];
    const headerRowAtestados = worksheetAtestados.getRow(1).values.slice(1); // Remover a primeira coluna

    worksheetAtestados.eachRow({ includeEmpty: false, skipHeader: true }, row => {
        const rowData = {};
        row.eachCell((cell, colNumber) => {
            rowData[headerRowAtestados[colNumber - 1]] = cell.value; // Ajustar o índice para corresponder ao header
        });
        dadosAtestados.push(rowData);
    });

    const dadosCombinados = dadosClientes.map(cliente => {
        const atestados = dadosAtestados.find(atestado => atestado.QRA === cliente.QRA);
        return {
            nome: cliente.QRA,
            atestados: atestados ? atestados.atestados : 0
        };
    }).filter(item => item.nome !== 'QRA'); // Filtra a primeira linha do cabeçalho
    
    dadosCombinados.sort((a, b) => b.atestados - a.atestados);
    
    return dadosCombinados.map((item, index) => ({
        nome: item.nome,
        atestados: item.atestados
    }));
}

consultarRankingAtestados()
    .then(ranking => {
        console.log(JSON.stringify(ranking, null, 2)); // Exibe o ranking no formato JSON
    })
    .catch(console.error);
