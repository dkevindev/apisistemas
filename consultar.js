const ExcelJS = require('exceljs');
const path = require('path');

const workbook = new ExcelJS.Workbook();
const filename = path.join(__dirname, 'bd.xlsx');


async function consultarDadosPorNome(parteNome) {
    await workbook.xlsx.readFile(filename);
    const worksheet = workbook.getWorksheet('Planilha1');

    if (!worksheet) {
        throw new Error("Planilha 'dados' não encontrada no arquivo Excel.");
    }

    const dados = [];
    const headerRow = worksheet.getRow(1).values;

    worksheet.eachRow({ includeEmpty: false }, row => {
        const rowData = {};
        row.eachCell((cell, colNumber) => {
            rowData[headerRow[colNumber]] = cell.value;
        });

        // Verifica se o valor da primeira coluna (nome) contém a parte do nome informada
        if (rowData[headerRow[1]].toLowerCase().includes(parteNome.toLowerCase())) {
            dados.push(rowData);
        }
    });

    return dados;
}

async function consultarRankingAtestados() {
    await workbook.xlsx.readFile(filename);
    const worksheet = workbook.getWorksheet('Planilha1');

    if (!worksheet) {
        throw new Error("Planilha 'dados' não encontrada no arquivo Excel.");
    }

    const dados = [];
    const headerRow = worksheet.getRow(1).values;

    worksheet.eachRow({ includeEmpty: false, skipHeader: true }, row => {
        const rowData = {};
        row.eachCell((cell, colNumber) => {
            rowData[headerRow[colNumber]] = cell.value;
        });

        // Verifica se o valor da coluna 'atestados' é um número
        if (!isNaN(rowData.atestados)) {
            dados.push(rowData);
        }
    });

    // Ordena os dados com base no número de atestados em ordem decrescente
    dados.sort((a, b) => b.atestados - a.atestados);

    // Cria um ranking baseado na posição de cada item nos dados
    const ranking = dados.map((item, index) => ({
        nome: item.QRA,
        atestados: item.atestados,
        ranking: index + 1
    }));

    return ranking;
}

consultarRankingAtestados()
    .then(ranking => {
        console.log(JSON.stringify(ranking, null, 2)); // Exibe o ranking no formato JSON
    })
    .catch(console.error);