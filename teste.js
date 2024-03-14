const ExcelJS = require('exceljs');
const path = require('path');

const workbook = new ExcelJS.Workbook();
const filename = path.join(__dirname, 'bd.xlsx');

async function lerDadosExcel() {
    await workbook.xlsx.readFile(filename);
    const worksheet = workbook.getWorksheet('Planilha1');

    if (!worksheet) {
        throw new Error("Planilha 'dados' não encontrada no arquivo Excel.");
    }

    const dados = [];
    const headerRow = worksheet.getRow(1).values;

    worksheet.spliceRows(1, 1); // Remove a primeira linha (cabeçalho)

    worksheet.eachRow({ includeEmpty: false }, row => {
        const rowData = {};
        row.eachCell((cell, colNumber) => {
            rowData[headerRow[colNumber]] = cell.value;
        });
        dados.push(rowData);
    });

    return dados;
}


async function adicionarDadosExcel(novoDado) {
    await workbook.xlsx.readFile(filename);
    const worksheet = workbook.getWorksheet('Planilha1');

    const newRow = worksheet.addRow([novoDado.nome, novoDado.idade, novoDado.email]);

    // Copiar a formatação da primeira linha para a nova linha adicionada
    if (worksheet.rowCount > 1) {
        const firstRow = worksheet.getRow(2); // A segunda linha é a primeira linha de dados (após o cabeçalho)
        newRow.eachCell((cell, colNumber) => {
            cell.style = Object.assign({}, firstRow.getCell(colNumber).style);
        });
    }

    await workbook.xlsx.writeFile(filename);

    return {
        nome: newRow.getCell(1).value,
        idade: newRow.getCell(2).value,
        email: newRow.getCell(3).value
    };
}

// Exemplo de uso
async function main() {
    const dados = await lerDadosExcel();
    console.log('Dados atuais:', dados);

    const novoDado = {
        nome: 'Novo Nome',
        idade: 30,
        email: 'novo@email.com'
    };

    const novoRegistro = await adicionarDadosExcel(novoDado);
    console.log('Novo registro adicionado:', novoRegistro);
}

main().catch(console.error);
