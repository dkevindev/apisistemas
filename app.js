const express = require('express');
const app = express();
const ExcelJS = require('exceljs');
const path = require('path');
const cors = require('cors');

app.use(cors());



async function lerDadosExcel() {

  const workbook = new ExcelJS.Workbook();
  const filename = path.join(__dirname, 'planodechamada.xlsx');
  await workbook.xlsx.readFile(filename);
  const worksheet = workbook.getWorksheet('plano');

  if (!worksheet) {
    throw new Error("Planilha 'dados' não encontrada no arquivo Excel.");
  }

  const dados = [];
  const headerRow = worksheet.getRow(7).values;

  worksheet.spliceRows(7, 1); // Remove a sétima linha (cabeçalho)

  worksheet.eachRow({ includeEmpty: false, firstRow: 8 }, row => {
    const rowData = {};
    row.eachCell((cell, colNumber) => {
      rowData[headerRow[colNumber]] = cell.value;
    });
    dados.push(rowData);
  });

  return dados;
}

async function lerDadosExcelAtestados() {
  const workbook = new ExcelJS.Workbook();
  const filename = path.join(__dirname, 'atestados.xlsx');
  await workbook.xlsx.readFile(filename);
  const worksheet = workbook.getWorksheet('atestados');

  if (!worksheet) {
    throw new Error("Planilha 'dados' não encontrada no arquivo Excel.");
  }

  const dados = [];
  const headerRow = worksheet.getRow(1).values;

  worksheet.spliceRows(7, 1); // Remove a sétima linha (cabeçalho)

  worksheet.eachRow({ includeEmpty: false, firstRow: 2 }, row => {
    const rowData = {};
    row.eachCell((cell, colNumber) => {
      rowData[headerRow[colNumber]] = cell.value;
    });
    dados.push(rowData);
  });

  return dados;
}




app.get('/api/dados', async (req, res) => {
  try {
    const dados = await lerDadosExcel();
    res.json(dados);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});


app.get('/api/tabela', async (req, res) => {
  try {
    const dados = await lerDadosExcelAtestados();
    res.json(dados);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});

