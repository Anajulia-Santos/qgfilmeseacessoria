const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(bodyParser.urlencoded({ extended: true }));

app.post('/processar-formulario', (req, res) => {
  try {
    const { name, email, message } = req.body;
    const filePath = path.join(__dirname, 'dados.xlsx'); // Caminho para o arquivo Excel

    console.log('filePath:', filePath); // Verificação do caminho do arquivo

    let data = [];

    if (fs.existsSync(filePath)) {
      // Ler dados existentes
      const workbook = xlsx.readFile(filePath);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      data = xlsx.utils.sheet_to_json(sheet);
    }

    // Adicionar nova entrada
    data.push({ name, email, message });

    // Criar novo workbook e adicionar dados
    const newWorkbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.json_to_sheet(data);
    xlsx.utils.book_append_sheet(newWorkbook, worksheet, 'Sheet1');

    console.log('newWorkbook:', newWorkbook); // Verificação do objeto workbook

    // Escrever dados no arquivo usando write
    const wbout = xlsx.write(newWorkbook, { bookType: 'xlsx', type: 'buffer' });
    fs.writeFileSync(filePath, wbout);

    res.send('Formulário enviado com sucesso!');
  } catch (error) {
    console.error('Erro ao processar o formulário:', error);
    res.status(500).send('Ocorreu um erro ao processar o formulário.');
  }
});

app.listen(3000, () => console.log('Servidor iniciado na porta 3000'));
