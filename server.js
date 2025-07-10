const express = require("express");
const bodyParser = require("body-parser");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const cors = require("cors");

const app = express();
const PORT = 3000;

app.use(cors());
app.use(bodyParser.json());

const registros = [];

app.post("/api/salvar", (req, res) => {
  const dados = req.body;
  registros.push(dados);
  res.json({ mensagem: "Registro recebido no servidor." });
});

app.post("/api/finalizar", (req, res) => {
  if (registros.length === 0) {
    return res.status(400).json({ erro: "Nenhum registro para salvar." });
  }

  const [dia, mes, ano] = registros[0].Data.split('/');
  const turno = registros[0].Turno || 'X';
  const nomeArquivo = `registro_turno_${turno}_${dia}-${mes}-${ano}.xlsx`;
  const pasta = path.join(__dirname, "arquivos");

  if (!fs.existsSync(pasta)) {
    fs.mkdirSync(pasta);
  }

  const caminho = path.join(pasta, nomeArquivo);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(registros);
  XLSX.utils.book_append_sheet(wb, ws, "Registros");
  XLSX.writeFile(wb, caminho);

  registros.length = 0;
  res.json({ mensagem: `Planilha "${nomeArquivo}" salva com sucesso no servidor.` });
});

app.listen(PORT, () => {
  console.log(`Servidor rodando em http://localhost:${PORT}`);
});
