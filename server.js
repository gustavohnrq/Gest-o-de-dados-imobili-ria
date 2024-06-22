const express = require('express');
const path = require('path');
const app = express();
const port = process.env.PORT || 3000;

// Middleware para servir arquivos estáticos
app.use(express.static(path.join(__dirname, 'public')));

// Rota para a página principal
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'MenuPrincipal.html'));
});

// Rotas para as outras páginas
app.get('/Form_Bairro', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'Form_Bairro.html'));
});

app.get('/Form_Caps', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'Form_Caps.html'));
});

app.get('/Form_Corretor', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'Form_Corretor.html'));
});

app.get('/Form_Estoque', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'Form_Estoque.html'));
});

app.get('/Form_Gerente', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'Form_Gerente.html'));
});

app.get('/Form_Saidas', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'Form_Saidas.html'));
});

app.get('/Form_Tipo', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'Form_Tipo.html'));
});

app.get('/Fomr_Vendas', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'Fomr_Vendas.html'));
});

// Inicia o servidor
app.listen(port, () => {
  console.log(`Servidor rodando em http://localhost:${port}`);
});
