const express = require('express');
const path = require('path');
const fs = require('fs');
const { google } = require('googleapis');
const app = express();
const port = process.env.PORT || 3000;

// Middleware para servir arquivos estáticos
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// Carregar credenciais do Google API
const credentials = JSON.parse(fs.readFileSync('credentials.json'));
const { client_secret, client_id, redirect_uris } = credentials.installed;
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

// Função para autenticar o cliente OAuth
function authenticate() {
    return new Promise((resolve, reject) => {
        fs.readFile('token.json', (err, token) => {
            if (err) return getAccessToken(oAuth2Client, resolve, reject);
            oAuth2Client.setCredentials(JSON.parse(token));
            resolve(oAuth2Client);
        });
    });
}

function getAccessToken(oAuth2Client, resolve, reject) {
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    console.log('Authorize this app by visiting this url:', authUrl);
    const rl = require('readline').createInterface({
        input: process.stdin,
        output: process.stdout,
    });
    rl.question('Enter the code from that page here: ', (code) => {
        rl.close();
        oAuth2Client.getToken(code, (err, token) => {
            if (err) return reject(err);
            oAuth2Client.setCredentials(token);
            fs.writeFile('token.json', JSON.stringify(token), (err) => {
                if (err) return reject(err);
                resolve(oAuth2Client);
            });
        });
    });
}

// Configure sua API de Rotas para Funções
app.post('/getCaptadores', async (req, res) => {
    try {
        const data = await getCaptadores();
        res.json(data);
    } catch (error) {
        res.status(500).send(error.message);
    }
});

// Implementação de Funções
async function getCaptadores() {
    const auth = await authenticate();
    const sheets = google.sheets({ version: 'v4', auth });
    const response = await sheets.spreadsheets.values.get({
        spreadsheetId: '1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw',
        range: 'Dim_Corretor!A2:C'
    });
    const rows = response.data.values;
    if (rows.length) {
        return rows.map(row => ({
            IdCorretor: row[0],
            Nome: row[1],
            IdGerente: row[2]
        }));
    } else {
        return [];
    }
}

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
