const express = require('express');
const path = require('path');
const fs = require('fs');
const { google } = require('googleapis');
const app = express();
const port = process.env.PORT || 3000;

// Middleware para servir arquivos estáticos
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// Carregar credenciais do Google API das variáveis de ambiente
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REDIRECT_URI = process.env.REDIRECT_URI;

if (!CLIENT_ID || !CLIENT_SECRET || !REDIRECT_URI) {
    console.error('Erro: CLIENT_ID, CLIENT_SECRET, e REDIRECT_URI precisam estar definidos.');
    process.exit(1);
}

const oAuth2Client = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI);

// Função para autenticar o cliente OAuth
async function authenticate() {
    return new Promise((resolve, reject) => {
        fs.readFile('token.json', (err, token) => {
            if (err) {
                return getAccessToken(oAuth2Client, resolve, reject);
            }
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
}

// Rota para receber o código de autorização
app.get('/oauth2callback', (req, res) => {
    const code = req.query.code;
    if (!code) {
        return res.status(400).send('Código de autorização ausente');
    }
    oAuth2Client.getToken(code, (err, token) => {
        if (err) {
            return res.status(400).send('Erro ao obter token de acesso');
        }
        oAuth2Client.setCredentials(token);
        fs.writeFile('token.json', JSON.stringify(token), (err) => {
            if (err) {
                return res.status(500).send('Erro ao salvar token');
            }
            res.send('Autorização bem-sucedida. Você pode fechar esta janela.');
        });
    });
});

// Funções de API
app.post('/getCaptadores', async (req, res) => {
    try {
        const auth = await authenticate();
        const sheets = google.sheets({ version: 'v4', auth });
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: '1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw',
            range: 'Dim_Corretor!A2:C'
        });
        const rows = response.data.values;
        const captadores = rows.map(row => ({
            IdCorretor: row[0],
            Nome: row[1],
            IdGerente: row[2]
        }));
        res.json(captadores);
    } catch (error) {
        res.status(500).send(error.message);
    }
});

app.post('/getBairros', async (req, res) => {
    try {
        const auth = await authenticate();
        const sheets = google.sheets({ version: 'v4', auth });
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: '1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw',
            range: 'Dim_Bairro!A2:B'
        });
        const rows = response.data.values;
        const bairros = rows.map(row => ({
            id: row[0],
            nome: row[1]
        }));
        res.json(bairros);
    } catch (error) {
        res.status(500).send(error.message);
    }
});

app.post('/getOptions', async (req, res) => {
    try {
        const auth = await authenticate();
        const sheets = google.sheets({ version: 'v4', auth });
        const [tiposResponse, bairrosResponse] = await Promise.all([
            sheets.spreadsheets.values.get({
                spreadsheetId: '1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw',
                range: 'Dim_Tipo!A2:B'
            }),
            sheets.spreadsheets.values.get({
                spreadsheetId: '1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw',
                range: 'Dim_Bairro!A2:B'
            })
        ]);
        const tipos = tiposResponse.data.values.map(row => ({
            id: row[0],
            nome: row[1]
        }));
        const bairros = bairrosResponse.data.values.map(row => ({
            id: row[0],
            nome: row[1]
        }));
        res.json({ tipos, bairros });
    } catch (error) {
        res.status(500).send(error.message);
    }
});

app.post('/getManager', async (req, res) => {
    const { idCorretor } = req.query;
    try {
        const auth = await authenticate();
        const sheets = google.sheets({ version: 'v4', auth });
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: '1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw',
            range: 'Dim_Corretor!A2:C'
        });
        const rows = response.data.values;
        const manager = rows.find(row => row[0] === idCorretor);
        res.json(manager ? manager[2] : '');
    } catch (error) {
        res.status(500).send(error.message);
    }
});

app.post('/submitData', async (req, res) => {
    const data = req.body;
    try {
        const auth = await authenticate();
        const sheets = google.sheets({ version: 'v4', auth });
        await sheets.spreadsheets.values.append({
            spreadsheetId: '1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw',
            range: 'Fato_Captacao!A2',
            valueInputOption: 'RAW',
            resource: {
                values: [[
                    data.codigo, data.captador1, data.captador2, data.captador3,
                    data.gerente, data.dataEntrada, data.tipo, data.valor,
                    data.bairro, data.focoPP, data.focoAC
                ]]
            }
        });
        res.json({ status: 'success' });
    } catch (error) {
        res.status(500).send(error.message);
    }
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

// Rota para a página principal
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'MenuPrincipal.html'));
});

// Inicia o servidor
app.listen(port, () => {
    console.log(`Servidor rodando em http://localhost:${port}`);
});
