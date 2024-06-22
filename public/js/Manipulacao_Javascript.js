
function showFormCaps() {
    const html = HtmlService.createHtmlOutputFromFile('Form_Caps')
        .setWidth(1280)
        .setHeight(720);
    SpreadsheetApp.getUi().showModalDialog(html, 'Registro de Captações');
}

function showExitForm() {
    const html = HtmlService.createHtmlOutputFromFile('Form_Saidas')
        .setWidth(1280)
        .setHeight(720);
    SpreadsheetApp.getUi().showModalDialog(html, 'Registro de Saídas');
}

function showSalesForm() {
    const html = HtmlService.createHtmlOutputFromFile('Form_Vendas')
        .setWidth(1280)
        .setHeight(720);
    SpreadsheetApp.getUi().showModalDialog(html, 'Registro de Vendas');
}

function showCorretorForm() {
    const html = HtmlService.createHtmlOutputFromFile('Form_Corretor')
        .setWidth(1280)
        .setHeight(720);
    const nextId = getNextCorretorId();
    const gerentes = getGerentes();
    const gerentesOptions = gerentes.map(g => `<option value="${g.id}">${g.nome}</option>`).join('');
    html.append(`<script>
        document.addEventListener('DOMContentLoaded', function() {
            document.getElementById('idCorretor').value = '${nextId}';
            const gerenteSelect = document.getElementById('idGerente');
            gerenteSelect.innerHTML = '<option value="">Selecionar Gerente</option>' + '${gerentesOptions}';
        });
    </script>`);
    SpreadsheetApp.getUi().showModalDialog(html, 'Cadastro de Corretor');
}

function showGerenteForm() {
    const html = HtmlService.createHtmlOutputFromFile('Form_Gerente')
        .setWidth(1280)
        .setHeight(720);
    const nextId = getNextGerenteId();
    html.append(`<script>
        document.addEventListener('DOMContentLoaded', function() {
            document.getElementById('idGerente').value = '${nextId}';
        });
    </script>`);
    SpreadsheetApp.getUi().showModalDialog(html, 'Cadastro de Gerente');
}

function showEstoqueForm() {
    const html = HtmlService.createHtmlOutputFromFile('Form_Estoque')
        .setWidth(1280)
        .setHeight(720);
    SpreadsheetApp.getUi().showModalDialog(html, 'Registrar Estoque');
}

function showTipoForm() {
    const html = HtmlService.createHtmlOutputFromFile('Form_Tipo')
        .setWidth(1280)
        .setHeight(720);
    SpreadsheetApp.getUi().showModalDialog(html, 'Cadastro de Tipo');
}

function showBairroForm() {
    const html = HtmlService.createHtmlOutputFromFile('Form_Bairro')
        .setWidth(1280)
        .setHeight(720);
    SpreadsheetApp.getUi().showModalDialog(html, 'Cadastro de Bairro');
}


function hideSheets(sheetNames) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheetNames.forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (sheet) {
            sheet.hideSheet();
        }
    });
}

function getCaptadores() {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheet = ss.getSheetByName('Dim_Corretor');
    const data = sheet.getRange('A2:C' + sheet.getLastRow()).getValues();
    return data.map(row => ({ IdCorretor: row[0], Nome: row[1], IdGerente: row[2] }));
}

function getManager(idCorretor) {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheet = ss.getSheetByName('Dim_Corretor');
    const data = sheet.getRange('A2:C' + sheet.getLastRow()).getValues();
    const manager = data.find(row => row[0] === idCorretor);
    return manager ? manager[2] : '';
}

function getGerentes() {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheet = ss.getSheetByName('Dim_Gerente');
    const data = sheet.getRange('A2:B' + sheet.getLastRow()).getValues();
    return data.map(row => ({ id: row[0], nome: row[1] }));
}

function getGerentesOptions() {
    const gerentes = getGerentes();
    return gerentes.map(g => `<option value="${g.id}">${g.nome}</option>`).join('');
}

function getOptions() {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheetTipos = ss.getSheetByName('Dim_Tipo');
    const sheetBairros = ss.getSheetByName('Dim_Bairro');

    const tiposData = sheetTipos.getRange('A2:B' + sheetTipos.getLastRow()).getValues();
    const bairrosData = sheetBairros.getRange('A2:B' + sheetBairros.getLastRow()).getValues();

    const tipos = tiposData.map(row => ({ id: row[0], nome: row[1] }));
    const bairros = bairrosData.map(row => ({ id: row[0], nome: row[1] }));

    return { tipos: tipos, bairros: bairros };
}// Ids Tipo e Bairro Form_Caps

function getBairros() {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheet = ss.getSheetByName('Dim_Bairro');
    const bairros = sheet.getRange('A2:B' + sheet.getLastRow()).getValues();
    return [...new Set(bairros.map(row => row[0]).filter(Boolean))];
}

function getDataForExit(codigo) {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheet = ss.getSheetByName('Fato_Estoque');
    const data = sheet.getRange('A2:F' + sheet.getLastRow()).getValues();
    const filteredData = data.filter(row => row[0].toString() === codigo.toString());

    if (filteredData.length > 0) {
        const sortedData = filteredData.sort((a, b) => new Date(b[5]) - new Date(a[5]));
        const latest = sortedData[0];
        return {
            captador1: latest[1],
            captador2: latest[2],
            captador3: latest[3],
            gerente: latest[4]
        };
    } else {
        return {}; // Retorna um objeto vazio se não encontrar dados
    }
}

function getNextCorretorId() {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheet = ss.getSheetByName('Dim_Corretor');
    const data = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat();
    const maxId = Math.max(...data.map(id => parseInt(id.replace('C61', ''))));
    const nextId = `C61${String(maxId + 1).padStart(3, '0')}`;
    return nextId;
}

function getNextGerenteId() {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheet = ss.getSheetByName('Dim_Gerente');
    const data = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat();
    const maxId = Math.max(...data.map(id => parseInt(id.replace('G61', ''))));
    const nextId = `G61${String(maxId + 1).padStart(3, '0')}`;
    return nextId;
}

function getNextTipoId() {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheet = ss.getSheetByName('Dim_Tipo');
    const data = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat();
    
    let maxId = 0;
    data.forEach(id => {
        const num = parseInt(id.replace('T', ''));
        if (!isNaN(num) && num > maxId) {
            maxId = num;
        }
    });

    const nextId = `T${String(maxId + 1).padStart(3, '0')}`;
    return nextId;
}

function getNextBairroId() {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheet = ss.getSheetByName('Dim_Bairro');
    const data = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat();
    
    let maxId = 0;
    data.forEach(id => {
        const num = parseInt(id.replace('B', ''));
        if (!isNaN(num) && num > maxId) {
            maxId = num;
        }
    });

    const nextId = `B${String(maxId + 1).padStart(3, '0')}`;
    return nextId;
}

function getCorretoresComGerentes() {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheetCorretores = ss.getSheetByName('Dim_Corretor');
    const sheetGerentes = ss.getSheetByName('Dim_Gerente');
    
    const corretores = sheetCorretores.getRange('A2:C' + sheetCorretores.getLastRow()).getValues();
    const gerentes = sheetGerentes.getRange('A2:C' + sheetGerentes.getLastRow()).getValues();
    const gerentesMap = new Map(gerentes.map(row => [row[0], row[1]]));
    
    return corretores.map(corretor => ({
        id: corretor[0],
        nome: corretor[1],
        idGerente: corretor[2],
        nomeGerente: gerentesMap.get(corretor[2])
    }));
}

function updateManagerSelect(corretorId, allData) {
    const selectedCorretor = document.getElementById(corretorId).value;
    const managerSelectId = corretorId.includes('vendedor') ?
                            corretorId.replace('vendedor', 'gerenteVenda') :
                            corretorId.replace('captador', 'gerenteCaptacao');
    const managerSelect = document.getElementById(managerSelectId);
    managerSelect.innerHTML = '<option value="">Selecionar Gerente</option>';

    const managerData = allData.find(item => item.id === selectedCorretor);
    if (managerData && managerData.nomeGerente) {
        const opt = document.createElement('option');
        opt.value = managerData.idGerente;
        opt.textContent = managerData.nomeGerente;
        managerSelect.appendChild(opt);
        managerSelect.value = managerData.idGerente;
    }
}

function populateDropdown(elementId, options) {
    const select = document.getElementById(elementId);
    select.innerHTML = '';
    options.forEach(option => {
        const opt = document.createElement('option');
        opt.value = option.id;
        opt.textContent = option.nome;
        select.appendChild(opt);
    });
}

function submitSalesData(data) {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheetVendas = ss.getSheetByName('Fato_Venda');
    const sheetImovel = ss.getSheetByName('Dim_Imovel');

    // Dados para Fato_Venda
    const rowDataVendas = [
        data.san,
        new Date(data.dataVenda),
        new Date(data.dataVenda),
        data.tempoVenda,
        data.bairro,
        data.quadra,
        data.tipo,
        data.focoPP === 'TRUE' || data.focoAC === 'TRUE', // Qualquer foco ativo define 'TRUE' aqui
        parseFloat(data.valorNegocio),
       
        parseFloat(data.valorComissao),
        parseFloat(data.porcentagemComissao),
        parseFloat(data.valorTotal61),
        data.participacao61,
        data.correcao61,
        data.correcaoVGV,
        data.nf61Imoveis,
        parseFloat(data.liquido61),
        data.correcaoVendedor1,
        data.v1,
        data.q1,
        data.vendedor1,
        data.imobiliaria,
        parseFloat(data.porcentagemVendedor1),
        parseFloat(data.valorVendedor1),
        data.gerenteVenda1,
        parseFloat(data.porcentagemGerenteVenda1),
        parseFloat(data.valorGerenteVenda1),
        data.correcaoVendedor2,
        data.v2,
        data.q2,
        data.vendedor2,
        parseFloat(data.porcentagemVendedor2),
        parseFloat(data.valorVendedor2),
        data.gerenteVenda2,
        parseFloat(data.porcentagemGerenteVenda2),
        parseFloat(data.valorGerenteVenda2),
        data.correcaoCap1,
        data.v3,
        data.q3,
        data.captador1,
        data.imobiliariaCaptador1,
        parseFloat(data.porcentagemCaptador1),
        parseFloat(data.valorCaptador1),
        data.gerenteCaptacao1,
        parseFloat(data.porcentagemGerenteCaptacao1),
        parseFloat(data.valorGerenteCaptacao1),
        data.correcaoCap2,
        data.v4,
        data.q4,
        data.captador2,
        parseFloat(data.porcentagemCaptador2),
        parseFloat(data.valorCaptador2),
        data.gerenteCaptacao2,
        parseFloat(data.porcentagemGerenteCaptacao2),
        parseFloat(data.valorGerenteCaptacao2),
        data.origemLeadVenda,
        data.tempoEmDias,
        new Date(data.entradaLead),
        data.idContrato
    ];

    sheetVendas.appendRow(rowDataVendas);

    // Dados para Dim_Imovel
    const rowDataImovel = [
        data.san,                     // Código
        data.tipo,                    // Tipo
        parseFloat(data.valorNegocio),// Valor
        data.bairro,                  // Bairro
        data.focoPP === 'TRUE',       // Foco PP
        data.focoAC === 'TRUE'        // Foco AC
    ];

    sheetImovel.appendRow(rowDataImovel);

    // Copiar fórmulas na planilha Fato_Venda, se aplicável
    const lastRow = sheetVendas.getLastRow() - 2;  // A última linha antes da inserção
    const range = sheetVendas.getRange(lastRow, 1, 1, sheetVendas.getLastColumn());
    const formulas = range.getFormulas()[0];

    if (formulas.filter(f => f !== '').length > 0) {
        const newFormulasRange = sheetVendas.getRange(lastRow + 1, 1, 1, formulas.length);
        newFormulasRange.setFormulas([formulas]);
    }

    return "Dados de venda registrados com sucesso e fórmulas copiadas.";
}

function submitData(data) {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheetCaptacao = ss.getSheetByName('Fato_Captacao');
    const sheetImovel = ss.getSheetByName('Dim_Imovel');

    sheetCaptacao.appendRow([
        data.codigo, data.captador1, data.captador2, data.captador3,
        data.gerente, data.dataEntrada
    ]);

    sheetImovel.appendRow([
        data.codigo, data.tipo, data.valor, data.bairro,
        data.focoPP ? 'TRUE' : 'FALSE', data.focoAC ? 'TRUE' : 'FALSE'
    ]);

    showSuccessMessage(data);
}

function submitExitData(data) {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheet = ss.getSheetByName('Fato_Saida');
    const dataDeSaida = data.dataSaida ? new Date(data.dataSaida) : new Date(); // Usa a data atual como fallback

    sheet.appendRow([
        data.codigo,
        data.captador1,
        data.captador2,
        data.captador3,
        data.gerente,
        data.motivo,
        dataDeSaida
    ]);

    showExitSuccessMessage(data);
}

function submitCorretorData(data) {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw'); // ID do seu Google Sheets
    const sheet = ss.getSheetByName('Dim_Corretor');
    sheet.appendRow([data.idCorretor, data.nomeCorretor, data.idGerente]);
}

function submitGerenteData(data) {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw'); // ID do seu Google Sheets
    const sheet = ss.getSheetByName('Dim_Gerente');
    sheet.appendRow([data.idGerente, data.nomeGerente]);
}

function processFile(data) {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw');
    const sheetCorretores = ss.getSheetByName('Dim_Corretor');
    const corretores = sheetCorretores.getRange('A2:C' + sheetCorretores.getLastRow()).getValues();

    const nomeParaId = {};
    const nomeParaGerente = {};

    corretores.forEach(corretor => {
        nomeParaId[corretor[1]] = corretor[0];
        nomeParaGerente[corretor[1]] = corretor[2];
    });

    const updatedRows = [];

    const headers = data[0];
    const rows = data.slice(1);

    rows.forEach(row => {
        const codigo = row[0];
        const captador1 = nomeParaId[row[1]] || row[1];
        const captador2 = nomeParaId[row[2]] || row[2];
        const captador3 = nomeParaId[row[3]] || row[3];
        const gerente = nomeParaGerente[row[1]] || '';
        const dataEstoque = row[5];
        updatedRows.push([codigo, captador1, captador2, captador3, gerente, dataEstoque]);
    });

    const sheetEstoque = ss.getSheetByName('Fato_Estoque');
    sheetEstoque.getRange(sheetEstoque.getLastRow() + 1, 1, updatedRows.length, updatedRows[0].length).setValues(updatedRows);

    return 'Arquivo processado e estoque atualizado com sucesso!';
}

function submitTipoData(data) {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw'); // ID do seu Google Sheets
    const sheet = ss.getSheetByName('Dim_Tipo');
    sheet.appendRow([data.idTipo, data.nomeTipo]);
}

function submitBairroData(data) {
    const ss = SpreadsheetApp.openById('1HQDdcbUMj276hnIbPs-WwdWHiUPzMhPRWt4HHRyYGnw'); // ID do seu Google Sheets
    const sheet = ss.getSheetByName('Dim_Bairro');
    sheet.appendRow([data.idBairro, data.nomeBairro]);
}

function showSuccessMessage(data) {
    const message = 'Sucesso! A captação foi registrada com sucesso:\n' +
                    'Código: ' + data.codigo + '\n' +
                    'Tipo: ' + data.tipo + '\n' +
                    'Valor: ' + data.valor + '\n' +
                    'Bairro: ' + data.bairro + '\n' +
                    'Data de Entrada: ' + data.dataEntrada;
    SpreadsheetApp.getUi().alert(message);
}

function showExitSuccessMessage(data) {
    const date = new Date(data.dataSaida);
    const formattedDate = date.isValid() ? date.toLocaleDateString() : 'Data inválida';

    const message = 'Sucesso! A saída foi registrada com sucesso:\n' +
                    'Código: ' + data.codigo + '\n' +
                    'Motivo: ' + data.motivo + '\n' +
                    'Data de Saída: ' + formattedDate;
    SpreadsheetApp.getUi().alert(message);
}

Date.prototype.isValid = function () {
    return this.getTime() === this.getTime();   // NaN não é igual a NaN, isso verifica se a data é NaN
};


