function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Gestão de Dados')
        .addItem('Registrar Captação', 'showForm')
        .addItem('Registrar Saída', 'showExitForm')
        .addItem('Registrar Venda', 'showSalesForm')
        .addToUi();
}

function showForm() {
    const html = HtmlService.createHtmlOutputFromFile('Formulario')
        .setWidth(500)
        .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Registro de Captações');
}

function showExitForm() {
    const html = HtmlService.createHtmlOutputFromFile('FormularioSaida')
        .setWidth(500)
        .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Registro de Saídas');
}

function showSalesForm() {
    const html = HtmlService.createHtmlOutputFromFile('FormularioVendas')
        .setWidth(500)
        .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Registro de Vendas');
}

function getCaptadores() {
    const ss = SpreadsheetApp.openById('1aTA1niEFGq6SMUSVUGVuo76sJOwh_-uIv9ARxmjkS0I');
    const sheet = ss.getSheetByName('Dim_Corretor');
    const data = sheet.getRange('A2:C' + sheet.getLastRow()).getValues();
    return data.map(row => ({ IdCorretor: row[0], Nome: row[1], IdGerente: row[2] }));
}

function getManager(idCorretor) {
    const ss = SpreadsheetApp.openById('1aTA1niEFGq6SMUSVUGVuo76sJOwh_-uIv9ARxmjkS0I');
    const sheet = ss.getSheetByName('Dim_Corretor');
    const data = sheet.getRange('A2:C' + sheet.getLastRow()).getValues();
    const manager = data.find(row => row[0] === idCorretor);
    return manager ? manager[2] : '';
}

function getBairros() {
    const ss = SpreadsheetApp.openById('1aTA1niEFGq6SMUSVUGVuo76sJOwh_-uIv9ARxmjkS0I');
    const sheet = ss.getSheetByName('Dim_Imovel');
    const bairros = sheet.getRange('D2:D' + sheet.getLastRow()).getValues();
    return [...new Set(bairros.map(row => row[0]).filter(Boolean))];
}

function getCorretoresComGerentes() {
    const ss = SpreadsheetApp.openById('1aTA1niEFGq6SMUSVUGVuo76sJOwh_-uIv9ARxmjkS0I');
    const sheetCorretores = ss.getSheetByName('Dim_Corretor');
    const sheetGerentes = ss.getSheetByName('Dim_Gerente');
    
    const corretores = sheetCorretores.getRange('A2:B' + sheetCorretores.getLastRow()).getValues();
    const gerentes = sheetGerentes.getRange('A2:C' + sheetGerentes.getLastRow()).getValues();
    const gerentesMap = new Map(gerentes.map(row => [row[0], row[1]]));
    
    return corretores.map(corretor => ({
        id: corretor[0],
        nome: corretor[1],
        idGerente: corretor[2],
        nomeGerente: gerentesMap.get(corretor[2])
    }));
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

function updateManagerSelect(elementId, options) {
    const select = document.getElementById(elementId);
    const selectedValue = select.value;
    const managerSelectId = elementId.replace('vendedor', 'gerenteVenda').replace('captador', 'gerenteCaptacao');
    const managerSelect = document.getElementById(managerSelectId);
    managerSelect.innerHTML = ''; // Limpar opções existentes
    const managerInfo = options.find(option => option.id === selectedValue);
    if (managerInfo) {
        const opt = document.createElement('option');
        opt.value = managerInfo.idGerente;
        opt.textContent = managerInfo.nomeGerente;
        managerSelect.appendChild(opt);
    }
}

function submitSalesData(data) {
    const ss = SpreadsheetApp.openById('1aTA1niEFGq6SMUSVUGVuo76sJOwh_-uIv9ARxmjkS0I');
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
