function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Produção')
    .addItem('Iniciar Transporte', 'showDialog')
    .addToUi();
}

function showDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(400)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Iniciar Transporte');
}

// Função que retorna os dados para o diálogo
function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Aba "Main" para produtos
  const demandaSheet = ss.getSheetByName('Main');
  if (!demandaSheet) throw new Error('Aba "Main" não encontrada.');

  const produtos = demandaSheet.getRange('A2:A').getValues().flat().filter(String);

  // Aba "Inventário" para matérias-primas
  const inventarioSheet = ss.getSheetByName('Inventário');
  if (!inventarioSheet) throw new Error('Aba "Inventário" não encontrada.');

  const tiers = inventarioSheet.getRange(2, 1, inventarioSheet.getLastRow() - 1, 1).getValues().flat();
  const categorias = inventarioSheet.getRange(1, 2, 1, inventarioSheet.getLastColumn() - 1).getValues()[0];
  const inventarioData = inventarioSheet.getRange(2, 2, tiers.length, categorias.length).getValues();

  const inventario = [];
  for(let i = 0; i < tiers.length; i++) {
    for(let j = 0; j < categorias.length; j++) {
      const nome = categorias[j] + " " + tiers[i];
      const qtd = inventarioData[i][j] || 0;
      inventario.push({ nome, quantidade: qtd });
    }
  }

  // Aba "Artefatos" para matérias-primas secundárias
  const artefatosSheet = ss.getSheetByName('Artefatos');
  let artefatos = [];
  if (artefatosSheet) {
    const lastRow = artefatosSheet.getLastRow();
    if (lastRow > 1) {
      artefatos = artefatosSheet.getRange(2, 1, lastRow - 1, 2).getValues()
        .filter(row => row[0] && row[1]);
    }
  }

  return { produtos, inventario, artefatos };
}

function processTransport(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logsSheet = ss.getSheetByName('Logs');
  const inventarioSheet = ss.getSheetByName('Inventário');
  const artefatosSheet = ss.getSheetByName('Artefatos');
  const mainSheet = ss.getSheetByName('Main');
  const now = new Date();

  // --- Atualiza Inventário ---
  const tiers = inventarioSheet.getRange(2, 1, inventarioSheet.getLastRow() - 1, 1).getValues().flat();
  const categorias = inventarioSheet.getRange(1, 2, 1, inventarioSheet.getLastColumn() - 1).getValues()[0];
  const inventarioQuantidades = inventarioSheet.getRange(2, 2, tiers.length, categorias.length).getValues();

  data.materiasPrimas.forEach(item => {
    const partes = item.nome.split(' ');
    const categoria = partes[0];
    const tier = partes.slice(1).join(' ');

    const linha = tiers.indexOf(tier);
    const coluna = categorias.indexOf(categoria);

    if (linha >= 0 && coluna >= 0) {
      let atual = inventarioQuantidades[linha][coluna];
      inventarioQuantidades[linha][coluna] = Math.max(0, atual - item.quantidade);
    }
  });

  inventarioSheet.getRange(2, 2, tiers.length, categorias.length).setValues(inventarioQuantidades);

  // --- Atualiza Artefatos ---
  if (artefatosSheet) {
    const artefatosDados = artefatosSheet.getRange(2, 1, artefatosSheet.getLastRow() - 1, 2).getValues();

    data.artefatos.forEach(item => {
      const idx = artefatosDados.findIndex(row => row[0] === item.nome);
      if (idx >= 0) {
        let atual = artefatosDados[idx][1];
        artefatosDados[idx][1] = Math.max(0, atual - item.quantidade);
      }
    });

    artefatosSheet.getRange(2, 2, artefatosDados.length, 1).setValues(artefatosDados.map(row => [row[1]]));
  }

  // --- Log ---
  logsSheet.appendRow([
    now,
    data.produto,
    JSON.stringify(data.materiasPrimas),
    JSON.stringify(data.artefatos),
    'Transporte iniciado e estoque atualizado'
  ]);

  // --- Marca Transporte como TRUE na aba Main ---
  if (mainSheet) {
    const produtos = mainSheet.getRange('A2:A' + mainSheet.getLastRow()).getValues().flat();
    const index = produtos.findIndex(p => typeof p === 'string' && p.trim().toLowerCase() === data.produto.trim().toLowerCase());

    if (index !== -1) {
      const linha = index + 2; // compensar cabeçalho
      mainSheet.getRange(linha, 3).setValue(true); // Coluna C = Transporte
    }
  }

  return true;
}
function onEdit(e) {
  const sheet = e.range.getSheet();
  const col = e.range.getColumn();
  const row = e.range.getRow();
  const newValue = e.value;
  const oldValue = e.oldValue;

  if (sheet.getName() !== 'Main') return;

  const ui = SpreadsheetApp.getUi();

  // COLUNA C (Transporte): Cancelar transporte
    // COLUNA C (Transporte): Cancelar transporte
  if (col === 3) {
    const isUnchecked = newValue === null || newValue === false || newValue === 'FALSE';
    const wasChecked = oldValue === true || oldValue === 'TRUE';

    if (wasChecked && isUnchecked) {
      const response = ui.alert('Cancelar esta produção?', ui.ButtonSet.YES_NO);
      if (response === ui.Button.NO) {
        sheet.getRange(row, col).setValue(true);
      }
    }
  }


  // COLUNA D (Produção): Ignorar
  if (col === 4) return;

  // COLUNA E (Entregue): Confirmar entrega
  if (col === 5 && newValue === 'TRUE') {
    const response = ui.alert('Entregar produto agora?', ui.ButtonSet.YES_NO);
    if (response === ui.Button.NO) {
      sheet.getRange(row, col).setValue(false);
      return;
    }

    // Confirmou a entrega — limpar linha
    const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const produto = rowValues[0]; // Coluna A
    const now = new Date();
    const logsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');

    // Logar ação
    logsSheet.appendRow([
      now,
      produto,
      '',
      '',
      'Produto entregue e linha limpa'
    ]);

    // Limpar a linha inteira
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).clearContent();
  }
}

