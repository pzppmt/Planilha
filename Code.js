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
  const now = new Date();

  // --- Atualiza Inventário ---
  // Mapear tiers e categorias para achar índices na planilha
  const tiers = inventarioSheet.getRange(2, 1, inventarioSheet.getLastRow() - 1, 1).getValues().flat();
  const categorias = inventarioSheet.getRange(1, 2, 1, inventarioSheet.getLastColumn() - 1).getValues()[0];

  // Pega toda a matriz de quantidades do inventário
  const inventarioQuantidades = inventarioSheet.getRange(2, 2, tiers.length, categorias.length).getValues();

  // Para cada matéria prima usada:
  data.materiasPrimas.forEach(item => {
    // Ex: nome = "Barras 4.0" => separar em categoria e tier
    const partes = item.nome.split(' ');
    const categoria = partes[0];
    const tier = partes.slice(1).join(' ');

    const linha = tiers.indexOf(tier);
    const coluna = categorias.indexOf(categoria);

    if (linha >= 0 && coluna >= 0) {
      // Atualizar a quantidade subtraindo o que foi usado
      let atual = inventarioQuantidades[linha][coluna];
      inventarioQuantidades[linha][coluna] = Math.max(0, atual - item.quantidade);
    }
  });

  // Grava as quantidades atualizadas no Inventário
  inventarioSheet.getRange(2, 2, tiers.length, categorias.length).setValues(inventarioQuantidades);

  // --- Atualiza Artefatos ---
  if (artefatosSheet) {
    // Pega os dados atuais de artefatos
    const artefatosDados = artefatosSheet.getRange(2, 1, artefatosSheet.getLastRow() - 1, 2).getValues();

    // Para cada artefato usado, subtrai da planilha
    data.artefatos.forEach(item => {
      const idx = artefatosDados.findIndex(row => row[0] === item.nome);
      if (idx >= 0) {
        let atual = artefatosDados[idx][1];
        artefatosDados[idx][1] = Math.max(0, atual - item.quantidade);
      }
    });

    // Grava as quantidades atualizadas de volta
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

  return true;
}

