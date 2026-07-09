/**
 * CHECKLIST FROTA — Integração com Google Sheets
 * ------------------------------------------------
 * Este script recebe os envios do Checklist Web (checklist_frota.html)
 * e grava cada um como uma nova linha na planilha "Checklis Frota - Seven",
 * separando os registros em duas abas:
 *   - "Checklis Diário"
 *   - "Checklis Semanal"
 *
 * Cada item do checklist vira uma coluna própria, preenchida com:
 *   "OK"                        -> quando o técnico não marcou o item
 *   "Divergência" ou
 *   "Divergência: <observação>" -> quando o técnico marcou o item como problema
 *
 * COMO INSTALAR:
 * 1. Abra a planilha "Checklis Frota - Seven" no Google Sheets.
 * 2. Menu Extensões > Apps Script.
 * 3. Apague o conteúdo padrão e cole todo este arquivo.
 * 4. Na barra de funções (topo do editor), selecione "configurarPlanilha"
 *    e clique em Executar uma vez — isso já cria as duas abas com os
 *    cabeçalhos corretos, mesmo antes do primeiro envio. Autorize o
 *    acesso quando solicitado.
 * 5. Clique em "Implantar" > "Nova implantação".
 * 6. Em "Tipo", escolha "App da Web".
 * 7. Em "Executar como", escolha "Eu" (sua conta).
 * 8. Em "Quem pode acessar", escolha "Qualquer pessoa".
 * 9. Clique em Implantar e autorize o acesso quando solicitado.
 * 10. Copie a "URL do app da Web" gerada.
 * 11. No Checklist Web, clique no ícone de engrenagem > aba "Integração"
 *     e cole essa URL em "URL do Google Apps Script (Web App)". Salve.
 *
 * Pronto — cada envio do checklist grava uma linha automaticamente na
 * aba correta ("Checklis Diário" ou "Checklis Semanal").
 *
 * OBS: se depois de editar o código você implantar novamente, use
 * "Gerenciar implantações" > editar (ícone de lápis) > "Nova versão"
 * para manter a mesma URL ativa.
 *
 * IMPORTANTE: a lista de itens abaixo (ITEMS) precisa ficar idêntica
 * (mesmo texto, mesma ordem) à constante CHECKLIST definida dentro do
 * checklist_frota.html. Se algum item for adicionado/removido/renomeado
 * no site, replique a mudança aqui também.
 */

const ITEMS = {
  'Diário': [
    ['Painel & Fluidos', 'Nível Combustível'],
    ['Painel & Fluidos', 'Nível água Arrefecimento'],
    ['Painel & Fluidos', 'Indicadores Painel'],
    ['Iluminação', 'Luz Freio Esq.'],
    ['Iluminação', 'Luz Freio Dir.'],
    ['Iluminação', 'Pisca Esq.'],
    ['Iluminação', 'Pisca Dir.'],
    ['Iluminação', 'Farol Esq.'],
    ['Iluminação', 'Farol Dir.'],
    ['Iluminação', 'Lanterna Dir.'],
    ['Iluminação', 'Lanterna Esq.'],
    ['Iluminação', 'Luz Placa'],
    ['Conforto & Acessórios', 'Ar. Condicionado'],
    ['Conforto & Acessórios', 'Limpeza Interna'],
    ['Conforto & Acessórios', 'Carregador Celular'],
    ['Conforto & Acessórios', 'Suporte p/ Celular'],
    ['Conforto & Acessórios', 'Cabo p/ Celular']
  ],
  'Semanal': [
    ['Iluminação', 'Luz Freio Esq.'],
    ['Iluminação', 'Luz Freio Dir.'],
    ['Iluminação', 'Pisca Esq.'],
    ['Iluminação', 'Pisca Dir.'],
    ['Iluminação', 'Farol Esq.'],
    ['Iluminação', 'Farol Dir.'],
    ['Iluminação', 'Lanterna Dir.'],
    ['Iluminação', 'Lanterna Esq.'],
    ['Iluminação', 'Luz Placa'],
    ['Iluminação', 'Pisca Alerta'],
    ['Mecânica & Fluidos', 'Nível Óleo Motor'],
    ['Mecânica & Fluidos', 'Nível Óleo Hidráulico'],
    ['Mecânica & Fluidos', 'Limp. Para brisa'],
    ['Mecânica & Fluidos', 'Para brisa Dianteiro'],
    ['Mecânica & Fluidos', 'Nível Água Para brisa'],
    ['Mecânica & Fluidos', 'Nível água Arrefecimento'],
    ['Mecânica & Fluidos', 'Freio de mão'],
    ['Mecânica & Fluidos', 'Pneus D/T'],
    ['Visão & Estrutura', 'Retrovisor Interno'],
    ['Visão & Estrutura', 'Retrovisor Dir.'],
    ['Visão & Estrutura', 'Retrovisor Esq.'],
    ['Visão & Estrutura', 'Vidros Laterais'],
    ['Visão & Estrutura', 'Vidros Elétricos'],
    ['Visão & Estrutura', 'Quebra Sol'],
    ['Visão & Estrutura', 'Sensor de ré'],
    ['Visão & Estrutura', 'Antena'],
    ['Segurança & Documentação', 'Triangulo'],
    ['Segurança & Documentação', 'Estepe'],
    ['Segurança & Documentação', 'Chave de Roda'],
    ['Segurança & Documentação', 'Macaco'],
    ['Segurança & Documentação', 'Documento Veicular'],
    ['Segurança & Documentação', 'Cintos de Segurança'],
    ['Segurança & Documentação', 'Chave Ignição'],
    ['Segurança & Documentação', 'Travamento das portas'],
    ['Segurança & Documentação', 'Indicadores Painel'],
    ['Conforto & Limpeza', 'Buzina'],
    ['Conforto & Limpeza', 'Ar. Condicionado'],
    ['Conforto & Limpeza', 'Rádio'],
    ['Conforto & Limpeza', 'Estofamento Bancos'],
    ['Conforto & Limpeza', 'Limpeza Externa'],
    ['Conforto & Limpeza', 'Limpeza Interna'],
    ['Conforto & Limpeza', 'Carregador Celular'],
    ['Conforto & Limpeza', 'Cabo Celular'],
    ['Conforto & Limpeza', 'Suporte p/ Celular'],
    ['Conforto & Limpeza', 'Tapetes']
  ]
};

const SHEET_NAMES = {
  'Diário': 'Checklis Diário',
  'Semanal': 'Checklis Semanal'
};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const tipo = (data.tipo === 'Semanal') ? 'Semanal' : 'Diário';
    const sheet = getOrCreateSheet_(tipo);
    const itemsDef = ITEMS[tipo];
    const itensRecebidos = Array.isArray(data.itens) ? data.itens : [];

    // Monta uma coluna por item, na mesma ordem definida em ITEMS.
    // Item não encontrado no payload (ou sem checkbox marcado) = "OK".
    const itemCols = itemsDef.map(function (def) {
      const grupo = def[0], label = def[1];
      const found = itensRecebidos.filter(function (it) {
        return it.grupo === grupo && it.label === label;
      })[0];
      if (!found || found.status !== 'Divergência') return 'OK';
      return found.obs ? ('Divergência: ' + found.obs) : 'Divergência';
    });

    const totalDivergencias = itemCols.filter(function (c) { return c !== 'OK'; }).length;

    const row = [
      new Date(),                 // Timestamp do envio (servidor)
      data.data || '',            // Data informada no checklist
      data.placa || '',
      data.veiculo || '',
      data.km || '',
      data.responsavel || '',
      data.revisor || ''
    ].concat(itemCols).concat([
      itemsDef.length,
      totalDivergencias,
      data.observacoes || ''
    ]);

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', aba: SHEET_NAMES[tipo] }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput('Checklist Frota — Web App ativo.');
}

/**
 * Cria (se necessário) as duas abas com os cabeçalhos corretos.
 * Pode ser executada manualmente pelo editor de Apps Script antes do
 * primeiro envio, ou roda automaticamente no primeiro POST de cada tipo.
 */
function configurarPlanilha() {
  getOrCreateSheet_('Diário');
  getOrCreateSheet_('Semanal');
}

function getOrCreateSheet_(tipo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = SHEET_NAMES[tipo];
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  if (sheet.getLastRow() === 0) {
    const header = ['Timestamp (envio)', 'Data', 'Placa', 'Veículo', 'KM', 'Responsável', 'Revisor']
      .concat(ITEMS[tipo].map(function (def) { return def[0] + ' — ' + def[1]; }))
      .concat(['Total de Itens', 'Total Divergências', 'Observações Gerais']);
    sheet.appendRow(header);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);
  }
  return sheet;
}
