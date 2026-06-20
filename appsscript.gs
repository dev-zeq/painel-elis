// 📊 GOOGLE SHEETS — AGENDAMENTO ELIS MIRANDA
// Versão corrigida: suporta JSON, form-encoded (no-cors) e todas as actions do painel

// ═══════════════════════════════════════════════════════════════════
// ⚙️ CONFIGURAÇÕES
// ═══════════════════════════════════════════════════════════════════

const CONFIG = {
  sheetsNames: {
    agendamentos: 'Agendamentos',
    dashboard:    'Dashboard',
    config:       'Configuração'
  },
  colors: {
    headerBg:    '#E91E63',
    headerText:  '#FFFFFF',
    dashboardBg: '#FFF5F8',
    goldenAccent:'#D4AF37'
  },
  headers: ['ID', 'Data', 'Hora', 'Nome Cliente', 'Telefone', 'Email', 'Serviço', 'Status', 'Anotações', 'Data Agendamento']
};

// ═══════════════════════════════════════════════════════════════════
// 🌐 doGet — Painel lê os agendamentos
// ═══════════════════════════════════════════════════════════════════

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'list';

  if (action === 'list')    return jsonResponse({ agendamentos: getAgendamentos() });
  if (action === 'blocked') return jsonResponse({ dates: getBlockedDates() });
  if (action === 'horarios') {
    const date = e.parameter && e.parameter.date;
    if (!date) return jsonResponse({ times: [] });
    return jsonResponse({ times: getHorariosOcupados(date) });
  }

  return jsonResponse({ success: false, error: 'Ação desconhecida: ' + action });
}

// ═══════════════════════════════════════════════════════════════════
// 🌐 doPost — Recebe do formulário (form-encoded / no-cors) e do painel (URLSearchParams)
// ═══════════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    let dados = parseRequest(e);
    const action = dados.action || 'add';

    if (action === 'add')         return addAgendamento(dados);
    if (action === 'updateStatus') return atualizarStatus(dados);
    if (action === 'blockDate')   return blockDate(dados);
    if (action === 'unblockDate') return unblockDate(dados);

    return jsonResponse({ success: false, error: 'Ação desconhecida: ' + action });

  } catch (error) {
    Logger.log('❌ Erro doPost: ' + error);
    return jsonResponse({ success: false, error: error.toString() });
  }
}

// ─── Tenta JSON primeiro, cai em e.parameter se for form-encoded ────────────
function parseRequest(e) {
  // Tentativa 1: JSON (n8n envia JSON)
  if (e.postData && e.postData.contents) {
    try {
      const parsed = JSON.parse(e.postData.contents);
      if (parsed && typeof parsed === 'object') return parsed;
    } catch (_) {}
  }

  // Tentativa 2: form-encoded (formulário com no-cors / URLSearchParams)
  if (e.parameter) {
    return e.parameter;
  }

  throw new Error('Não foi possível ler os dados da requisição');
}

// ═══════════════════════════════════════════════════════════════════
// ➕ Adicionar novo agendamento
// ═══════════════════════════════════════════════════════════════════

function addAgendamento(dados) {
  if (!dados.name || !dados.phone || !dados.date || !dados.time) {
    return jsonResponse({ success: false, error: 'Campos obrigatórios: name, phone, date, time' });
  }

  const aba = getAbaAgendamentos();
  const novoId = gerarId();

  const novaLinha = [
    novoId,                                      // ID
    dados.date,                                  // Data
    dados.time,                                  // Hora
    dados.name,                                  // Nome Cliente
    dados.phone,                                 // Telefone
    dados.email   || '',                         // Email
    dados.servico || 'Estética',                 // Serviço
    dados.status  || 'Pendente',                 // Status
    dados.notes   || dados.anotacoes || '',       // Anotações
    new Date().toLocaleDateString('pt-BR')        // Data do agendamento
  ];

  aba.appendRow(novaLinha);

  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha % 2 === 0) {
    aba.getRange(ultimaLinha, 1, 1, CONFIG.headers.length).setBackground('#F9F9F9');
  }

  Logger.log('✅ Agendamento: ' + dados.name + ' — ' + dados.date + ' ' + dados.time);

  try { enviarConfirmacaoEmail(dados); } catch (_) {}

  return jsonResponse({ success: true, id: novoId, message: 'Agendamento salvo!' });
}

// ═══════════════════════════════════════════════════════════════════
// ✏️ Atualizar status de um agendamento
// ═══════════════════════════════════════════════════════════════════

function atualizarStatus(dados) {
  if (!dados.id || !dados.status) {
    return jsonResponse({ success: false, error: 'id e status são obrigatórios' });
  }

  const aba  = getAbaAgendamentos();
  const rows = aba.getDataRange().getValues();

  // Coluna A = ID (índice 0), Coluna H = Status (índice 7)
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(dados.id)) {
      aba.getRange(i + 1, 8).setValue(dados.status); // coluna 8 = Status
      Logger.log('✅ Status atualizado: ID ' + dados.id + ' → ' + dados.status);
      return jsonResponse({ success: true });
    }
  }

  return jsonResponse({ success: false, error: 'ID não encontrado: ' + dados.id });
}

// ═══════════════════════════════════════════════════════════════════
// 📋 Ler todos os agendamentos (para o painel)
// ═══════════════════════════════════════════════════════════════════

function getAgendamentos() {
  const aba  = getAbaAgendamentos();
  const rows = aba.getDataRange().getValues();
  const lista = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] && !r[1]) continue; // linha vazia

    lista.push({
      id:      r[0] || (i + 1),
      date:    formatarData(r[1]),
      time:    r[2]   || '',
      name:    r[3]   || '',
      phone:   r[4]   || '',
      email:   r[5]   || '',
      servico: r[6]   || '',
      status:  r[7]   || 'Pendente',
      notes:   r[8]   || '',
    });
  }

  return lista;
}

// ═══════════════════════════════════════════════════════════════════
// 🚫 Dias Bloqueados
// ═══════════════════════════════════════════════════════════════════

function getBlockedDates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const aba   = sheet.getSheetByName('Bloqueios');
  if (!aba || aba.getLastRow() === 0) return [];

  return aba.getDataRange().getValues()
    .map(r => formatarData(r[0]))
    .filter(Boolean);
}

function blockDate(dados) {
  if (!dados.date) return jsonResponse({ success: false, error: 'date é obrigatório' });

  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  let aba = sheet.getSheetByName('Bloqueios');
  if (!aba) aba = sheet.insertSheet('Bloqueios');

  // Evita duplicata
  const existing = getBlockedDates();
  if (existing.includes(dados.date)) return jsonResponse({ success: true, message: 'Já bloqueado' });

  aba.appendRow([dados.date]);
  Logger.log('🚫 Dia bloqueado: ' + dados.date);
  return jsonResponse({ success: true });
}

function unblockDate(dados) {
  if (!dados.date) return jsonResponse({ success: false, error: 'date é obrigatório' });

  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const aba   = sheet.getSheetByName('Bloqueios');
  if (!aba) return jsonResponse({ success: false, error: 'Nenhum dia bloqueado' });

  const rows = aba.getDataRange().getValues();
  for (let i = 0; i < rows.length; i++) {
    if (formatarData(rows[i][0]) === dados.date) {
      aba.deleteRow(i + 1);
      Logger.log('✅ Dia desbloqueado: ' + dados.date);
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ success: false, error: 'Data não encontrada' });
}

// ═══════════════════════════════════════════════════════════════════
// 🔧 Helpers internos
// ═══════════════════════════════════════════════════════════════════

function getAbaAgendamentos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const aba   = sheet.getSheetByName(CONFIG.sheetsNames.agendamentos);
  if (!aba) throw new Error('Aba "Agendamentos" não encontrada. Execute setupPlanilha() primeiro.');
  return aba;
}

function getHorariosOcupados(date) {
  const aba  = getAbaAgendamentos();
  const rows = aba.getDataRange().getValues();
  const times = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] && !r[1]) continue;
    const rowDate   = formatarData(r[1]);
    const rowStatus = String(r[7] || 'Pendente');
    if (rowDate === date && rowStatus !== 'Cancelado') {
      const hora = formatarHora(r[2]);
      if (hora) times.push(hora);
    }
  }
  return times;
}

function formatarHora(valor) {
  if (!valor) return '';
  if (valor instanceof Date) {
    const h = String(valor.getHours()).padStart(2, '0');
    const m = String(valor.getMinutes()).padStart(2, '0');
    return h + ':' + m;
  }
  return String(valor);
}

function gerarId() {
  return Date.now();
}

function formatarData(valor) {
  if (!valor) return '';
  if (valor instanceof Date) return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(valor);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════════
// 📧 Email de confirmação (opcional)
// ═══════════════════════════════════════════════════════════════════

function enviarConfirmacaoEmail(dados) {
  const sheet    = SpreadsheetApp.getActiveSpreadsheet();
  const abaConf  = sheet.getSheetByName(CONFIG.sheetsNames.config);
  if (!abaConf) return;

  const emailElis = abaConf.getRange('B5').getValue();
  if (!emailElis) return;

  const assunto = '📅 Novo Agendamento — ' + dados.name;
  const corpo   = '<h2>Novo Agendamento!</h2>'
    + '<p><strong>Cliente:</strong> '  + dados.name  + '</p>'
    + '<p><strong>Data:</strong> '     + dados.date  + '</p>'
    + '<p><strong>Hora:</strong> '     + dados.time  + '</p>'
    + '<p><strong>Telefone:</strong> ' + dados.phone + '</p>';

  GmailApp.sendEmail(emailElis, assunto, '', { htmlBody: corpo });
  Logger.log('📧 Email enviado para ' + emailElis);
}

// ═══════════════════════════════════════════════════════════════════
// 🚀 SETUP INICIAL (execute uma vez)
// ═══════════════════════════════════════════════════════════════════

function setupPlanilha() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  criarAbas(sheet);
  formatarAbaAgendamentos(sheet);
  criarDashboard(sheet);
  criarConfiguracao(sheet);
  Logger.log('✅ Planilha configurada!');
}

function criarAbas(sheet) {
  [CONFIG.sheetsNames.agendamentos, CONFIG.sheetsNames.dashboard, CONFIG.sheetsNames.config].forEach(nome => {
    try { sheet.insertSheet(nome); } catch (_) {}
  });
}

function formatarAbaAgendamentos(sheet) {
  const aba = sheet.getSheetByName(CONFIG.sheetsNames.agendamentos);
  if (aba.getLastRow() === 0) {
    aba.appendRow(CONFIG.headers);
  }
  const hr = aba.getRange(1, 1, 1, CONFIG.headers.length);
  hr.setBackground(CONFIG.colors.headerBg);
  hr.setFontColor(CONFIG.colors.headerText);
  hr.setFontWeight('bold');
  hr.setFontSize(12);
  aba.setFrozenRows(1);

  const widths = [80, 110, 90, 150, 130, 150, 120, 100, 200, 150];
  widths.forEach((w, i) => aba.setColumnWidth(i + 1, w));

  const statusRange = aba.getRange('H2:H1000');
  const validacao   = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendente', 'Confirmado', 'Cancelado', 'Realizado'])
    .setAllowInvalid(false).build();
  statusRange.setDataValidation(validacao);
}

function criarDashboard(sheet) {
  const aba = sheet.getSheetByName(CONFIG.sheetsNames.dashboard);
  aba.clearContents();
  aba.getRange('A1').setValue('📊 DASHBOARD — ELIS MIRANDA').setFontSize(22).setFontWeight('bold').setFontColor(CONFIG.colors.headerBg);
  aba.getRange('A3').setValue('Total').setFontWeight('bold');
  aba.getRange('B3').setFormula('=COUNTA(Agendamentos!B2:B1000)');
  aba.getRange('A4').setValue('Confirmados');
  aba.getRange('B4').setFormula('=COUNTIF(Agendamentos!H2:H1000,"Confirmado")');
  aba.getRange('A5').setValue('Pendentes');
  aba.getRange('B5').setFormula('=COUNTIF(Agendamentos!H2:H1000,"Pendente")');
  aba.getRange('A6').setValue('Realizados');
  aba.getRange('B6').setFormula('=COUNTIF(Agendamentos!H2:H1000,"Realizado")');
}

function criarConfiguracao(sheet) {
  const aba = sheet.getSheetByName(CONFIG.sheetsNames.config);
  aba.clearContents();
  const configs = [
    ['Nome do Profissional', 'Elis Miranda'],
    ['Especialidade',        'Estética e Beleza'],
    ['Cidade',               'Porto Alegre, RS'],
    ['Telefone WhatsApp',    ''],
    ['Email',                'elis@elismiranda.com.br'],
    ['URL Webhook n8n',      ''],
    ['Token Evolution GO',   ''],
    ['Horário',              '9:00 - 17:00'],
  ];
  configs.forEach((c, i) => {
    aba.getRange(i + 1, 1).setValue(c[0]).setFontWeight('bold');
    aba.getRange(i + 1, 2).setValue(c[1]);
  });
  aba.setColumnWidth(1, 220);
  aba.setColumnWidth(2, 300);
}

// ═══════════════════════════════════════════════════════════════════
// ⏰ LEMBRETES AUTOMÁTICOS (trigger diário às 8h)
// ═══════════════════════════════════════════════════════════════════

function criarTriggers() {
  ScriptApp.newTrigger('enviarLembretes24hAntes').timeBased().atHour(8).everyDays(1).create();
  Logger.log('⏰ Trigger criado');
}

function enviarLembretes24hAntes() {
  const aba   = getAbaAgendamentos();
  const rows  = aba.getDataRange().getValues();
  const amanha = new Date();
  amanha.setDate(amanha.getDate() + 1);
  const dataAmanha = amanha.toLocaleDateString('pt-BR');

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (formatarData(r[1]) === dataAmanha && r[7] === 'Confirmado') {
      Logger.log('🔔 Lembrete para: ' + r[3] + ' — ' + r[2]);
    }
  }
}
