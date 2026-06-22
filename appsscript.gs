// 📊 GOOGLE SHEETS — AGENDAMENTO ELIS MIRANDA
// Versão corrigida: suporta JSON, form-encoded (no-cors) e todas as actions do painel

// ═══════════════════════════════════════════════════════════════════
// ⚙️ CONFIGURAÇÕES
// ═══════════════════════════════════════════════════════════════════

const CONFIG = {
  sheetsNames: {
    agendamentos:  'Agendamentos',
    dashboard:     'Dashboard',
    config:        'Configuração',
    servicos:      'Serviços',
    bloqueiosHora: 'BloqueiosHora'
  },
  colors: {
    headerBg:    '#E91E63',
    headerText:  '#FFFFFF',
    dashboardBg: '#FFF5F8',
    goldenAccent:'#D4AF37'
  },
  headers: ['ID', 'Data', 'Hora', 'Nome Cliente', 'Telefone', 'Email', 'Serviço', 'Status', 'Anotações', 'Data Agendamento', 'LembreteEnviado', 'DuracaoMin']
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
  if (action === 'servicos')      return jsonResponse({ servicos: getServicos() });
  if (action === 'blockedHours')  return jsonResponse({ hours: getBlockedHours() });

  return jsonResponse({ success: false, error: 'Ação desconhecida: ' + action });
}

// ═══════════════════════════════════════════════════════════════════
// 🌐 doPost — Recebe do formulário (form-encoded / no-cors) e do painel (URLSearchParams)
// ═══════════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    let dados = parseRequest(e);
    const action = dados.action || 'add';

    if (action === 'add')           return addAgendamento(dados);
    if (action === 'updateStatus')  return atualizarStatus(dados);
    if (action === 'blockDate')     return blockDate(dados);
    if (action === 'unblockDate')   return unblockDate(dados);
    if (action === 'addServico')    return addServico(dados);
    if (action === 'updateServico') return updateServico(dados);
    if (action === 'deleteServico')      return deleteServico(dados);
    if (action === 'deleteAgendamento')  return deleteAgendamento(dados);
    if (action === 'blockHour')          return blockHour(dados);
    if (action === 'unblockHour')   return unblockHour(dados);

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

function parseDuracaoMin(str) {
  if (!str) return 30;
  str = String(str).toLowerCase().trim();
  let mins = 0;
  const h = str.match(/(\d+)\s*h/);
  const m = str.match(/(\d+)\s*min/);
  if (h) mins += parseInt(h[1]) * 60;
  if (m) mins += parseInt(m[1]);
  if (!mins) { const n = parseInt(str); mins = isNaN(n) ? 30 : n; }
  return Math.max(30, mins);
}

function addAgendamento(dados) {
  if (!dados.name || !dados.phone || !dados.date || !dados.time) {
    return jsonResponse({ success: false, error: 'Campos obrigatórios: name, phone, date, time' });
  }

  const aba = getAbaAgendamentos();
  const novoId = gerarId();
  const duracaoMin = parseDuracaoMin(dados.duracao);

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
    new Date().toLocaleDateString('pt-BR'),       // Data do agendamento
    '',                                          // LembreteEnviado
    duracaoMin                                   // DuracaoMin
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
// 🗑️ Deletar agendamento
// ═══════════════════════════════════════════════════════════════════

function deleteAgendamento(dados) {
  if (!dados.id) {
    return jsonResponse({ success: false, error: 'id é obrigatório' });
  }

  const aba  = getAbaAgendamentos();
  const rows = aba.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(dados.id)) {
      aba.deleteRow(i + 1);
      Logger.log('🗑️ Agendamento deletado: ID ' + dados.id);
      return jsonResponse({ success: true });
    }
  }

  return jsonResponse({ success: false, error: 'ID não encontrado: ' + dados.id });
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
      time:    formatarHora(r[2]),
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
// ⏰ Horários Bloqueados — CRUD
// ═══════════════════════════════════════════════════════════════════

function getBlockedHours() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const aba   = sheet.getSheetByName(CONFIG.sheetsNames.bloqueiosHora);
  if (!aba || aba.getLastRow() === 0) return [];

  return aba.getDataRange().getValues()
    .map(r => ({ date: formatarData(r[0]), hora: String(r[1] || '') }))
    .filter(r => r.date && r.hora);
}

function blockHour(dados) {
  if (!dados.date || !dados.hora) return jsonResponse({ success: false, error: 'date e hora são obrigatórios' });

  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  let aba = sheet.getSheetByName(CONFIG.sheetsNames.bloqueiosHora);
  if (!aba) aba = sheet.insertSheet(CONFIG.sheetsNames.bloqueiosHora);

  // Evita duplicata
  const existing = getBlockedHours();
  if (existing.some(r => r.date === dados.date && r.hora === dados.hora)) {
    return jsonResponse({ success: true, message: 'Já bloqueado' });
  }

  aba.appendRow([dados.date, dados.hora]);
  Logger.log('⏰ Horário bloqueado: ' + dados.date + ' ' + dados.hora);
  return jsonResponse({ success: true });
}

function unblockHour(dados) {
  if (!dados.date || !dados.hora) return jsonResponse({ success: false, error: 'date e hora são obrigatórios' });

  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const aba   = sheet.getSheetByName(CONFIG.sheetsNames.bloqueiosHora);
  if (!aba)   return jsonResponse({ success: false, error: 'Nenhum horário bloqueado' });

  const rows = aba.getDataRange().getValues();
  for (let i = 0; i < rows.length; i++) {
    if (formatarData(rows[i][0]) === dados.date && String(rows[i][1]) === dados.hora) {
      aba.deleteRow(i + 1);
      Logger.log('✅ Horário desbloqueado: ' + dados.date + ' ' + dados.hora);
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ success: false, error: 'Horário não encontrado' });
}

// ═══════════════════════════════════════════════════════════════════
// 💅 Serviços — CRUD
// ═══════════════════════════════════════════════════════════════════

function getServicos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const aba   = sheet.getSheetByName(CONFIG.sheetsNames.servicos);
  if (!aba || aba.getLastRow() <= 1) return [];

  const rows = aba.getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    lista.push({
      id:        String(r[0]),
      categoria: r[1] || '',
      servico:   r[2] || '',
      preco:     Number(r[3]) || 0,
      duracao:   r[4] || '',
      ativo:     r[5] !== false && r[5] !== 'false' && r[5] !== 0,
      obs:       r[6] || ''
    });
  }
  return lista;
}

function addServico(dados) {
  if (!dados.categoria || !dados.servico || !dados.preco) {
    return jsonResponse({ success: false, error: 'categoria, servico e preco são obrigatórios' });
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  let aba = sheet.getSheetByName(CONFIG.sheetsNames.servicos);
  if (!aba) { setupServicos(); aba = sheet.getSheetByName(CONFIG.sheetsNames.servicos); }

  const novoId = Date.now();
  aba.appendRow([novoId, dados.categoria, dados.servico, Number(dados.preco), dados.duracao || '', true, dados.obs || '']);
  return jsonResponse({ success: true, id: novoId });
}

function updateServico(dados) {
  if (!dados.id) return jsonResponse({ success: false, error: 'id é obrigatório' });
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const aba   = sheet.getSheetByName(CONFIG.sheetsNames.servicos);
  if (!aba)   return jsonResponse({ success: false, error: 'Aba Serviços não encontrada' });

  const rows = aba.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(dados.id)) {
      if (dados.categoria !== undefined) aba.getRange(i + 1, 2).setValue(dados.categoria);
      if (dados.servico   !== undefined) aba.getRange(i + 1, 3).setValue(dados.servico);
      if (dados.preco     !== undefined) aba.getRange(i + 1, 4).setValue(Number(dados.preco));
      if (dados.duracao   !== undefined) aba.getRange(i + 1, 5).setValue(dados.duracao);
      if (dados.ativo     !== undefined) aba.getRange(i + 1, 6).setValue(dados.ativo === 'true' || dados.ativo === true);
      if (dados.obs       !== undefined) aba.getRange(i + 1, 7).setValue(dados.obs);
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ success: false, error: 'Serviço não encontrado: ' + dados.id });
}

function deleteServico(dados) {
  if (!dados.id) return jsonResponse({ success: false, error: 'id é obrigatório' });
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const aba   = sheet.getSheetByName(CONFIG.sheetsNames.servicos);
  if (!aba)   return jsonResponse({ success: false, error: 'Aba Serviços não encontrada' });

  const rows = aba.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(dados.id)) {
      aba.deleteRow(i + 1);
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ success: false, error: 'Serviço não encontrado: ' + dados.id });
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
      if (hora) {
        const duracaoMin = Number(r[11]) || 30;
        times.push({ time: hora, duracao: duracaoMin });
      }
    }
  }

  // Inclui horários bloqueados manualmente (duração mínima 30min)
  try {
    const bloqueados = getBlockedHours().filter(b => b.date === date);
    bloqueados.forEach(b => {
      if (!times.some(t => t.time === b.hora)) {
        times.push({ time: b.hora, duracao: 30 });
      }
    });
  } catch (_) {}

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
  setupServicos();
  Logger.log('✅ Planilha configurada!');
}

function criarAbas(sheet) {
  [CONFIG.sheetsNames.agendamentos, CONFIG.sheetsNames.dashboard, CONFIG.sheetsNames.config, CONFIG.sheetsNames.servicos].forEach(nome => {
    try { sheet.insertSheet(nome); } catch (_) {}
  });
}

function setupServicos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  let aba = sheet.getSheetByName(CONFIG.sheetsNames.servicos);
  if (!aba) aba = sheet.insertSheet(CONFIG.sheetsNames.servicos);

  if (aba.getLastRow() > 0) return; // já tem dados

  const headers = ['ID', 'Categoria', 'Serviço', 'Preço', 'Duração', 'Ativo', 'Observação'];
  aba.appendRow(headers);
  const hr = aba.getRange(1, 1, 1, headers.length);
  hr.setBackground('#C9A96E').setFontColor('#FFFFFF').setFontWeight('bold');
  aba.setFrozenRows(1);

  const servicos = [
    [Date.now()+1,  'Extensão de Cílios', 'Volume Brasileiro',                   150, '~2h',    true, ''],
    [Date.now()+2,  'Extensão de Cílios', 'Volume Egípcio 3D',                   150, '~2h',    true, ''],
    [Date.now()+3,  'Extensão de Cílios', 'Volume Egípcio 5D',                   170, '~2h',    true, ''],
    [Date.now()+4,  'Extensão de Cílios', 'Volume Fox',                           180, '~2h',    true, ''],
    [Date.now()+5,  'Extensão de Cílios', 'Copping',                              220, '~2h30',  true, ''],
    [Date.now()+6,  'Manutenção de Cílios', 'Manutenção até 20 dias',             85,  '~1h',    true, ''],
    [Date.now()+7,  'Manutenção de Cílios', 'Manutenção até 30 dias',             100, '~1h',    true, ''],
    [Date.now()+8,  'Design e Epilação', 'Design de Sobrancelhas',                30,  '~20min', true, ''],
    [Date.now()+9,  'Design e Epilação', 'Epilação Buço',                         20,  '~5min',  true, ''],
    [Date.now()+10, 'Design e Epilação', 'Epilação Queixo',                       15,  '~10min', true, ''],
    [Date.now()+11, 'Design e Epilação', 'Epilação Nariz',                        15,  '~10min', true, ''],
    [Date.now()+12, 'Design e Epilação', 'Combo Epilação',                        60,  '~30min', true, ''],
    [Date.now()+13, 'Design e Epilação', 'Epilação Total Rosto',                  80,  '~1h',    true, ''],
    [Date.now()+14, 'Manicure e Unhas', 'Manicure Simples',                       30,  '~1h',    true, ''],
    [Date.now()+15, 'Manicure e Unhas', 'Esmaltação em Gel Mãos',                 80,  '~2h',    true, ''],
    [Date.now()+16, 'Manicure e Unhas', 'Alongamento Mãos',                       150, '~3h',    true, ''],
    [Date.now()+17, 'Manicure e Unhas', 'Manutenção Alongamento até 20 dias',     100, '~3h',    true, ''],
    [Date.now()+18, 'Manicure e Unhas', 'Esmaltação em Gel Pés',                  80,  '~2h',    true, ''],
    [Date.now()+19, 'Micropigmentação', 'Sobrancelhas',                           250, '~2h',    true, ''],
    [Date.now()+20, 'Micropigmentação', 'Retoque após 1 mês',                     100, '~2h',    true, ''],
    [Date.now()+21, 'Micropigmentação', 'Correção',                               450, '~2h',    true, ''],
    [Date.now()+22, 'Micropigmentação', 'Retoques 2x a cada 30 dias',             100, '~3h',    true, 'Valor cada sessão'],
    [Date.now()+23, 'Estética',         'Tratamento de Estrias',                  250, '~1h',    true, 'Mediante avaliação presencial — valor por sessão'],
  ];

  const widths = [80, 180, 240, 80, 80, 60, 260];
  widths.forEach((w, i) => aba.setColumnWidth(i + 1, w));

  servicos.forEach(row => aba.appendRow(row));
  Logger.log('✅ Aba Serviços criada com ' + servicos.length + ' serviços');
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
  ScriptApp.newTrigger('enviarAgendaDia').timeBased().atHour(7).everyDays(1).create();
  Logger.log('⏰ Triggers criados: lembrete 8h + agenda 7h');
}

function enviarAgendaDia() {
  const aba   = getAbaAgendamentos();
  const rows  = aba.getDataRange().getValues();
  const hoje  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const [year, month, day] = hoje.split('-');
  const dateFormatted = day + '/' + month + '/' + year;

  const weekdays = ['domingo','segunda-feira','terça-feira','quarta-feira','quinta-feira','sexta-feira','sábado'];
  const weekday  = weekdays[new Date(hoje + 'T12:00:00').getDay()];

  const AGENDA_WEBHOOK = 'https://flow.ezstudio.com.br/webhook/agenda-dia-elis';

  const lista = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    const rowDate   = formatarData(r[1]);
    const rowStatus = String(r[7] || '');
    if (rowDate === hoje && rowStatus !== 'Cancelado') {
      lista.push({
        hora:    formatarHora(r[2]),
        nome:    String(r[3] || ''),
        servico: String(r[6] || ''),
        status:  rowStatus
      });
    }
  }

  lista.sort((a, b) => a.hora.localeCompare(b.hora));

  try {
    UrlFetchApp.fetch(AGENDA_WEBHOOK, {
      method:      'POST',
      contentType: 'application/json',
      payload:     JSON.stringify({ date: hoje, dateFormatted, weekday, total: lista.length, agendamentos: lista }),
      muteHttpExceptions: true
    });
    Logger.log('✅ Agenda do dia enviada: ' + lista.length + ' atendimento(s)');
  } catch (err) {
    Logger.log('❌ Erro agenda dia: ' + err);
  }
}

function enviarLembretes24hAntes() {
  const aba    = getAbaAgendamentos();
  const rows   = aba.getDataRange().getValues();
  const amanha = new Date();
  amanha.setDate(amanha.getDate() + 1);
  const dataAmanha = Utilities.formatDate(amanha, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const LEMBRETE_WEBHOOK = 'https://flow.ezstudio.com.br/webhook/lembrete-elis';
  const COL_LEMBRETE = 11; // coluna K — LembreteEnviado

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;

    const rowDate        = formatarData(r[1]);
    const rowStatus      = String(r[7] || '');
    const lembreteJaEnviado = String(r[COL_LEMBRETE - 1] || '');

    if (rowDate !== dataAmanha) continue;
    if (rowStatus !== 'Confirmado' && rowStatus !== 'Agendado') continue;
    if (lembreteJaEnviado === 'true') continue;

    const payload = {
      name:    String(r[3] || ''),
      phone:   String(r[4] || ''),
      date:    rowDate,
      time:    formatarHora(r[2]),
      servico: String(r[6] || '')
    };

    try {
      UrlFetchApp.fetch(LEMBRETE_WEBHOOK, {
        method:      'POST',
        contentType: 'application/json',
        payload:     JSON.stringify(payload),
        muteHttpExceptions: true
      });
      aba.getRange(i + 1, COL_LEMBRETE).setValue('true');
      Logger.log('✅ Lembrete enviado: ' + payload.name + ' — ' + rowDate + ' ' + payload.time);
    } catch (err) {
      Logger.log('❌ Erro lembrete ' + payload.name + ': ' + err);
    }
  }
}
