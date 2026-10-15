// ============================================================
// Code.gs — SERVER-SIDE (Google Apps Script)
// Sistema de Gestão de Orçamentos via WhatsApp
// ============================================================

const SHEET_NAME = 'Orçamentos';
const DASH_NAME = 'Dashboard';
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbz3uHYnblqw38Jda8E6FcG7c1TkQEhOTeFR6OORwV0QyufXvI8KRwaejw39vhaxlRi7/exec'; 

// ============================================================
// ROUTING (doGet)
// ============================================================
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};

  if (params.id) {
    var tpl = HtmlService.createTemplateFromFile('Cliente');
    tpl.orcamentoId = params.id;
    return tpl.evaluate()
      .setTitle('Orçamento - Resposta')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  if (params.page === 'bi') {
    return HtmlService.createHtmlOutputFromFile('BI')
      .setTitle('Dashboard Executivo')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  return HtmlService.createHtmlOutputFromFile('Formulario')
    .setTitle('Novo Orçamento')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ============================================================
// HELPER: Obter URL do Web App
// ============================================================
function getWebAppUrl() {
  if (WEB_APP_URL && WEB_APP_URL.length > 0) return WEB_APP_URL;
  try { return ScriptApp.getService().getUrl(); } catch (err) { return ''; }
}

// ============================================================
// HELPER: Gerar ID único (8 chars)
// ============================================================
function gerarId() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  var id = '';
  for (var i = 0; i < 8; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return id;
}

// ============================================================
// HELPER: Limpar telefone para WhatsApp
// ============================================================
function limparTelefone(tel) {
  if (!tel) return '';
  var limpo = tel.toString().replace(/\D/g, '');
  if (limpo.length === 11 || limpo.length === 10) {
    limpo = '55' + limpo;
  }
  return limpo;
}

// ============================================================
// HELPER: Formatar data
// ============================================================
function formatarData(data) {
  if (!data) return '';
  var d = (data instanceof Date) ? data : new Date(data);
  if (isNaN(d.getTime())) return '';
  return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear();
}

function formatarDataHora(data) {
  if (!data) return '';
  var d = (data instanceof Date) ? data : new Date(data);
  if (isNaN(d.getTime())) return '';
  return formatarData(d) + ' ' + ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2);
}

function formatarMoeda(valor) {
  if (!valor || isNaN(valor)) return 'R$ 0,00';
  return 'R$ ' + Number(valor).toFixed(2).replace('.', ',');
}

// ============================================================
// MONTAR MENSAGEM WHATSAPP
// ============================================================
function montarMensagemWhatsApp(dados, linkResposta, valor) {
  var linhas = [];
  linhas.push('*ORÇAMENTO - EVENTO*');
  linhas.push('');
  linhas.push('Olá, ' + (dados.nomeCliente || 'Cliente') + '!');
  linhas.push('');
  linhas.push('Evento: ' + (dados.tipoEvento || ''));
  if (dados.dataEvento) {
    var d = new Date(dados.dataEvento + 'T00:00:00');
    linhas.push('Data do Evento: ' + formatarData(d));
  }
  if (dados.convidados) {
    linhas.push('Convidados: ' + dados.convidados);
  }
  if (dados.local) {
    linhas.push('Local: ' + dados.local);
  }
  linhas.push('Valor: *' + formatarMoeda(valor) + '*');
  if (dados.descricao) {
    linhas.push('');
    linhas.push('Serviços: ' + dados.descricao);
  }
  linhas.push('');
  linhas.push('Para responder, clique no link abaixo:');
  linhas.push(linkResposta);
  linhas.push('');
  linhas.push('Obrigado!');
  return linhas.join('\n');
}

// ============================================================
// SALVAR ORÇAMENTO
// ============================================================
function salvarOrcamento(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, error: 'Aba não encontrada' };

    var id = gerarId();
    var agora = new Date();

    var dataEvento = null;
    if (dados.dataEvento) {
      dataEvento = new Date(dados.dataEvento + 'T00:00:00');
    }

    var valor = parseFloat(dados.valor) || 0;
    var convidados = parseInt(dados.convidados) || 0;

    var linha = [
      agora,
      dados.nomeCliente || '',
      dados.telefone || '',
      dados.tipoEvento || 'Outro',
      dataEvento,
      convidados,
      dados.local || '',
      valor,
      dados.descricao || '',
      'Pendente',
      '',
      '',
      id
    ];

    sheet.appendRow(linha);

    var lastRow = sheet.getLastRow();
    try { sheet.getRange(lastRow, 8, 1, 1).setNumberFormat('"R$" #,##0.00'); } catch (e) {}
    try { sheet.getRange(lastRow, 6, 1, 1).setNumberFormat('#,##0'); } catch (e) {}

    atualizarDashboard();

    var baseUrl = getWebAppUrl();
    var linkResposta = baseUrl + '?id=' + id;
    var telefoneLimpo = limparTelefone(dados.telefone);
    var msg = montarMensagemWhatsApp(dados, linkResposta, valor);
    var linkWhatsApp = 'https://wa.me/' + telefoneLimpo + '?text=' + encodeURIComponent(msg);

    return {
      success: true,
      id: id,
      linkResposta: linkResposta,
      linkWhatsApp: linkWhatsApp,
      mensagem: msg
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// GET ORÇAMENTO (por ID) — para página do cliente
// ============================================================
function getOrcamento(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, error: 'Aba não encontrada' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][12] === id) {
        return {
          success: true,
          orcamento: {
            nomeCliente: data[i][1],
            telefone: data[i][2],
            tipoEvento: data[i][3],
            dataEvento: data[i][4] ? formatarData(data[i][4]) : '',
            convidados: data[i][5],
            local: data[i][6],
            valor: data[i][7],
            descricao: data[i][8],
            status: data[i][9],
            dataResposta: data[i][10] ? formatarDataHora(data[i][10]) : '',
            resposta: data[i][11],
            id: id
          }
        };
      }
    }
    return { success: false, error: 'Orçamento não encontrado' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// REGISTRAR RESPOSTA (cliente)
// ============================================================
function registrarResposta(id, resposta) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, error: 'Aba não encontrada' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][12] === id) {
        var statusMap = {
          'aceitar': 'Aceito',
          'negociar': 'Negociar',
          'recusar': 'Negado'
        };
        var novoStatus = statusMap[resposta] || 'Pendente';

        // Verifica se já foi respondido
        if (data[i][9] !== 'Pendente' && data[i][9] !== '') {
          return { success: false, jaRespondido: true, statusAtual: data[i][9] };
        }

        sheet.getRange(i + 1, 10).setValue(novoStatus);
        sheet.getRange(i + 1, 11).setValue(new Date());
        sheet.getRange(i + 1, 12).setValue('Cliente respondeu: ' + novoStatus);

        try { sheet.getRange(i + 1, 11, 1, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss'); } catch (e) {}

        atualizarDashboard();

        return { success: true, status: novoStatus };
      }
    }
    return { success: false, error: 'Orçamento não encontrado' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// ATUALIZAR DASHBOARD (cálculo em JS, NÃO usa fórmulas)
// ============================================================
function atualizarDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    var dash = ss.getSheetByName(DASH_NAME);
    if (!sheet || !dash) return;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      dash.getRange(4, 2, 6, 1).setValues([[0],[0],[0],[0],[0],[0]]);
      dash.getRange(12, 2, 4, 1).setValues([[0],[0],[0],[0]]);
      dash.getRange(18, 2, 4, 1).setValues([[0],[0],[0],[0]]);
      return;
    }

    var orcamentos = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][12] && data[i][12] !== '') {
        orcamentos.push({
          dataEnvio: data[i][0],
          nome: data[i][1],
          telefone: data[i][2],
          tipo: data[i][3],
          dataEvento: data[i][4],
          convidados: data[i][5] || 0,
          local: data[i][6],
          valor: Number(data[i][7]) || 0,
          descricao: data[i][8],
          status: data[i][9] || 'Pendente',
          dataResposta: data[i][10],
          resposta: data[i][11],
          id: data[i][12]
        });
      }
    }

    var total = orcamentos.length;
    var pendentes = 0, aceitos = 0, negados = 0, negociar = 0;
    var valFaturado = 0, valAberto = 0, valNegociar = 0;
    var convidadosTotal = 0, eventosConfirmados = 0;

    var porTipo = {
      'Aniversário': { qtd: 0, fat: 0, aceitos: 0 },
      'Casamento':   { qtd: 0, fat: 0, aceitos: 0 },
      'Corporativo': { qtd: 0, fat: 0, aceitos: 0 },
      'Outro':       { qtd: 0, fat: 0, aceitos: 0 }
    };

    for (var j = 0; j < orcamentos.length; j++) {
      var o = orcamentos[j];
      var tipoKey = porTipo[o.tipo] ? o.tipo : 'Outro';
      porTipo[tipoKey].qtd++;

      if (o.status === 'Pendente') {
        pendentes++;
        valAberto += o.valor;
      } else if (o.status === 'Aceito') {
        aceitos++;
        valFaturado += o.valor;
        eventosConfirmados++;
        convidadosTotal += o.convidados;
        porTipo[tipoKey].fat += o.valor;
        porTipo[tipoKey].aceitos++;
      } else if (o.status === 'Negado') {
        negados++;
      } else if (o.status === 'Negociar') {
        negociar++;
        valNegociar += o.valor;
      }
    }

    var taxaConv = total > 0 ? (aceitos / total * 100) : 0;
    var ticketMedio = aceitos > 0 ? (valFaturado / aceitos) : 0;
    var mediaConvidados = eventosConfirmados > 0 ? (convidadosTotal / eventosConfirmados) : 0;

    var agora = new Date();
    var limite30 = new Date(agora.getTime() + 30 * 24 * 60 * 60 * 1000);
    var proximos30 = 0;
    var eventosAgenda = [];

    for (var k = 0; k < orcamentos.length; k++) {
      if (orcamentos[k].status === 'Aceito' && orcamentos[k].dataEvento) {
        var dEvt = new Date(orcamentos[k].dataEvento);
        var diasAte = Math.ceil((dEvt - agora) / (24 * 60 * 60 * 1000));
        if (diasAte >= 0 && dEvt <= limite30) proximos30++;
        if (diasAte >= 0) {
          eventosAgenda.push({
            dataEvento: dEvt,
            tipo: orcamentos[k].tipo,
            nome: orcamentos[k].nome,
            convidados: orcamentos[k].convidados,
            local: orcamentos[k].local,
            diasAte: diasAte
          });
        }
      }
    }
    eventosAgenda.sort(function(a, b) { return a.dataEvento - b.dataEvento; });

    // --- SEÇÃO 1: KPIs ---
    dash.getRange(4, 2, 6, 1).setValues([
      [total],
      [pendentes],
      [aceitos],
      [negados],
      [negociar],
      [taxaConv / 100]
    ]);
    try { dash.getRange(9, 2).setNumberFormat('0.0%'); } catch (e) {}

    // --- SEÇÃO 2: Financeiro ---
    dash.getRange(12, 2, 4, 1).setValues([
      [valFaturado],
      [valAberto],
      [valNegociar],
      [ticketMedio]
    ]);
    try { dash.getRange(12, 2, 4, 1).setNumberFormat('"R$" #,##0.00'); } catch (e) {}

    // --- SEÇÃO 3: Eventos ---
    dash.getRange(18, 2, 4, 1).setValues([
      [eventosConfirmados],
      [proximos30],
      [convidadosTotal],
      [Math.round(mediaConvidados)]
    ]);

    // --- SEÇÃO 4: Por Tipo (CORRIGIDO: 3 colunas, não 4) ---
    var tiposArr = ['Aniversário', 'Casamento', 'Corporativo', 'Outro'];
    for (var t = 0; t < tiposArr.length; t++) {
      var tk = tiposArr[t];
      var qtd = porTipo[tk].qtd;
      var fat = porTipo[tk].fat;
      var conv = qtd > 0 ? (porTipo[tk].aceitos / qtd) : 0;
      dash.getRange(25 + t, 2, 1, 3).setValues([[qtd, fat, conv]]); // ← AQUI: 3 colunas
    }
    try { dash.getRange(25, 3, 4, 1).setNumberFormat('"R$" #,##0.00'); } catch (e) {}
    try { dash.getRange(25, 4, 4, 1).setNumberFormat('0.0%'); } catch (e) {}

    // --- SEÇÃO 5: Projeções ---
    var primeiraData = orcamentos.length > 0 ? new Date(orcamentos[0].dataEnvio) : agora;
    var diasDecorridos = Math.max(1, Math.ceil((agora - primeiraData) / (24 * 60 * 60 * 1000)));
    var mediaDiaria = total / diasDecorridos;
    var projMensal = mediaDiaria * 30;
    var fatRealista = (valFaturado / diasDecorridos) * 30;
    var fatPessimista = fatRealista * 0.70;
    var fatOtimista = fatRealista * 1.15;

    var limite7 = new Date(agora.getTime() - 7 * 24 * 60 * 60 * 1000);
    var enviados7 = 0, aceitos7 = 0;
    for (var m = 0; m < orcamentos.length; m++) {
      var dEnv = new Date(orcamentos[m].dataEnvio);
      if (dEnv >= limite7) {
        enviados7++;
        if (orcamentos[m].status === 'Aceito') aceitos7++;
      }
    }
    var conv7 = enviados7 > 0 ? (aceitos7 / enviados7 * 100) : 0;
    var tendencia = 'Estável';
    if (conv7 > taxaConv + 5) tendencia = 'Melhorando';
    else if (conv7 < taxaConv - 5) tendencia = 'Declinando';

    dash.getRange(31, 2, 8, 1).setValues([
      [diasDecorridos],
      [mediaDiaria.toFixed(1)],
      [projMensal.toFixed(0)],
      [fatRealista],
      [fatPessimista],
      [fatOtimista],
      [conv7.toFixed(1) + '%'],
      [tendencia]
    ]);
    try { dash.getRange(34, 2, 3, 1).setNumberFormat('"R$" #,##0.00'); } catch (e) {}

        // --- DADOS PARA GRÁFICOS ---
    // Pizza: E6:F9 — ordem: Aceito, Pendente, Negado, Negociar
    dash.getRange(6, 5, 4, 2).setValues([
      ['Aceito', aceitos],
      ['Pendente', pendentes],
      ['Negado', negados],
      ['Negociar', negociar]
    ]);

    // Barras valores: F12:I13 (cada status = série com cor própria)
    // Header já existe (F12:I12). Só atualizar valores em F13:I13
    dash.getRange(13, 6, 1, 4).setValues([[
      valFaturado,  // F13: Aceito
      valAberto,    // G13: Pendente
      0,            // H13: Negado
      valNegociar   // I13: Negociar
    ]]);
    try { dash.getRange(13, 6, 1, 4).setNumberFormat('"R$" #,##0.00'); } catch (e) {}

    // Barras tipos: E18:F21
    dash.getRange(18, 5, 4, 2).setValues([
      ['Aniversário', porTipo['Aniversário'].qtd],
      ['Casamento', porTipo['Casamento'].qtd],
      ['Corporativo', porTipo['Corporativo'].qtd],
      ['Outro', porTipo['Outro'].qtd]
    ]);

    // --- SEÇÃO 6: AGENDA ---
    dash.getRange(42, 1, 100, 6).clearContent();
    dash.getRange(42, 1, 100, 6).clearFormat();

    for (var n = 0; n < eventosAgenda.length; n++) {
      var ev = eventosAgenda[n];
      var row = 42 + n;
      dash.getRange(row, 1, 1, 6).setValues([[
        formatarData(ev.dataEvento),
        ev.tipo,
        ev.nome,
        ev.convidados,
        ev.local || '',
        ev.diasAte + ' dia(s)'
      ]]);

      if (ev.diasAte <= 7) {
        dash.getRange(row, 1, 1, 6).setBackground('#ffcdd2');
      } else if (ev.diasAte <= 30) {
        dash.getRange(row, 1, 1, 6).setBackground('#ffe0b2');
      }
    }

  } catch (err) {
    Logger.log('Erro atualizarDashboard: ' + err.toString());
  }
}

// ============================================================
// GET MÉTRICAS BI (retorna JSON para o BI web)
// ============================================================
function getMetricasBI() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, error: 'Aba não encontrada' };

    var data = sheet.getDataRange().getValues();
    var orcamentos = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][12]) {
        orcamentos.push({
          dataEnvio: data[i][0],
          nome: data[i][1],
          telefone: data[i][2],
          tipo: data[i][3],
          dataEvento: data[i][4],
          convidados: data[i][5] || 0,
          local: data[i][6],
          valor: Number(data[i][7]) || 0,
          status: data[i][9] || 'Pendente',
          dataResposta: data[i][10]
        });
      }
    }

    var total = orcamentos.length;
    var pendentes = 0, aceitos = 0, negados = 0, negociar = 0;
    var valFaturado = 0, valAberto = 0, valNegociar = 0;
    var convidadosTotal = 0, eventosConfirmados = 0;
    var porTipo = { 'Aniversário': {qtd:0,fat:0,ace:0}, 'Casamento': {qtd:0,fat:0,ace:0}, 'Corporativo': {qtd:0,fat:0,ace:0}, 'Outro': {qtd:0,fat:0,ace:0} };

    for (var j = 0; j < orcamentos.length; j++) {
      var o = orcamentos[j];
      var tk = porTipo[o.tipo] ? o.tipo : 'Outro';
      porTipo[tk].qtd++;
      if (o.status === 'Pendente') { pendentes++; valAberto += o.valor; }
      else if (o.status === 'Aceito') { aceitos++; valFaturado += o.valor; eventosConfirmados++; convidadosTotal += o.convidados; porTipo[tk].fat += o.valor; porTipo[tk].ace++; }
      else if (o.status === 'Negado') { negados++; }
      else if (o.status === 'Negociar') { negociar++; valNegociar += o.valor; }
    }

    var taxaConv = total > 0 ? (aceitos / total * 100) : 0;
    var ticketMedio = aceitos > 0 ? (valFaturado / aceitos) : 0;
    var mediaConvidados = eventosConfirmados > 0 ? Math.round(convidadosTotal / eventosConfirmados) : 0;

    // Evolução 14 dias
    var agora = new Date();
    var evolucao = [];
    for (var d = 13; d >= 0; d--) {
      var dia = new Date(agora.getTime() - d * 24 * 60 * 60 * 1000);
      dia.setHours(0, 0, 0, 0);
      var diaFim = new Date(dia.getTime() + 24 * 60 * 60 * 1000);
      var enviados = 0, aceitosDia = 0;
      for (var e = 0; e < orcamentos.length; e++) {
        var dEnv = new Date(orcamentos[e].dataEnvio);
        if (dEnv >= dia && dEnv < diaFim) enviados++;
        if (orcamentos[e].status === 'Aceito' && orcamentos[e].dataResposta) {
          var dResp = new Date(orcamentos[e].dataResposta);
          if (dResp >= dia && dResp < diaFim) aceitosDia++;
        }
      }
      evolucao.push({
        data: ('0' + dia.getDate()).slice(-2) + '/' + ('0' + (dia.getMonth() + 1)).slice(-2),
        enviados: enviados,
        aceitos: aceitosDia
      });
    }

    // Funil
    var respondidos = aceitos + negados + negociar;

    // Projeções
    var primeiraData = orcamentos.length > 0 ? new Date(orcamentos[0].dataEnvio) : agora;
    var diasDecorridos = Math.max(1, Math.ceil((agora - primeiraData) / (24 * 60 * 60 * 1000)));
    var mediaDiaria = total / diasDecorridos;
    var projMensal = mediaDiaria * 30;
    var fatRealista = (valFaturado / diasDecorridos) * 30;
    var fatPessimista = fatRealista * 0.70;
    var fatOtimista = fatRealista * 1.15;

    // Conversão 7 dias
    var limite7 = new Date(agora.getTime() - 7 * 24 * 60 * 60 * 1000);
    var enviados7 = 0, aceitos7 = 0;
    for (var m = 0; m < orcamentos.length; m++) {
      if (new Date(orcamentos[m].dataEnvio) >= limite7) {
        enviados7++;
        if (orcamentos[m].status === 'Aceito') aceitos7++;
      }
    }
    var conv7 = enviados7 > 0 ? (aceitos7 / enviados7 * 100) : 0;
    var tendencia = 'Estável';
    if (conv7 > taxaConv + 5) tendencia = 'Melhorando';
    else if (conv7 < taxaConv - 5) tendencia = 'Declinando';

    // Próximos eventos
    var limite30 = new Date(agora.getTime() + 30 * 24 * 60 * 60 * 1000);
    var proximos30 = 0;
    var eventosAgenda = [];
    for (var k = 0; k < orcamentos.length; k++) {
      if (orcamentos[k].status === 'Aceito' && orcamentos[k].dataEvento) {
        var dEvt = new Date(orcamentos[k].dataEvento);
        var diasAte = Math.ceil((dEvt - agora) / (24 * 60 * 60 * 1000));
        if (diasAte >= 0 && dEvt <= limite30) proximos30++;
        if (diasAte >= 0) {
          eventosAgenda.push({
            data: formatarData(dEvt),
            tipo: orcamentos[k].tipo,
            nome: orcamentos[k].nome,
            convidados: orcamentos[k].convidados,
            local: orcamentos[k].local || '',
            diasAte: diasAte
          });
        }
      }
    }
    eventosAgenda.sort(function(a, b) { return a.diasAte - b.diasAte; });

    return {
      success: true,
      kpis: {
        total: total, pendentes: pendentes, aceitos: aceitos,
        negados: negados, negociar: negociar, taxaConv: taxaConv,
        valFaturado: valFaturado, valAberto: valAberto, valNegociar: valNegociar,
        ticketMedio: ticketMedio, mediaConvidados: mediaConvidados,
        eventosConfirmados: eventosConfirmados, proximos30: proximos30,
        convidadosTotal: convidadosTotal
      },
      porTipo: porTipo,
      evolucao: evolucao,
      funil: { enviados: total, respondidos: respondidos, aceitos: aceitos },
      projecoes: {
        diasDecorridos: diasDecorridos, mediaDiaria: mediaDiaria,
        projMensal: projMensal, fatRealista: fatRealista,
        fatPessimista: fatPessimista, fatOtimista: fatOtimista,
        conv7: conv7, tendencia: tendencia
      },
      eventosAgenda: eventosAgenda,
      atualizadoEm: formatarDataHora(agora)
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// SETUP: Configurar planilha, dashboard e gráficos
// ============================================================
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- ABA ORÇAMENTOS ---
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  var headers = [
    'Data Envio', 'Nome do Cliente', 'Telefone/WhatsApp',
    'Tipo de Evento', 'Data do Evento', 'Qtd Convidados',
    'Local do Evento', 'Valor Total', 'Descrição/Serviços',
    'Status', 'Data Resposta', 'Resposta', 'ID Único'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff').setHorizontalAlignment('center');

  var larguras = [18, 25, 20, 18, 15, 12, 25, 15, 40, 12, 18, 30, 12];
  for (var i = 0; i < larguras.length; i++) {
    sheet.setColumnWidth(i + 1, larguras[i] * 7);
  }

  var regraTipo = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Aniversário', 'Casamento', 'Corporativo', 'Outro'], true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 4, 1000, 1).setDataValidation(regraTipo);

  var regraStatus = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendente', 'Aceito', 'Negado', 'Negociar'], true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 10, 1000, 1).setDataValidation(regraStatus);

  try { sheet.getRange(2, 1, 1000, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss'); } catch (e) {}
  try { sheet.getRange(2, 5, 1000, 1).setNumberFormat('dd/MM/yyyy'); } catch (e) {}
  try { sheet.getRange(2, 6, 1000, 1).setNumberFormat('#,##0'); } catch (e) {}
  try { sheet.getRange(2, 8, 1000, 1).setNumberFormat('"R$" #,##0.00'); } catch (e) {}
  try { sheet.getRange(2, 11, 1000, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss'); } catch (e) {}

  sheet.setFrozenRows(1);

  // --- ABA DASHBOARD ---
  var dash = ss.getSheetByName(DASH_NAME);
  if (dash) {
    var charts = dash.getCharts();
    for (var c = 0; c < charts.length; c++) {
      try { dash.removeChart(charts[c]); } catch (e) {}
    }
    ss.deleteSheet(dash);
  }
  dash = ss.insertSheet(DASH_NAME);

  // Larguras
  dash.setColumnWidth(1, 280);  // A
  dash.setColumnWidth(2, 150);  // B
  dash.setColumnWidth(3, 150);  // C
  dash.setColumnWidth(4, 150);  // D
  dash.setColumnWidth(5, 120);  // E
  dash.setColumnWidth(6, 80);   // F
  dash.setColumnWidth(7, 90);   // G
  dash.setColumnWidth(8, 90);   // H
  dash.setColumnWidth(9, 90);   // I

  // Título principal
  dash.getRange('A1:D1').merge();
  dash.getRange('A1').setValue('DASHBOARD — GESTÃO DE ORÇAMENTOS');
  dash.getRange('A1').setFontWeight('bold').setFontSize(14)
    .setBackground('#4285f4').setFontColor('#ffffff').setHorizontalAlignment('center');

  // === SEÇÕES COM LIMITE DE COLUNAS ===
  // Seção 1 — apenas A:B
  dash.getRange(3, 1, 1, 2).merge();
  dash.getRange(3, 1).setValue('SEÇÃO 1 — KPIs BÁSICOS');
  dash.getRange(3, 1, 1, 2).setBackground('#e3f2fd').setFontWeight('bold');

  // Seção 2 — apenas A:B
  dash.getRange(11, 1, 1, 2).merge();
  dash.getRange(11, 1).setValue('SEÇÃO 2 — FINANCEIRO');
  dash.getRange(11, 1, 1, 2).setBackground('#e8f5e9').setFontWeight('bold');

  // Seção 3 — apenas A:B
  dash.getRange(17, 1, 1, 2).merge();
  dash.getRange(17, 1).setValue('SEÇÃO 3 — MÉTRICAS DE EVENTOS');
  dash.getRange(17, 1, 1, 2).setBackground('#fff3e0').setFontWeight('bold');

  // Seção 4 — A:D
  dash.getRange(23, 1, 1, 4).merge();
  dash.getRange(23, 1).setValue('SEÇÃO 4 — POR TIPO DE EVENTO');
  dash.getRange(23, 1, 1, 4).setBackground('#f3e5f5').setFontWeight('bold');

  // Seção 5 — apenas A:B
  dash.getRange(30, 1, 1, 2).merge();
  dash.getRange(30, 1).setValue('SEÇÃO 5 — PROJEÇÕES E CENÁRIOS');
  dash.getRange(30, 1, 1, 2).setBackground('#e0f7fa').setFontWeight('bold');

  // Seção 6 — A:F (tem 6 colunas de dados)
  dash.getRange(40, 1, 1, 6).merge();
  dash.getRange(40, 1).setValue('SEÇÃO 6 — AGENDA DE PRÓXIMOS EVENTOS');
  dash.getRange(40, 1, 1, 6).setBackground('#fce4ec').setFontWeight('bold');

  // === DADOS DAS SEÇÕES ===
  // Seção 1 — KPIs (A4:B9)
  dash.getRange(4, 1, 6, 2).setValues([
    ['Total de Orçamentos', ''], ['Pendentes', ''], ['Aceitos', ''],
    ['Negados', ''], ['Em Negociação', ''], ['Taxa de Conversão', '']
  ]);

  // Seção 2 — Financeiro (A12:B15)
  dash.getRange(12, 1, 4, 2).setValues([
    ['Valor Faturado (Aceito)', ''], ['Valor em Aberto (Pendente)', ''],
    ['Valor em Negociação', ''], ['Ticket Médio', '']
  ]);

  // Seção 3 — Eventos (A18:B21)
  dash.getRange(18, 1, 4, 2).setValues([
    ['Total Eventos Confirmados', ''], ['Próximos Eventos (30 dias)', ''],
    ['Convidados Totais', ''], ['Média de Convidados', '']
  ]);

  // Seção 4 — Por Tipo (A24:D28)
  dash.getRange(24, 1, 1, 4).setValues([['Tipo', 'Quantidade', 'Faturamento', 'Taxa Conv.']]);
  dash.getRange(24, 1, 1, 4).setFontWeight('bold');
  dash.getRange(25, 1, 4, 4).setValues([
    ['Aniversário', '', '', ''], ['Casamento', '', '', ''],
    ['Corporativo', '', '', ''], ['Outro', '', '', '']
  ]);

  // Seção 5 — Projeções (A31:B38)
  dash.getRange(31, 1, 8, 2).setValues([
    ['Dias Decorridos', ''], ['Média Diária', ''], ['Projeção Mensal', ''],
    ['Faturamento Realista', ''], ['Faturamento Pessimista (-30%)', ''],
    ['Faturamento Otimista (+15%)', ''], ['Conversão Últimos 7 dias', ''],
    ['Tendência', '']
  ]);

  // Seção 6 — Agenda (A41:F41 header)
  dash.getRange(41, 1, 1, 6).setValues([
    ['Data', 'Tipo', 'Cliente', 'Convidados', 'Local', 'Dias até evento']
  ]);
  dash.getRange(41, 1, 1, 6).setFontWeight('bold');

  // === TÍTULOS DOS DADOS DE GRÁFICOS (colunas E+) ===
  // Pizza — título em E3
  dash.getRange(3, 5, 1, 2).merge();
  dash.getRange(3, 5).setValue('PIZZA — STATUS');
  dash.getRange(3, 5, 1, 2).setBackground('#fff8e1').setFontWeight('bold').setHorizontalAlignment('center');

  // Barras valores — título em E11
  dash.getRange(11, 5, 1, 5).merge();
  dash.getRange(11, 5).setValue('BARRAS — VALORES POR STATUS (R$)');
  dash.getRange(11, 5, 1, 5).setBackground('#fff8e1').setFontWeight('bold').setHorizontalAlignment('center');

  // Barras tipos — título em E17
  dash.getRange(17, 5, 1, 2).merge();
  dash.getRange(17, 5).setValue('BARRAS — QTDE POR TIPO');
  dash.getRange(17, 5, 1, 2).setBackground('#fff8e1').setFontWeight('bold').setHorizontalAlignment('center');

  // === DADOS PARA GRÁFICOS ===
  // Pizza: E6:F9 — ORDEM: Aceito, Pendente, Negado, Negociar
  dash.getRange(6, 5, 4, 2).setValues([
    ['Aceito', 0],
    ['Pendente', 0],
    ['Negado', 0],
    ['Negociar', 0]
  ]);

  // Barras valores: E12:I13 — cada status como série (cor individual)
  // Header row 12: F12="Aceito" G12="Pendente" H12="Negado" I12="Negociar"
  // Data row 13: E13="Valor" F13=val G13=val H13=0 I13=val
  dash.getRange(12, 6, 1, 4).setValues([['Aceito', 'Pendente', 'Negado', 'Negociar']]);
  dash.getRange(12, 6, 1, 4).setFontWeight('bold');
  dash.getRange(13, 5, 1, 5).setValues([['Valor (R$)', 0, 0, 0, 0]]);

  // Barras tipos: E18:F21
  dash.getRange(18, 5, 4, 2).setValues([
    ['Aniversário', 0], ['Casamento', 0], ['Corporativo', 0], ['Outro', 0]
  ]);

  // Criar gráficos
  criarGraficos(dash);

  // Atualizar com dados reais
  atualizarDashboard();

  SpreadsheetApp.getActiveSpreadsheet().toast('Setup concluído! Dashboard e gráficos criados.', '✅', 5);
  return 'Setup concluído com sucesso!';
}

// ============================================================
// CRIAR GRÁFICOS
// ============================================================
function criarGraficos(dash) {
  if (!dash) {
    Logger.log('criarGraficos: dash é undefined');
    return;
  }

  // 1. Pizza: E6:F9 — SEM cabeçalho
  //    Ordem: Aceito(0)=verde, Pendente(1)=laranja, Negado(2)=vermelho, Negociar(3)=roxo
  var pieRange = dash.getRange(6, 5, 4, 2);
  var pieChart = dash.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(pieRange)
    .setPosition(3, 7, 0, 0)
    .setOption('title', 'Distribuição de Status')
    .setOption('width', 350)
    .setOption('height', 250)
    .setOption('slices', {
      0: { color: '#34a853' }, // Aceito — Verde
      1: { color: '#ff9800' }, // Pendente — Laranja
      2: { color: '#ea4335' }, // Negado — Vermelho
      3: { color: '#9c27b0' }  // Negociar — Roxo
    })
    .setOption('pieSliceText', 'percentage')
    .setOption('legend', { position: 'right' })
    .build();
  dash.insertChart(pieChart);

  // 2. Barras: Valores por status (E12:I13) — cada status = série com cor própria
  //    Cores: Aceito=verde, Pendente=laranja, Negado=vermelho, Negociar=roxo
  var barRange1 = dash.getRange(12, 6, 2, 4); // F12:I13 (4 séries)
  var barChart1 = dash.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(barRange1)
    .setPosition(17, 7, 0, 0)
    .setOption('title', 'Valores por Status (R$)')
    .setOption('width', 350)
    .setOption('height', 250)
    .setOption('legend', { position: 'top' })
    .setOption('colors', ['#34a853', '#ff9800', '#ea4335', '#9c27b0'])
    .setOption('isStacked', 'false')
    .build();
  dash.insertChart(barChart1);

  // 3. Barras: Quantidade por tipo (E18:F21)
  var barRange2 = dash.getRange(18, 5, 4, 2);
  var barChart2 = dash.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(barRange2)
    .setPosition(31, 7, 0, 0)
    .setOption('title', 'Quantidade por Tipo de Evento')
    .setOption('width', 350)
    .setOption('height', 250)
    .setOption('legend', { position: 'none' })
    .setOption('series', { 0: { color: '#34a853' } })
    .build();
  dash.insertChart(barChart2);
}
