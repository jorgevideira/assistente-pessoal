/********************************************************************
 *  ASSISTENTE PESSOAL — Telegram + Google Agenda + Planilha de Gastos
 *  + Agenda do Outlook (trabalho) sincronizada via link ICS
 *  IA: Google Gemini (camada gratuita) | Hospedagem: Google Apps Script
 *
 *  Versão 2
 *
 *  Instalação do zero:
 *   1. Preencha TELEGRAM_TOKEN e GEMINI_API_KEY abaixo e salve.
 *   2. Execute  configurar   (autorize o acesso).
 *   3. Implantar > Nova implantação > App da Web
 *        Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 *   4. Cole a URL (termina em /exec) em WEB_APP_URL e salve.
 *   5. Execute  registrarWebhook   e depois   criarGatilhos
 *   6. No Telegram, mande /start para o seu bot.
 *   7. (Opcional) Outlook: cole o link ICS em OUTLOOK_ICS_URL e
 *      execute  configurarOutlook
 *
 *  Ao alterar o código depois: Implantar > Gerenciar implantações >
 *  lápis > Versão: "Nova versão" > Implantar (a URL continua a mesma).
 ********************************************************************/

// ====================== 1. PREENCHA AQUI ======================
const CONFIG_INICIAL = {
  TELEGRAM_TOKEN:  'COLE_AQUI_O_TOKEN_DO_BOTFATHER',
  GEMINI_API_KEY:  'COLE_AQUI_A_CHAVE_DO_GEMINI',
  WEB_APP_URL:     'COLE_AQUI_A_URL_DO_APP_DA_WEB',
  OUTLOOK_ICS_URL: 'COLE_AQUI_O_LINK_ICS_DO_OUTLOOK', // opcional
};

// ====================== 2. AJUSTES OPCIONAIS ======================
const MODELO_GEMINI    = 'gemini-3.1-flash-lite';
const MODELOS_RESERVA  = ['gemini-flash-lite-latest', 'gemini-3.5-flash'];
const FUSO             = 'America/Sao_Paulo';
const OFFSET           = '-03:00';
const HORA_RESUMO      = 7;    // resumo diário às 7h
// Avisos antes de cada compromisso, em minutos (máx. 360). Ex.: [60, 15] = avisa 1h antes e 15 min antes.
const LEMBRETES_MINUTOS = [15];
const NOME_AGENDA_TRABALHO = 'Trabalho (Outlook)';
// Outras agendas do Google para o bot consultar (somente leitura), pelo nome exato que aparece no Google Agenda.
const AGENDAS_EXTRAS   = []; // ex.: ['Família', 'Clube do Livro']
const SYNC_DIAS        = 60;   // sincroniza do dia anterior até 60 dias à frente
const CATEGORIAS = ['Alimentação', 'Mercado', 'Transporte', 'Moradia', 'Contas',
                    'Saúde', 'Lazer', 'Educação', 'Compras', 'Outros'];

const DIA_MS = 86400000;
const DIAS_SEMANA = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira',
                     'quinta-feira', 'sexta-feira', 'sábado'];

const AJUDA =
  'Oi! Eu cuido da sua agenda e dos seus gastos. Exemplos:\n\n' +
  '💰 "gastei 32,50 no almoço"\n' +
  '💰 "quanto gastei com mercado esse mês?"\n' +
  '📷 mande a foto de um cupom ou comprovante\n' +
  '📅 "marca dentista sexta às 14h"\n' +
  '📅 "o que tenho amanhã?"\n' +
  '📅 "muda a reunião de amanhã para 16h"\n' +
  '🎤 ou mande tudo isso por áudio\n\n' +
  'Compromissos do trabalho (Outlook) aparecem com 💼 e são só para consulta.\n\n' +
  'Comandos: /resumo  /desfazer (apaga o último gasto)  /sincronizar (atualiza o Outlook agora)  /limpar (zera a memória da conversa)  /ajuda';


// =================================================================
//  FUNÇÕES DE INSTALAÇÃO (execute pelo editor)
// =================================================================

function configurar() {
  const c = CONFIG_INICIAL;
  if (c.TELEGRAM_TOKEN.indexOf('COLE') === 0) throw new Error('Preencha TELEGRAM_TOKEN no topo do código.');
  if (c.GEMINI_API_KEY.indexOf('COLE') === 0) throw new Error('Preencha GEMINI_API_KEY no topo do código.');

  const p = props_();
  p.setProperty('TELEGRAM_TOKEN', c.TELEGRAM_TOKEN.trim());
  p.setProperty('GEMINI_API_KEY', c.GEMINI_API_KEY.trim());
  if (!p.getProperty('SEGREDO')) p.setProperty('SEGREDO', Utilities.getUuid().replace(/-/g, ''));

  if (!p.getProperty('PLANILHA_ID')) {
    const ss = SpreadsheetApp.create('Assistente - Gastos');
    ss.setSpreadsheetTimeZone(FUSO);
    const sh = ss.getSheets()[0];
    sh.setName('Gastos');
    sh.appendRow(['Data', 'Valor', 'Categoria', 'Descrição', 'Registrado em']);
    sh.setFrozenRows(1);
    sh.getRange('1:1').setFontWeight('bold');
    sh.getRange('A2:A').setNumberFormat('dd/mm/yyyy');
    sh.getRange('B2:B').setNumberFormat('"R$" #,##0.00');
    sh.getRange('E2:E').setNumberFormat('dd/mm/yyyy hh:mm');
    p.setProperty('PLANILHA_ID', ss.getId());
    Logger.log('📊 Planilha criada: ' + ss.getUrl());
  }

  Logger.log('📅 Agenda: ' + CalendarApp.getDefaultCalendar().getName());
  const me = tg_('getMe');
  if (!me.ok) throw new Error('Token do Telegram inválido: ' + JSON.stringify(me));
  Logger.log('🤖 Bot: @' + me.result.username);
  Logger.log('✅ Configuração salva. Próximo passo: implantar como App da Web.');
}

function registrarWebhook() {
  const url = CONFIG_INICIAL.WEB_APP_URL.trim();
  if (!url || url.indexOf('COLE') === 0) throw new Error('Preencha WEB_APP_URL com a URL da implantação.');
  const r = tg_('setWebhook', {
    url: url + '?token=' + prop_('SEGREDO'),
    allowed_updates: ['message'],
    drop_pending_updates: true,
  });
  Logger.log(JSON.stringify(r));
  removerGatilho_('buscarMensagens');
  if (r.ok) Logger.log('✅ Webhook registrado.');
}

function criarGatilhos() {
  removerGatilho_('resumoDiario');
  removerGatilho_('verificarLembretes');
  ScriptApp.newTrigger('resumoDiario').timeBased()
    .atHour(HORA_RESUMO).everyDays(1).inTimezone(FUSO).create();
  ScriptApp.newTrigger('verificarLembretes').timeBased().everyMinutes(5).create();
  Logger.log('✅ Gatilhos criados: resumo às ' + HORA_RESUMO + 'h e lembretes a cada 5 min.');
}

/** Liga a sincronização Outlook → Google (somente leitura). */
function configurarOutlook() {
  const url = String(CONFIG_INICIAL.OUTLOOK_ICS_URL || '').trim().replace(/^webcal:\/\//i, 'https://');
  if (!url || url.indexOf('COLE') === 0) throw new Error('Preencha OUTLOOK_ICS_URL no topo do código.');
  const p = props_();
  p.setProperty('OUTLOOK_ICS_URL', url);

  if (!calendarioTrabalho_()) {
    const cal = CalendarApp.createCalendar(NOME_AGENDA_TRABALHO, { timeZone: FUSO, color: CalendarApp.Color.BLUE });
    p.setProperty('CAL_TRABALHO_ID', cal.getId());
    Logger.log('📅 Agenda criada no Google: ' + NOME_AGENDA_TRABALHO);
  }

  CacheService.getScriptCache().remove('sync_outlook');
  Logger.log(sincronizarOutlook());

  removerGatilho_('sincronizarOutlook');
  ScriptApp.newTrigger('sincronizarOutlook').timeBased().everyMinutes(15).create();
  Logger.log('✅ Sincronização automática ativa (a cada 15 min).');
}

/** Desliga a sincronização (a agenda "Trabalho (Outlook)" continua no Google). */
function desativarOutlook() {
  removerGatilho_('sincronizarOutlook');
  props_().deleteProperty('OUTLOOK_ICS_URL');
  Logger.log('Sincronização desligada. Se quiser, apague a agenda "' + NOME_AGENDA_TRABALHO + '" no Google Agenda.');
}

/** Plano B: se o bot responder em dobro ou atrasar, use este modo (responde em até ~1 min). */
function ativarModoPolling() {
  tg_('deleteWebhook', { drop_pending_updates: false });
  removerGatilho_('buscarMensagens');
  ScriptApp.newTrigger('buscarMensagens').timeBased().everyMinutes(1).create();
  Logger.log('✅ Modo polling ativo. Para voltar ao webhook, execute registrarWebhook.');
}

function diagnosticar() {
  Logger.log(JSON.stringify(tg_('getWebhookInfo'), null, 2));
  Logger.log('Dono registrado: ' + (prop_('DONO_CHAT_ID') || 'nenhum ainda'));
  Logger.log('Outlook: ' + (prop_('OUTLOOK_ICS_URL') ? 'configurado' : 'não configurado'));
}

/** Lista todas as agendas que o script enxerga e quantos eventos cada uma tem nos próximos 30 dias. */
function listarAgendas() {
  const hoje = new Date(), fim = new Date(hoje.getTime() + 30 * DIA_MS);
  CalendarApp.getAllCalendars().forEach(function (c) {
    Logger.log('"' + c.getName() + '" → ' + c.getEvents(hoje, fim).length + ' eventos nos próximos 30 dias');
  });
  Logger.log('Extras encontradas pelo script: ' + calendariosExtras_().map(function (c) { return c.nome; }).join(', '));
}

function testarGemini() {
  const modelos = [MODELO_GEMINI].concat(MODELOS_RESERVA, ['gemini-flash-latest']);
  modelos.forEach(function (m) {
    const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': prop_('GEMINI_API_KEY') },
      payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: 'diga oi' }] }] }),
    });
    const cod = r.getResponseCode();
    let info = '';
    try { const j = JSON.parse(r.getContentText()); info = cod === 200 ? '✅ FUNCIONA' : j.error.message; } catch (e) {}
    Logger.log(m + ' → ' + cod + ' ' + String(info).slice(0, 150));
  });
}


// =================================================================
//  ENTRADA DE MENSAGENS
// =================================================================

function doPost(e) {
  try {
    if (!e || !e.parameter || e.parameter.token !== prop_('SEGREDO')) return ok_();
    const update = JSON.parse(e.postData.contents);
    if (jaProcessado_(update.update_id)) return ok_();
    if (update.message) tratarMensagem_(update.message);
  } catch (err) {
    avisarErro_(err);
  }
  return ok_();
}

function doGet() {
  return HtmlService.createHtmlOutput('Assistente no ar ✅');
}

function buscarMensagens() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const offset = Number(prop_('ULTIMO_UPDATE') || 0) + 1;
    const r = tg_('getUpdates', { offset: offset, timeout: 0, allowed_updates: ['message'] });
    if (!r.ok) return;
    r.result.forEach(function (u) {
      props_().setProperty('ULTIMO_UPDATE', String(u.update_id));
      if (!u.message) return;
      try { tratarMensagem_(u.message); } catch (err) { avisarErro_(err); }
    });
  } finally {
    lock.releaseLock();
  }
}

function tratarMensagem_(msg) {
  const chatId = String(msg.chat.id);
  const dono = prop_('DONO_CHAT_ID');

  if (!dono) {
    props_().setProperty('DONO_CHAT_ID', chatId);
    enviar_(chatId, '✅ Pronto! Este chat foi registrado como dono do assistente. Só você poderá usá-lo.\n\n' + AJUDA);
    return;
  }
  if (chatId !== dono) return;

  const texto = String(msg.text || msg.caption || '').trim();
  const comando = texto.split(/[\s@]/)[0].toLowerCase();
  if (comando === '/start' || comando === '/ajuda') return enviar_(chatId, AJUDA);
  if (comando === '/resumo') return enviar_(chatId, montarResumo_('📋 Resumo'));
  if (comando === '/desfazer') return enviar_(chatId, desfazerUltimoGasto_().mensagem);
  if (comando === '/sincronizar') {
    if (!prop_('OUTLOOK_ICS_URL')) return enviar_(chatId, 'O Outlook ainda não foi configurado.');
    CacheService.getScriptCache().remove('sync_outlook');
    return enviar_(chatId, '🔄 ' + sincronizarOutlook());
  }
  if (comando === '/limpar') {
    CacheService.getScriptCache().remove('historico');
    return enviar_(chatId, '🧹 Memória da conversa apagada.');
  }

  const partes = [];
  const rotulo = [];

  const audio = msg.voice || msg.audio;
  if (audio) {
    partes.push(baixarArquivo_(audio.file_id, audio.mime_type || 'audio/ogg'));
    partes.push({ text: '(O usuário mandou este áudio. Entenda o pedido falado e atenda.)' });
    rotulo.push('[mensagem de voz]');
  }
  if (msg.photo && msg.photo.length) {
    const foto = msg.photo[msg.photo.length - 1];
    partes.push(baixarArquivo_(foto.file_id, 'image/jpeg'));
    if (!texto) partes.push({ text: '(O usuário mandou esta imagem, provavelmente um cupom ou comprovante. Extraia valor total, estabelecimento e data e registre o gasto.)' });
    rotulo.push('[foto]');
  }
  if (msg.document && /^(image\/|application\/pdf)/.test(msg.document.mime_type || '')) {
    partes.push(baixarArquivo_(msg.document.file_id, msg.document.mime_type));
    if (!texto) partes.push({ text: '(O usuário mandou este arquivo, provavelmente um comprovante. Extraia valor total, estabelecimento e data e registre o gasto.)' });
    rotulo.push('[arquivo]');
  }
  if (texto) { partes.push({ text: texto }); rotulo.push(texto); }

  if (!partes.length) return enviar_(chatId, 'Ainda não entendo esse tipo de mensagem. Mande texto, áudio ou foto 🙂');

  tg_('sendChatAction', { chat_id: chatId, action: 'typing' });
  enviar_(chatId, conversar_(partes, rotulo.join(' ')));
}


// =================================================================
//  IA (GEMINI) COM FERRAMENTAS
// =================================================================

const FERRAMENTAS = [
  {
    name: 'registrar_gasto',
    description: 'Registra um gasto na planilha.',
    parameters: {
      type: 'OBJECT',
      properties: {
        valor:     { type: 'NUMBER', description: 'Valor em reais, ex.: 32.5' },
        categoria: { type: 'STRING', enum: CATEGORIAS },
        descricao: { type: 'STRING', description: 'Descrição curta, ex.: "Almoço no restaurante X"' },
        data:      { type: 'STRING', description: 'YYYY-MM-DD. Omitir se for hoje.' },
      },
      required: ['valor', 'categoria', 'descricao'],
    },
  },
  {
    name: 'consultar_gastos',
    description: 'Soma e lista gastos de um período (datas inclusivas), opcionalmente de uma categoria.',
    parameters: {
      type: 'OBJECT',
      properties: {
        data_inicio: { type: 'STRING', description: 'YYYY-MM-DD' },
        data_fim:    { type: 'STRING', description: 'YYYY-MM-DD' },
        categoria:   { type: 'STRING', enum: CATEGORIAS },
      },
      required: ['data_inicio', 'data_fim'],
    },
  },
  {
    name: 'desfazer_ultimo_gasto',
    description: 'Apaga o último gasto registrado na planilha.',
  },
  {
    name: 'criar_evento',
    description: 'Cria um compromisso na agenda pessoal.',
    parameters: {
      type: 'OBJECT',
      properties: {
        titulo:    { type: 'STRING' },
        inicio:    { type: 'STRING', description: 'YYYY-MM-DDTHH:MM:SS, ou YYYY-MM-DD para evento de dia inteiro' },
        fim:       { type: 'STRING', description: 'YYYY-MM-DDTHH:MM:SS. Omitir para duração de 1 hora.' },
        local:     { type: 'STRING' },
        descricao: { type: 'STRING' },
      },
      required: ['titulo', 'inicio'],
    },
  },
  {
    name: 'listar_eventos',
    description: 'Lista compromissos pessoais e do trabalho entre duas datas (inclusivas). Retorna o id e a origem de cada evento.',
    parameters: {
      type: 'OBJECT',
      properties: {
        data_inicio: { type: 'STRING', description: 'YYYY-MM-DD' },
        data_fim:    { type: 'STRING', description: 'YYYY-MM-DD' },
      },
      required: ['data_inicio', 'data_fim'],
    },
  },
  {
    name: 'editar_evento',
    description: 'Altera título e/ou horário de um evento pessoal. Use listar_eventos antes para obter o id.',
    parameters: {
      type: 'OBJECT',
      properties: {
        id:          { type: 'STRING' },
        novo_titulo: { type: 'STRING' },
        novo_inicio: { type: 'STRING', description: 'YYYY-MM-DDTHH:MM:SS' },
        novo_fim:    { type: 'STRING', description: 'YYYY-MM-DDTHH:MM:SS. Omitir para manter a duração.' },
      },
      required: ['id'],
    },
  },
  {
    name: 'cancelar_evento',
    description: 'Exclui um evento pessoal da agenda. Use listar_eventos antes para obter o id.',
    parameters: {
      type: 'OBJECT',
      properties: { id: { type: 'STRING' } },
      required: ['id'],
    },
  },
];

function executarFerramenta_(nome, a) {
  switch (nome) {
    case 'registrar_gasto':       return registrarGasto_(a);
    case 'consultar_gastos':      return consultarGastos_(a);
    case 'desfazer_ultimo_gasto': return desfazerUltimoGasto_();
    case 'criar_evento':          return criarEvento_(a);
    case 'listar_eventos':        return listarEventos_(a);
    case 'editar_evento':         return editarEvento_(a);
    case 'cancelar_evento':       return cancelarEvento_(a);
    default:                      return { erro: 'Ferramenta desconhecida: ' + nome };
  }
}

function instrucoes_() {
  const agora = new Date();
  return [
    'Você é um assistente pessoal que conversa pelo Telegram, em português do Brasil.',
    'Agora é ' + diaSemana_(agora) + ', ' + Utilities.formatDate(agora, FUSO, 'dd/MM/yyyy HH:mm') +
      ' (horário de Brasília). Data de hoje em ISO: ' + Utilities.formatDate(agora, FUSO, 'yyyy-MM-dd') + '.',
    'Você gerencia a agenda (Google Agenda) e os gastos (planilha) do usuário por meio das ferramentas.',
    '',
    'Regras:',
    '- Sempre use as ferramentas para registrar, alterar ou consultar. Nunca invente dados nem confirme algo que não foi executado.',
    '- Interprete datas relativas (amanhã, sexta, semana que vem, mês passado) a partir da data de hoje.',
    '- Gastos: valor numérico em reais. Escolha a categoria mais adequada entre: ' + CATEGORIAS.join(', ') + '. Sem data informada = hoje.',
    '- Se a mensagem tiver vários gastos, registre cada um separadamente.',
    '- Compromissos sem duração informada duram 1 hora. Se faltar o horário, pergunte antes de criar.',
    '- Novos compromissos sempre vão para a agenda pessoal.',
    '- Eventos com origem "Trabalho (Outlook)" são somente leitura: não tente editar nem cancelar; diga que a alteração deve ser feita no Outlook. Ao listar, marque-os com 💼.',
    '- Eventos de outras agendas (' + AGENDAS_EXTRAS.join(', ') + ') também são somente leitura; ao listar, mostre o nome da agenda.',
    '- Para editar ou cancelar, liste os eventos do período primeiro. Se houver mais de um candidato, pergunte qual é.',
    '- Em cupons e comprovantes, use o valor TOTAL pago.',
    '- Responda de forma curta e direta, em texto simples (sem markdown, sem asteriscos), com emojis com moderação.',
    '- Valores no formato R$ 1.234,56 e datas como dd/mm.',
    '- Se o pedido não tiver relação com agenda ou gastos, ajude normalmente, de forma breve.',
  ].join('\n');
}

function conversar_(partesUsuario, rotulo) {
  const historico = lerHistorico_();
  const contents = historico.concat([{ role: 'user', parts: partesUsuario }]);

  for (let i = 0; i < 6; i++) {
    const r = gemini_(contents);
    const cand = r.candidates && r.candidates[0];
    if (!cand || !cand.content || !cand.content.parts) {
      return 'Não consegui processar isso agora. Pode tentar de novo com outras palavras?';
    }
    contents.push(cand.content);

    const chamadas = cand.content.parts.filter(function (p) { return p.functionCall; });
    if (!chamadas.length) {
      const resposta = cand.content.parts
        .filter(function (p) { return p.text && !p.thought; })
        .map(function (p) { return p.text; }).join('').trim() || 'Feito ✅';
      salvarHistorico_(historico, rotulo, resposta);
      return resposta;
    }

    const respostas = chamadas.map(function (c) {
      let resultado;
      try {
        resultado = executarFerramenta_(c.functionCall.name, c.functionCall.args || {});
      } catch (err) {
        resultado = { erro: err.message };
      }
      const fr = { name: c.functionCall.name, response: resultado };
      if (c.functionCall.id) fr.id = c.functionCall.id;
      return { functionResponse: fr };
    });
    contents.push({ role: 'user', parts: respostas });
  }
  return 'O pedido ficou complexo demais para mim. Pode dividir em partes?';
}

function gemini_(contents) {
  const corpo = {
    systemInstruction: { parts: [{ text: instrucoes_() }] },
    contents: contents,
    tools: [{ functionDeclarations: FERRAMENTAS }],
  };
  const modelos = [MODELO_GEMINI].concat(MODELOS_RESERVA);
  let ultimoErro = '';
  for (let m = 0; m < modelos.length; m++) {
    for (let tentativa = 0; tentativa < 2; tentativa++) {
      const resp = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + modelos[m] + ':generateContent', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-goog-api-key': prop_('GEMINI_API_KEY') },
        payload: JSON.stringify(corpo),
      });
      const codigo = resp.getResponseCode();
      if (codigo === 200) return JSON.parse(resp.getContentText());
      ultimoErro = modelos[m] + ' ' + codigo + ': ' + resp.getContentText().slice(0, 150);
      if (codigo === 429 || codigo >= 500) { Utilities.sleep(2000); continue; }
      break;
    }
  }
  throw new Error('A IA não respondeu agora (' + ultimoErro + '). Tente de novo em instantes.');
}

function lerHistorico_() {
  const h = CacheService.getScriptCache().get('historico');
  return h ? JSON.parse(h) : [];
}

function salvarHistorico_(historico, rotulo, resposta) {
  const novo = historico.concat([
    { role: 'user',  parts: [{ text: rotulo || '(mensagem)' }] },
    { role: 'model', parts: [{ text: resposta }] },
  ]).slice(-10);
  try {
    CacheService.getScriptCache().put('historico', JSON.stringify(novo), 1800);
  } catch (e) { /* ignora */ }
}


// =================================================================
//  GASTOS (Google Sheets)
// =================================================================

function planilha_() {
  return SpreadsheetApp.openById(prop_('PLANILHA_ID')).getSheetByName('Gastos');
}

function lerGastos_() {
  const sh = planilha_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .filter(function (l) { return l[0] instanceof Date && typeof l[1] === 'number'; })
    .map(function (l) { return { data: l[0], valor: l[1], categoria: String(l[2] || 'Outros'), descricao: String(l[3] || '') }; });
}

function registrarGasto_(a) {
  const valor = Number(String(a.valor).replace(',', '.'));
  if (!isFinite(valor) || valor <= 0) return { erro: 'Valor inválido.' };
  const data = a.data ? parseData_(a.data) : hojeInicio_();
  planilha_().appendRow([data, valor, a.categoria || 'Outros', a.descricao || '', new Date()]);
  return { ok: true, data: fmtData_(data), valor: moeda_(valor), categoria: a.categoria, descricao: a.descricao };
}

function consultarGastos_(a) {
  const ini = parseData_(a.data_inicio);
  const fim = new Date(parseData_(a.data_fim).getTime() + DIA_MS);
  const cat = a.categoria ? String(a.categoria).toLowerCase() : null;
  const itens = lerGastos_().filter(function (g) {
    return g.data >= ini && g.data < fim && (!cat || g.categoria.toLowerCase() === cat);
  });

  let total = 0;
  const porCategoria = {};
  itens.forEach(function (g) {
    total += g.valor;
    porCategoria[g.categoria] = (porCategoria[g.categoria] || 0) + g.valor;
  });
  Object.keys(porCategoria).forEach(function (k) { porCategoria[k] = moeda_(porCategoria[k]); });

  return {
    periodo: fmtData_(ini) + ' a ' + fmtData_(new Date(fim.getTime() - DIA_MS)),
    total: moeda_(total),
    quantidade: itens.length,
    por_categoria: porCategoria,
    lancamentos: itens.slice(-30).map(function (g) {
      return fmtData_(g.data) + ' | ' + moeda_(g.valor) + ' | ' + g.categoria + ' | ' + g.descricao;
    }),
    observacao: itens.length > 30 ? 'Mostrando só os 30 lançamentos mais recentes.' : undefined,
  };
}

function desfazerUltimoGasto_() {
  const sh = planilha_();
  const n = sh.getLastRow();
  if (n < 2) return { ok: false, mensagem: 'Não há gastos para apagar.' };
  const l = sh.getRange(n, 1, 1, 4).getValues()[0];
  sh.deleteRow(n);
  const desc = moeda_(Number(l[1])) + ' — ' + l[3] + ' (' + l[2] + ')';
  return { ok: true, removido: desc, mensagem: '↩️ Apaguei o último gasto: ' + desc };
}


// =================================================================
//  AGENDA (Google Calendar) — pessoal + trabalho
// =================================================================

function calendario_() { return CalendarApp.getDefaultCalendar(); }

function calendarioTrabalho_() {
  const id = prop_('CAL_TRABALHO_ID');
  return id ? CalendarApp.getCalendarById(id) : null;
}

/** Agendas extras (somente leitura) encontradas pelo nome. */
function calendariosExtras_() {
  const lista = [];
  AGENDAS_EXTRAS.forEach(function (nome) {
    const cals = CalendarApp.getCalendarsByName(nome);
    if (cals && cals.length) lista.push({ cal: cals[0], nome: nome });
    else console.warn('Agenda extra não encontrada: ' + nome);
  });
  return lista;
}

/** Eventos de todas as agendas no intervalo, em ordem de horário. */
function eventosEntre_(ini, fim) {
  const lista = [];
  calendario_().getEvents(ini, fim).forEach(function (ev) { lista.push({ ev: ev, trabalho: false }); });
  const t = calendarioTrabalho_();
  if (t) t.getEvents(ini, fim).forEach(function (ev) { lista.push({ ev: ev, trabalho: true }); });
  calendariosExtras_().forEach(function (c) {
    c.cal.getEvents(ini, fim).forEach(function (ev) { lista.push({ ev: ev, trabalho: false, extra: c.nome }); });
  });
  lista.sort(function (a, b) { return a.ev.getStartTime().getTime() - b.ev.getStartTime().getTime(); });
  return lista;
}

/** Marcador mostrado antes do título no resumo e nos lembretes. */
function marca_(x) {
  if (x.trabalho) return '💼 ';
  if (x.extra) return '[' + x.extra + '] ';
  return '';
}

function descreverEvento_(ev, trabalho, extra) {
  const diaTodo = ev.isAllDayEvent();
  return {
    id: ev.getId(),
    origem: trabalho ? 'Trabalho (Outlook) – somente leitura' : extra ? extra + ' – somente leitura' : 'Pessoal',
    titulo: ev.getTitle(),
    inicio: diaTodo ? fmtData_(ev.getAllDayStartDate()) + ' (dia todo)' : fmtDataHora_(ev.getStartTime()),
    fim: diaTodo ? undefined : fmtDataHora_(ev.getEndTime()),
    local: ev.getLocation() || undefined,
  };
}

function criarEvento_(a) {
  const cal = calendario_();
  const opcoes = {};
  if (a.local) opcoes.location = a.local;
  if (a.descricao) opcoes.description = a.descricao;

  let ev;
  if (/^\d{4}-\d{2}-\d{2}$/.test(a.inicio)) {
    ev = cal.createAllDayEvent(a.titulo, parseData_(a.inicio), opcoes);
  } else {
    const ini = parseDataHora_(a.inicio);
    const fim = a.fim ? parseDataHora_(a.fim) : new Date(ini.getTime() + 3600000);
    ev = cal.createEvent(a.titulo, ini, fim, opcoes);
  }
  return { ok: true, evento: descreverEvento_(ev, false) };
}

function listarEventos_(a) {
  const ini = parseData_(a.data_inicio);
  const fim = new Date(parseData_(a.data_fim).getTime() + DIA_MS);
  const eventos = eventosEntre_(ini, fim).slice(0, 80).map(function (x) { return descreverEvento_(x.ev, x.trabalho, x.extra); });
  return { quantidade: eventos.length, eventos: eventos };
}

function eventoPessoal_(id) {
  const ev = calendario_().getEventById(id);
  if (ev) return { ev: ev };
  const t = calendarioTrabalho_();
  if (t && t.getEventById(id)) return { erro: 'Esse compromisso vem do Outlook do trabalho e é somente leitura. Altere pelo Outlook.' };
  const extra = calendariosExtras_().filter(function (c) { return c.cal.getEventById(id); })[0];
  if (extra) return { erro: 'Esse compromisso é da agenda "' + extra.nome + '", que é somente leitura para mim.' };
  return { erro: 'Evento não encontrado.' };
}

function editarEvento_(a) {
  const r = eventoPessoal_(a.id);
  if (r.erro) return r;
  const ev = r.ev;
  if (a.novo_titulo) ev.setTitle(a.novo_titulo);
  if (a.novo_inicio) {
    const ini = parseDataHora_(a.novo_inicio);
    const duracao = ev.getEndTime().getTime() - ev.getStartTime().getTime();
    const fim = a.novo_fim ? parseDataHora_(a.novo_fim) : new Date(ini.getTime() + duracao);
    ev.setTime(ini, fim);
  }
  return { ok: true, evento: descreverEvento_(ev, false) };
}

function cancelarEvento_(a) {
  const r = eventoPessoal_(a.id);
  if (r.erro) return r;
  const info = descreverEvento_(r.ev, false);
  r.ev.deleteEvent();
  return { ok: true, cancelado: info };
}


// =================================================================
//  OUTLOOK (link ICS) → GOOGLE AGENDA  —  sincronização em mão única
// =================================================================
//  • Lê o ICS publicado pelo Outlook e espelha os compromissos na agenda
//    "Trabalho (Outlook)" do Google: cria o que é novo, atualiza o que
//    mudou e apaga o que foi removido/cancelado no Outlook.
//  • Só mexe em eventos que ele mesmo criou nessa agenda.
//  • Nada é enviado de volta ao Outlook.

const TZ_WINDOWS = {
  'E. South America Standard Time': 'America/Sao_Paulo',
  'Central Brazilian Standard Time': 'America/Cuiaba',
  'SA Western Standard Time': 'America/Manaus',
  'SA Eastern Standard Time': 'America/Fortaleza',
  'Bahia Standard Time': 'America/Bahia',
  'Tocantins Standard Time': 'America/Araguaina',
  'Argentina Standard Time': 'America/Argentina/Buenos_Aires',
  'SA Pacific Standard Time': 'America/Bogota',
  'Pacific SA Standard Time': 'America/Santiago',
  'Montevideo Standard Time': 'America/Montevideo',
  'Paraguay Standard Time': 'America/Asuncion',
  'Eastern Standard Time': 'America/New_York',
  'Central Standard Time': 'America/Chicago',
  'Mountain Standard Time': 'America/Denver',
  'Pacific Standard Time': 'America/Los_Angeles',
  'Central Standard Time (Mexico)': 'America/Mexico_City',
  'GMT Standard Time': 'Europe/London',
  'Greenwich Standard Time': 'Atlantic/Reykjavik',
  'W. Europe Standard Time': 'Europe/Berlin',
  'Romance Standard Time': 'Europe/Paris',
  'Central Europe Standard Time': 'Europe/Budapest',
  'Central European Standard Time': 'Europe/Warsaw',
  'GTB Standard Time': 'Europe/Bucharest',
  'FLE Standard Time': 'Europe/Kiev',
  'Russian Standard Time': 'Europe/Moscow',
  'India Standard Time': 'Asia/Kolkata',
  'China Standard Time': 'Asia/Shanghai',
  'Singapore Standard Time': 'Asia/Singapore',
  'Tokyo Standard Time': 'Asia/Tokyo',
  'AUS Eastern Standard Time': 'Australia/Sydney',
  'UTC': 'UTC',
  'Coordinated Universal Time': 'UTC',
  'tzone://Microsoft/Utc': 'UTC',
};

let TZ_FIXOS_ = {};       // offsets lidos dos blocos VTIMEZONE (plano B)
const CACHE_OFFSET_ = {};

function sincronizarOutlook() {
  const url = prop_('OUTLOOK_ICS_URL');
  const cal = calendarioTrabalho_();
  if (!url || !cal) return 'Outlook não configurado.';

  const cache = CacheService.getScriptCache();
  if (cache.get('sync_outlook')) return 'Já existe uma sincronização em andamento.';
  cache.put('sync_outlook', '1', 360);

  const inicioExec = Date.now();
  const LIMITE_MS = 4.5 * 60 * 1000;
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const texto = resp.getContentText();
    if (resp.getResponseCode() !== 200 || texto.indexOf('BEGIN:VCALENDAR') < 0) {
      throw new Error('Não consegui ler o link ICS do Outlook (HTTP ' + resp.getResponseCode() + '). Confira se o calendário continua publicado.');
    }

    const hoje = hojeInicio_();
    const janIni = new Date(hoje.getTime() - DIA_MS);
    const janFim = new Date(hoje.getTime() + SYNC_DIAS * DIA_MS);
    const desejados = {};
    lerIcs_(texto, janIni.getTime(), janFim.getTime()).forEach(function (o) { desejados[o.chave] = o; });

    let criados = 0, removidos = 0, mantidos = 0, pendentes = false;
    cal.getEvents(janIni, janFim).forEach(function (ev) {
      const chave = ev.getTag('outlook_chave');
      if (!chave) return; // evento criado à mão nessa agenda: não mexe
      const o = desejados[chave];
      if (o && ev.getTag('outlook_hash') === o.hash) { delete desejados[chave]; mantidos++; return; }
      if (Date.now() - inicioExec > LIMITE_MS) { pendentes = true; return; }
      ev.deleteEvent(); removidos++; pausa_(); // se mudou, é recriado abaixo
    });

    Object.keys(desejados).forEach(function (k) {
      if (Date.now() - inicioExec > LIMITE_MS) { pendentes = true; return; }
      criarEventoOutlook_(cal, desejados[k]); criados++; pausa_();
    });

    cache.remove('erro_sync');
    return 'Outlook sincronizado: ' + criados + ' criado(s), ' + removidos + ' removido(s)/atualizado(s), ' +
      mantidos + ' sem mudança.' + (pendentes ? ' Faltou terminar; continua na próxima rodada.' : '');
  } catch (err) {
    console.error(err);
    if (!cache.get('erro_sync')) {
      cache.put('erro_sync', '1', 21600); // avisa no máximo a cada 6h
      const dono = prop_('DONO_CHAT_ID');
      if (dono) enviar_(dono, '⚠️ Sincronização do Outlook: ' + err.message);
    }
    return 'Erro: ' + err.message;
  } finally {
    cache.remove('sync_outlook');
  }
}

function criarEventoOutlook_(cal, o) {
  const opc = { description: 'Sincronizado do Outlook. Alterações feitas aqui serão desfeitas.' };
  if (o.local) opc.location = o.local;
  let ev;
  if (o.diaTodo) {
    const ini = dataDeNaive_(o.inicioNaive);
    ev = (o.fimNaive - o.inicioNaive > DIA_MS)
      ? cal.createAllDayEvent(o.titulo, ini, dataDeNaive_(o.fimNaive), opc)
      : cal.createAllDayEvent(o.titulo, ini, opc);
  } else {
    ev = cal.createEvent(o.titulo, new Date(o.inicio), new Date(Math.max(o.fim, o.inicio)), opc);
  }
  ev.setTag('outlook_chave', o.chave);
  ev.setTag('outlook_hash', o.hash);
}

/** Lê o texto ICS e devolve as ocorrências (com repetições expandidas) dentro da janela. */
function lerIcs_(texto, janIni, janFim) {
  TZ_FIXOS_ = {};
  const linhas = texto.replace(/\r?\n[ \t]/g, '').split(/\r?\n/);
  const eventos = [];
  let atual = null, emAlarme = false, tzAtual = null, emStandard = false;

  linhas.forEach(function (linha) {
    if (!linha) return;
    const p = propIcs_(linha);
    if (p.nome === 'BEGIN') {
      if (p.valor === 'VEVENT') { atual = { EXDATE: [] }; }
      else if (p.valor === 'VALARM') emAlarme = true;
      else if (p.valor === 'VTIMEZONE') tzAtual = {};
      else if (p.valor === 'STANDARD') emStandard = true;
      return;
    }
    if (p.nome === 'END') {
      if (p.valor === 'VEVENT' && atual) { eventos.push(atual); atual = null; }
      else if (p.valor === 'VALARM') emAlarme = false;
      else if (p.valor === 'STANDARD') emStandard = false;
      else if (p.valor === 'VTIMEZONE') {
        if (tzAtual && tzAtual.id && tzAtual.offset != null) TZ_FIXOS_[tzAtual.id] = tzAtual.offset;
        tzAtual = null;
      }
      return;
    }
    if (tzAtual && !atual) {
      if (p.nome === 'TZID') tzAtual.id = p.valor;
      if (p.nome === 'TZOFFSETTO' && emStandard && tzAtual.offset == null) tzAtual.offset = offsetTexto_(p.valor);
      return;
    }
    if (!atual || emAlarme) return;
    if (p.nome === 'EXDATE') {
      p.valor.split(',').forEach(function (v) { atual.EXDATE.push({ valor: v, params: p.params }); });
    } else if (['DTSTART', 'DTEND', 'RECURRENCE-ID'].indexOf(p.nome) >= 0) {
      atual[p.nome] = { valor: p.valor, params: p.params };
    } else if (['UID', 'SUMMARY', 'LOCATION', 'STATUS', 'RRULE', 'DURATION'].indexOf(p.nome) >= 0) {
      atual[p.nome] = p.valor;
    }
  });

  // Ocorrências que foram alteradas/canceladas individualmente
  const excecoes = {};
  eventos.forEach(function (e) {
    if (e['RECURRENCE-ID'] && e.UID) {
      const r = tempoIcs_(e['RECURRENCE-ID']);
      (excecoes[e.UID] = excecoes[e.UID] || []).push(paraMs_(r.naive, r.tz));
    }
  });

  const saida = [];
  eventos.forEach(function (e) {
    if (!e.DTSTART || !e.UID) return;
    if (String(e.STATUS || '').toUpperCase() === 'CANCELLED') return;

    const t0 = tempoIcs_(e.DTSTART);
    let duracao;
    if (e.DTEND) duracao = tempoIcs_(e.DTEND).naive - t0.naive;
    else if (e.DURATION) duracao = duracaoIcs_(e.DURATION);
    else duracao = t0.diaTodo ? DIA_MS : 0;
    if (duracao < 0) duracao = 0;

    const inicios = (e.RRULE && !e['RECURRENCE-ID'])
      ? expandirRegra_(e.RRULE, t0, janIni - 2 * DIA_MS, janFim + 2 * DIA_MS)
      : [t0.naive];

    const exclui = (excecoes[e.UID] && !e['RECURRENCE-ID']) ? excecoes[e.UID].slice() : [];
    e.EXDATE.forEach(function (x) {
      const t = tempoIcs_(x);
      exclui.push(t.diaTodo ? 'D' + t.naive : paraMs_(t.naive, t.tz || t0.tz));
    });

    const titulo = textoIcs_(e.SUMMARY) || '(sem título)';
    const local = textoIcs_(e.LOCATION);
    const uidHash = hash_(e.UID);

    inicios.forEach(function (naive) {
      const tzUso = t0.diaTodo ? FUSO : t0.tz;
      const ini = paraMs_(naive, tzUso);
      const fim = paraMs_(naive + duracao, tzUso);
      const diaIni = naive - (naive % DIA_MS);
      const excluido = exclui.some(function (x) {
        return typeof x === 'string' ? x === 'D' + diaIni : Math.abs(x - ini) < 60000;
      });
      if (excluido) return;
      if (!(fim > janIni || (fim === ini && ini >= janIni)) || ini >= janFim) return;
      saida.push({
        chave: uidHash + '|' + ini,
        titulo: titulo, local: local, diaTodo: t0.diaTodo,
        inicio: ini, fim: fim, inicioNaive: naive, fimNaive: naive + duracao,
        hash: hash_([titulo, local, t0.diaTodo, ini, fim].join('|')),
      });
    });
  });
  return saida;
}

function propIcs_(linha) {
  let i = 0, aspas = false;
  for (; i < linha.length; i++) {
    const c = linha.charAt(i);
    if (c === '"') aspas = !aspas;
    else if (c === ':' && !aspas) break;
  }
  const cab = linha.slice(0, i).split(';');
  const params = {};
  cab.slice(1).forEach(function (kv) {
    const j = kv.indexOf('=');
    if (j > 0) params[kv.slice(0, j).toUpperCase()] = kv.slice(j + 1).replace(/^"|"$/g, '');
  });
  return { nome: cab[0].toUpperCase(), params: params, valor: linha.slice(i + 1).trim() };
}

function textoIcs_(s) {
  return String(s || '').replace(/\\n/gi, ' ').replace(/\\([,;\\])/g, '$1').trim();
}

/** Converte data ICS em "horário de parede" (ms como se fosse UTC) + fuso. */
function tempoIcs_(p) {
  const v = String(p.valor).trim();
  const m = v.match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?/);
  if (!m) throw new Error('Data inválida no ICS: ' + v);
  return {
    naive: Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)),
    diaTodo: !m[4] || p.params.VALUE === 'DATE',
    tz: m[7] ? 'UTC' : (p.params.TZID || null),
  };
}

function duracaoIcs_(s) {
  const m = String(s).match(/P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?/);
  if (!m) return 0;
  return ((+(m[1] || 0) * 7 + +(m[2] || 0)) * 24 * 3600 + +(m[3] || 0) * 3600 + +(m[4] || 0) * 60 + +(m[5] || 0)) * 1000;
}

function offsetTexto_(s) {
  const m = String(s).match(/([+-])(\d{2}):?(\d{2})/);
  return m ? (m[1] === '-' ? -1 : 1) * (+m[2] * 60 + +m[3]) : 0;
}

function resolverTz_(tz) {
  if (!tz) return FUSO;
  if (TZ_WINDOWS[tz]) return TZ_WINDOWS[tz];
  const limpo = tz.replace(/^\/+/, '');
  if (/^[A-Za-z]+\/[A-Za-z_\/+-]+$/.test(limpo) && limpo.indexOf('tzone') < 0) return limpo;
  return null;
}

function offsetMin_(iana, ms) {
  const k = iana + '|' + Math.floor(ms / 3600000);
  if (CACHE_OFFSET_[k] == null) CACHE_OFFSET_[k] = offsetTexto_(Utilities.formatDate(new Date(ms), iana, 'Z'));
  return CACHE_OFFSET_[k];
}

/** "Horário de parede" num fuso → instante real (ms). */
function paraMs_(naive, tz) {
  if (tz === 'UTC') return naive;
  const iana = resolverTz_(tz);
  if (iana) {
    const off1 = offsetMin_(iana, naive);
    let ms = naive - off1 * 60000;
    const off2 = offsetMin_(iana, ms);
    if (off2 !== off1) ms = naive - off2 * 60000;
    return ms;
  }
  const fixo = TZ_FIXOS_[tz] != null ? TZ_FIXOS_[tz] : -180;
  return naive - fixo * 60000;
}

function naiveDeMs_(ms, tz) {
  if (tz === 'UTC') return ms;
  const iana = resolverTz_(tz);
  if (iana) return ms + offsetMin_(iana, ms) * 60000;
  return ms + (TZ_FIXOS_[tz] != null ? TZ_FIXOS_[tz] : -180) * 60000;
}

/** Expande uma RRULE (DAILY/WEEKLY/MONTHLY/YEARLY) em inícios ("horário de parede"). */
function expandirRegra_(regraTxt, t0, desdeMs, ateMs) {
  const r = {};
  regraTxt.split(';').forEach(function (kv) {
    const j = kv.indexOf('=');
    if (j > 0) r[kv.slice(0, j).toUpperCase()] = kv.slice(j + 1);
  });
  const freq = r.FREQ;
  const intervalo = Math.max(1, Number(r.INTERVAL || 1));
  const count = r.COUNT ? Number(r.COUNT) : null;
  const COD = { SU: 0, MO: 1, TU: 2, WE: 3, TH: 4, FR: 5, SA: 6 };
  const byday = r.BYDAY ? r.BYDAY.split(',').map(function (s) {
    const m = s.trim().match(/^([+-]?\d+)?([A-Z]{2})$/);
    return m ? { n: m[1] ? Number(m[1]) : null, wd: COD[m[2]] } : null;
  }).filter(Boolean) : null;
  const bymonthday = r.BYMONTHDAY ? r.BYMONTHDAY.split(',').map(Number) : null;
  const bymonth = r.BYMONTH ? r.BYMONTH.split(',').map(Number) : null;
  const bysetpos = r.BYSETPOS ? r.BYSETPOS.split(',').map(Number) : null;

  let ate = null;
  if (r.UNTIL) {
    const u = tempoIcs_({ valor: r.UNTIL, params: {} });
    if (u.tz === 'UTC') ate = naiveDeMs_(u.naive, t0.tz);
    else ate = u.diaTodo ? u.naive + DIA_MS - 1 : u.naive;
  }

  const hora = ((t0.naive % DIA_MS) + DIA_MS) % DIA_MS;
  const dia0 = t0.naive - hora;
  const s = new Date(dia0);
  const segunda = function (d) { return d - ((new Date(d).getUTCDay() + 6) % 7) * DIA_MS; };

  function casaNoMes(dt) {
    const dia = dt.getUTCDate(), dow = dt.getUTCDay();
    const ultimo = new Date(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth() + 1, 0)).getUTCDate();
    if (bymonthday) return bymonthday.some(function (x) { return (x > 0 ? x : ultimo + x + 1) === dia; });
    if (byday) {
      if (bysetpos) {
        const dias = byday.map(function (b) { return b.wd; });
        const cand = [];
        for (let i = 1; i <= ultimo; i++) {
          if (dias.indexOf((((dow + i - dia) % 7) + 7) % 7) >= 0) cand.push(i);
        }
        return bysetpos.some(function (p) { return cand[p > 0 ? p - 1 : cand.length + p] === dia; });
      }
      return byday.some(function (b) {
        if (b.wd !== dow) return false;
        if (b.n == null) return true;
        return b.n > 0 ? Math.ceil(dia / 7) === b.n : Math.ceil((ultimo - dia + 1) / 7) === -b.n;
      });
    }
    return dia === s.getUTCDate();
  }

  function casa(d) {
    if (d === dia0) return true;
    const dt = new Date(d);
    const mes = dt.getUTCMonth() + 1, ano = dt.getUTCFullYear(), dow = dt.getUTCDay();
    if (bymonth && bymonth.indexOf(mes) < 0) return false;
    if (freq === 'DAILY') {
      if (Math.round((d - dia0) / DIA_MS) % intervalo) return false;
      return !byday || byday.some(function (b) { return b.wd === dow; });
    }
    if (freq === 'WEEKLY') {
      if (Math.round((segunda(d) - segunda(dia0)) / (7 * DIA_MS)) % intervalo) return false;
      return byday ? byday.some(function (b) { return b.wd === dow; }) : dow === s.getUTCDay();
    }
    if (freq === 'MONTHLY') {
      const meses = (ano - s.getUTCFullYear()) * 12 + (mes - (s.getUTCMonth() + 1));
      return meses % intervalo === 0 && casaNoMes(dt);
    }
    if (freq === 'YEARLY') {
      if ((ano - s.getUTCFullYear()) % intervalo) return false;
      if (!bymonth && mes !== s.getUTCMonth() + 1) return false;
      return casaNoMes(dt);
    }
    return false;
  }

  // Sem COUNT dá para começar perto da janela; com COUNT é preciso contar desde o início.
  const desdeNaive = desdeMs - DIA_MS;
  const ateNaive = ateMs + DIA_MS;
  let d = (!count && desdeNaive > dia0) ? desdeNaive - (((desdeNaive % DIA_MS) + DIA_MS) % DIA_MS) : dia0;
  const res = [];
  let n = 0;
  for (let iter = 0; d <= ateNaive && iter < 8000; d += DIA_MS, iter++) {
    if (!casa(d)) continue;
    const occ = d + hora;
    if (ate != null && occ > ate) break;
    n++;
    if (count && n > count) break;
    if (occ >= desdeNaive) res.push(occ);
  }
  return res;
}

function dataDeNaive_(naive) {
  return parseData_(new Date(naive).toISOString().slice(0, 10));
}

function pausa_() { Utilities.sleep(150); } // evita o limite de "muitas alterações seguidas" do Google Agenda


// =================================================================
//  ROTINAS AUTOMÁTICAS (gatilhos)
// =================================================================

function resumoDiario() {
  const dono = prop_('DONO_CHAT_ID');
  if (dono) enviar_(dono, montarResumo_('☀️ Bom dia!'));
}

function montarResumo_(titulo) {
  const hoje = hojeInicio_();
  const amanha = new Date(hoje.getTime() + DIA_MS);
  const ontem = new Date(hoje.getTime() - DIA_MS);
  const inicioMes = parseData_(Utilities.formatDate(new Date(), FUSO, 'yyyy-MM') + '-01');

  const eventos = eventosEntre_(hoje, amanha);
  const agenda = eventos.length
    ? eventos.map(function (x) {
        const ev = x.ev;
        const hora = ev.isAllDayEvent() ? 'Dia todo' : Utilities.formatDate(ev.getStartTime(), FUSO, 'HH:mm');
        return '• ' + hora + ' – ' + marca_(x) + ev.getTitle();
      }).join('\n')
    : 'Nada marcado 🙌';

  const gastos = lerGastos_();
  const soma = function (de, ate) {
    return gastos.filter(function (g) { return g.data >= de && g.data < ate; })
                 .reduce(function (t, g) { return t + g.valor; }, 0);
  };
  const porCat = {};
  gastos.filter(function (g) { return g.data >= inicioMes && g.data < amanha; })
        .forEach(function (g) { porCat[g.categoria] = (porCat[g.categoria] || 0) + g.valor; });
  const top = Object.keys(porCat)
    .sort(function (x, y) { return porCat[y] - porCat[x]; }).slice(0, 3)
    .map(function (c) { return c + ' ' + moeda_(porCat[c]); }).join(' | ');

  return titulo + ' Hoje é ' + diaSemana_(new Date()) + ', ' + Utilities.formatDate(new Date(), FUSO, 'dd/MM') + '.\n\n' +
    '📅 Agenda de hoje\n' + agenda + '\n\n' +
    '💰 Gastos\nOntem: ' + moeda_(soma(ontem, hoje)) +
    '\nNo mês: ' + moeda_(soma(inicioMes, amanha)) +
    (top ? '\nMaiores: ' + top : '');
}

function verificarLembretes() {
  const dono = prop_('DONO_CHAT_ID');
  if (!dono) return;
  const limites = LEMBRETES_MINUTOS.slice().sort(function (a, b) { return a - b; });
  const maior = limites[limites.length - 1];
  const agora = new Date();
  const cache = CacheService.getScriptCache();

  eventosEntre_(agora, new Date(agora.getTime() + maior * 60000)).forEach(function (x) {
    const ev = x.ev;
    if (ev.isAllDayEvent() || ev.getStartTime() < agora) return;
    const faltam = (ev.getStartTime().getTime() - agora.getTime()) / 60000;
    const aviso = limites.filter(function (m) { return m >= faltam; })[0]; // menor aviso que já "venceu"
    if (aviso == null) return;
    const base = 'lemb_' + hash_(ev.getId() + '|' + ev.getStartTime().getTime()) + '_';
    if (cache.get(base + aviso)) return;
    limites.filter(function (m) { return m >= aviso; }).forEach(function (m) { cache.put(base + m, '1', 21600); });

    const min = Math.max(1, Math.round(faltam));
    const quando = min >= 60 ? Math.floor(min / 60) + 'h' + (min % 60 ? String(min % 60).padStart(2, '0') : '') : min + ' min';
    enviar_(dono, '⏰ ' + marca_(x) + 'Em ' + quando + ': ' + ev.getTitle() +
      ' (' + Utilities.formatDate(ev.getStartTime(), FUSO, 'HH:mm') + ')' +
      (ev.getLocation() ? '\n📍 ' + ev.getLocation() : ''));
  });
}


// =================================================================
//  UTILITÁRIOS
// =================================================================

function props_() { return PropertiesService.getScriptProperties(); }
function prop_(k) { return props_().getProperty(k); }
function ok_() { return HtmlService.createHtmlOutput('ok'); }

function tg_(metodo, dados) {
  const resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + prop_('TELEGRAM_TOKEN') + '/' + metodo, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(dados || {}),
    muteHttpExceptions: true,
  });
  return JSON.parse(resp.getContentText());
}

function enviar_(chatId, texto) {
  texto = String(texto || '').trim() || '(sem resposta)';
  for (let i = 0; i < texto.length; i += 4000) {
    tg_('sendMessage', { chat_id: chatId, text: texto.slice(i, i + 4000) });
  }
}

function baixarArquivo_(fileId, mime) {
  const info = tg_('getFile', { file_id: fileId });
  if (!info.ok) throw new Error('Não consegui baixar o arquivo do Telegram (limite de 20 MB).');
  const blob = UrlFetchApp.fetch('https://api.telegram.org/file/bot' + prop_('TELEGRAM_TOKEN') + '/' + info.result.file_path).getBlob();
  return { inlineData: { mimeType: mime, data: Utilities.base64Encode(blob.getBytes()) } };
}

function jaProcessado_(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cache = CacheService.getScriptCache();
    const chave = 'upd_' + id;
    if (cache.get(chave)) return true;
    cache.put(chave, '1', 21600);
    return false;
  } finally {
    lock.releaseLock();
  }
}

function avisarErro_(err) {
  console.error(err && err.stack ? err.stack : err);
  const dono = prop_('DONO_CHAT_ID');
  if (dono) {
    try { enviar_(dono, '⚠️ Deu um erro: ' + (err && err.message ? err.message : err)); } catch (e) { /* ignora */ }
  }
}

function removerGatilho_(nomeFuncao) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === nomeFuncao) ScriptApp.deleteTrigger(t);
  });
}

function parseData_(s) {
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) throw new Error('Data inválida: ' + s);
  return new Date(m[1] + '-' + m[2] + '-' + m[3] + 'T00:00:00' + OFFSET);
}

function parseDataHora_(s) {
  s = String(s).trim().replace(' ', 'T');
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) s += 'T09:00:00';
  if (/T\d{2}:\d{2}$/.test(s)) s += ':00';
  if (!/([zZ]|[+-]\d{2}:?\d{2})$/.test(s)) s += OFFSET;
  const d = new Date(s);
  if (isNaN(d.getTime())) throw new Error('Data/hora inválida: ' + s);
  return d;
}

function hojeInicio_() { return parseData_(Utilities.formatDate(new Date(), FUSO, 'yyyy-MM-dd')); }
function fmtData_(d) { return Utilities.formatDate(d, FUSO, 'dd/MM/yyyy'); }
function fmtDataHora_(d) { return Utilities.formatDate(d, FUSO, 'dd/MM/yyyy HH:mm'); }
function diaSemana_(d) { return DIAS_SEMANA[Number(Utilities.formatDate(d, FUSO, 'u')) % 7]; }

function moeda_(v) {
  const partes = Math.abs(Number(v) || 0).toFixed(2).split('.');
  return (v < 0 ? '-' : '') + 'R$ ' + partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + partes[1];
}

function hash_(s) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s));
}
