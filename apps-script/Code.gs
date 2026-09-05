const CFG = {
  // Estrutura atual da Nova Camisa — mantida por compatibilidade e estabilidade.
  SHEET_PEDIDOS: 'Pedidos',
  SHEET_PRODUTOS: 'Produtos',
  SHEET_TAMANHOS: 'Tamanhos',
  SHEET_CONFIG: 'Configurações',
  SHEET_LOG: 'Log',

  // Commerce Core — fundação reutilizável para futuras vendas do Grêmio.
  SHEET_VARIANTES: 'Variantes',
  SHEET_ITENS_PEDIDO: 'ItensPedido',
  SHEET_CLIENTES: 'Clientes',
  SHEET_PAGAMENTOS: 'Pagamentos',
  SHEET_ESTOQUE: 'Estoque',
  SHEET_CUPONS: 'Cupons',
  SHEET_FINANCEIRO: 'Movimentações Financeiras',
  SHEET_ENTREGAS: 'Entregas',
  SHEET_CATEGORIAS: 'Categorias',
  SHEET_CAMPANHAS: 'Campanhas',
  SHEET_ASAAS_EVENTOS: 'Eventos Asaas',
  SHEET_RELACAO_GRAFICA: 'Relação Gráfica',

  PAGE_DEFAULT: 'estudantes',
  ADMIN_CACHE_PREFIX: 'CAMISAS_ADMIN_SESSION_',
  ADMIN_SESSION_SECONDS: 21600,
  COMMERCE_VERSION: '7.4.0',
  CARD_MAX_INSTALLMENTS: 12,
  CARD_MIN_INSTALLMENT_VALUE: 5
};

/**
 * MOCKUPS DA PÁGINA "CONHEÇA OS MODELOS"
 * ----------------------------------------
 * Quando as imagens finais estiverem prontas, envie-as ao Google Drive
 * e cole SOMENTE o ID de cada arquivo entre as aspas abaixo.
 * Enquanto estiver vazio, o site mostra um placeholder profissional.
 */
const MOCKUP_IDS = {
  // Fallback legado. Os valores oficiais agora são administrados pela Área do Grêmio
  // e ficam salvos na aba Configurações.
  estudante: '',
  servidorPolo: '',
  servidorCareca: ''
};

const MOCKUP_CONFIG_KEYS = {
  estudante: 'ID mockup estudantes',
  servidorPolo: 'ID mockup profissionais polo',
  servidorCareca: 'ID mockup profissionais careca'
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛍️ Loja do Grêmio')
    .addItem('Validar/formatar sistema', 'configurarSistema')
    .addItem('Preparar estrutura futura da Loja', 'configurarCommerceCore')
    .addItem('Mostrar links do Web App', 'mostrarLinksWebApp')
    .addSeparator()
    .addItem('🔌 Verificar conexão Asaas', 'testarConexaoAsaas')
    .addItem('✅ Validar ativação oficial Asaas', 'validarAtivacaoOficialAsaas')
    .addItem('🩺 Diagnóstico Asaas', 'diagnosticarAsaas')
    .addItem('💳 Consultar parcelamento Asaas', 'testarParcelamentoCartaoSelecionado')
    .addItem('🔄 Sincronizar pagamentos agora', 'sincronizarPagamentosAsaas')
    .addSeparator()
    .addItem('🏭 Gerar relação para gráfica', 'gerarRelacaoGrafica')
    .addItem('✅ Marcar relação atual como enviada à gráfica', 'marcarRelacaoGraficaComoEnviada')
    .addSeparator()
    .addItem('🔐 Configurar acesso da Área do Grêmio', 'configurarCodigoAdminPeloMenu')
    .addSeparator()
    .addItem('🌐 Configurar ponte Vercel', 'configurarPonteVercel')
    .addItem('🌐 Mostrar dados da ponte Vercel', 'mostrarDadosPonteVercel')
    .addSeparator()
    .addItem('Abrir Dashboard', 'abrirDashboard')
    .addToUi();
}


/**
 * Gera uma relação limpa para envio à gráfica.
 *
 * Regras:
 * - inclui apenas pedidos com pagamento PAGO ou CONFIRMADO;
 * - inclui apenas pedidos ainda NÃO ENVIADOS à gráfica;
 * - inclui somente os produtos atuais de camisas;
 * - estudantes aparecem consolidados por tamanho;
 * - profissionais aparecem detalhados por modelo/tamanho, com personalização.
 *
 * A aba "Relação Gráfica" é um snapshot do momento da geração. As colunas
 * internas de controle ficam ocultas e permitem marcar exatamente aqueles
 * pedidos como "ENVIADO À GRÁFICA" depois que a relação for enviada.
 */
function gerarRelacaoGrafica() {
  const ss = SpreadsheetApp.getActive();
  const shPedidos = ss.getSheetByName(CFG.SHEET_PEDIDOS);
  if (!shPedidos) throw new Error('A aba Pedidos não foi encontrada.');

  const lastRow = shPedidos.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Ainda não existem pedidos para gerar a relação da gráfica.');
    return;
  }

  const dados = shPedidos.getRange(2, 1, lastRow - 1, 38).getValues();
  const codigosCamisa = ['EST-POLO', 'SERV-POLO', 'SERV-CARECA'];
  const statusOk = ['PAGO', 'CONFIRMADO'];

  const pedidos = dados.filter(r => {
    const pagamento = String(r[18] || '').trim().toUpperCase();
    const producao = String(r[19] || '').trim().toUpperCase();
    const codigo = String(r[7] || '').trim().toUpperCase();
    return statusOk.includes(pagamento)
      && producao === 'NÃO ENVIADO'
      && codigosCamisa.includes(codigo);
  });

  if (!pedidos.length) {
    SpreadsheetApp.getUi().alert(
      'Não há camisas pagas/confirmadas pendentes de envio à gráfica.\n\n' +
      'A relação considera somente pedidos com status PAGO ou CONFIRMADO e produção = NÃO ENVIADO.'
    );
    return;
  }

  const ordemTamanhos = ['PP', 'P', 'M', 'G', 'GG', 'XGG', 'EXGG'];
  const rankTamanho = t => {
    const i = ordemTamanhos.indexOf(String(t || '').trim().toUpperCase());
    return i >= 0 ? i : 999;
  };

  const estudantes = pedidos.filter(r => String(r[2] || '').trim().toUpperCase() === 'ESTUDANTE');
  const profissionais = pedidos.filter(r => String(r[2] || '').trim().toUpperCase() === 'SERVIDOR');

  const contagem = {};
  pedidos.forEach(r => {
    const codigo = String(r[7] || '').trim().toUpperCase();
    const tam = String(r[9] || '').trim().toUpperCase() || 'SEM TAMANHO';
    const chave = codigo + '|' + tam;
    contagem[chave] = (contagem[chave] || 0) + 1;
  });

  const porCodigoETamanho = (codigo, tam) => contagem[codigo + '|' + tam] || 0;

  let sh = ss.getSheetByName(CFG.SHEET_RELACAO_GRAFICA);
  if (!sh) {
    sh = ss.insertSheet(CFG.SHEET_RELACAO_GRAFICA);
  } else {
    sh.showSheet();
    sh.clear();
    sh.clearConditionalFormatRules();
    sh.getRange(1, 1, sh.getMaxRows(), Math.min(sh.getMaxColumns(), 12)).breakApart();
  }

  if (sh.getMaxColumns() < 12) sh.insertColumnsAfter(sh.getMaxColumns(), 12 - sh.getMaxColumns());
  const minRowsRelacao = Math.max(200, pedidos.length + 100);
  if (sh.getMaxRows() < minRowsRelacao) sh.insertRowsAfter(sh.getMaxRows(), minRowsRelacao - sh.getMaxRows());

  const env = getAsaasConfig_().env;
  if (env !== 'PRODUCTION') {
    throw new Error('A relação oficial para a gráfica só pode ser gerada depois que o Asaas de Produção estiver ativado.');
  }
  const agora = new Date();
  const geradoEm = Utilities.formatDate(agora, 'America/Fortaleza', 'dd/MM/yyyy HH:mm');
  const total = pedidos.length;

  sh.getRange('A1:I1').merge().setValue('RELAÇÃO PARA PRODUÇÃO — NOVA CAMISA OFICIAL')
    .setBackground('#063667').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(15).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);

  sh.getRange('A2:I2').merge().setValue('RELAÇÃO OFICIAL PARA PRODUÇÃO')
    .setBackground('#E8F5E9')
    .setFontColor('#1B5E20')
    .setFontWeight('bold').setHorizontalAlignment('center');

  sh.getRange('A3:B4').setValues([
    ['Gerado em', geradoEm],
    ['Total de peças', total]
  ]);
  sh.getRange('A3:A4').setFontWeight('bold').setBackground('#EAF1F8');
  sh.getRange('B3:B4').setFontWeight('bold');

  let row = 6;
  sh.getRange(row, 1, 1, 9).merge().setValue('RESUMO DE PRODUÇÃO')
    .setBackground('#0D4F8B').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('left');
  row++;

  const resumoHeaders = ['Modelo', ...ordemTamanhos, 'Total'];
  sh.getRange(row, 1, 1, resumoHeaders.length).setValues([resumoHeaders])
    .setBackground('#D9E8F5').setFontWeight('bold').setHorizontalAlignment('center');
  row++;

  const linhasResumo = [
    ['Estudantes — Gola Polo', ...ordemTamanhos.map(t => porCodigoETamanho('EST-POLO', t)),
      estudantes.length],
    ['Profissionais — Gola Polo', ...ordemTamanhos.map(t => porCodigoETamanho('SERV-POLO', t)),
      profissionais.filter(r => String(r[7] || '').trim().toUpperCase() === 'SERV-POLO').length],
    ['Profissionais — Gola Careca', ...ordemTamanhos.map(t => porCodigoETamanho('SERV-CARECA', t)),
      profissionais.filter(r => String(r[7] || '').trim().toUpperCase() === 'SERV-CARECA').length]
  ];
  sh.getRange(row, 1, linhasResumo.length, resumoHeaders.length).setValues(linhasResumo);
  sh.getRange(row, 2, linhasResumo.length, resumoHeaders.length - 1).setHorizontalAlignment('center');
  row += linhasResumo.length + 2;

  sh.getRange(row, 1, 1, 3).merge().setValue('ESTUDANTES — GOLA POLO')
    .setBackground('#0D4F8B').setFontColor('#FFFFFF').setFontWeight('bold');
  row++;
  sh.getRange(row, 1, 1, 2).setValues([['Tamanho', 'Quantidade']])
    .setBackground('#D9E8F5').setFontWeight('bold').setHorizontalAlignment('center');
  row++;

  const tamanhosEstudantesExtras = [...new Set(estudantes.map(r => String(r[9] || '').trim().toUpperCase()).filter(Boolean))]
    .filter(t => !ordemTamanhos.includes(t)).sort();
  const tamanhosEstudantes = [...ordemTamanhos, ...tamanhosEstudantesExtras];
  const linhasEst = tamanhosEstudantes
    .map(t => [t, porCodigoETamanho('EST-POLO', t)])
    .filter(r => r[1] > 0);

  if (linhasEst.length) {
    sh.getRange(row, 1, linhasEst.length, 2).setValues(linhasEst);
    sh.getRange(row, 1, linhasEst.length, 2).setHorizontalAlignment('center');
    row += linhasEst.length;
  } else {
    sh.getRange(row, 1, 1, 2).merge().setValue('Nenhum pedido de estudante pendente neste lote.')
      .setFontStyle('italic').setFontColor('#666666');
    row++;
  }
  sh.getRange(row, 1).setValue('TOTAL').setFontWeight('bold');
  sh.getRange(row, 2).setValue(estudantes.length).setFontWeight('bold').setHorizontalAlignment('center');
  row += 2;

  const escreverProfissionais = (titulo, codigo) => {
    sh.getRange(row, 1, 1, 6).merge().setValue(titulo)
      .setBackground('#0D4F8B').setFontColor('#FFFFFF').setFontWeight('bold');
    row++;
    const headers = ['Tamanho', 'Nome do profissional', 'Nome na camisa', 'Cargo/Área/Função', 'Qtd.', 'Pedido'];
    sh.getRange(row, 1, 1, headers.length).setValues([headers])
      .setBackground('#D9E8F5').setFontWeight('bold').setHorizontalAlignment('center');
    row++;

    const itens = profissionais
      .filter(r => String(r[7] || '').trim().toUpperCase() === codigo)
      .sort((a, b) => {
        const tamA = String(a[9] || '').trim().toUpperCase();
        const tamB = String(b[9] || '').trim().toUpperCase();
        const d = rankTamanho(tamA) - rankTamanho(tamB);
        if (d !== 0) return d;
        if (tamA !== tamB && d === 0) return tamA.localeCompare(tamB, 'pt-BR');
        return String(a[3] || '').localeCompare(String(b[3] || ''), 'pt-BR');
      })
      .map(r => [
        String(r[9] || '').trim().toUpperCase() || '—',
        String(r[3] || '').trim() || '—',
        String(r[10] || '').trim() || 'SEM PERSONALIZAÇÃO',
        String(r[11] || '').trim() || '—',
        1,
        String(r[0] || '').trim()
      ]);

    if (itens.length) {
      sh.getRange(row, 1, itens.length, headers.length).setValues(itens);
      sh.getRange(row, 1, itens.length, 1).setHorizontalAlignment('center');
      sh.getRange(row, 5, itens.length, 1).setHorizontalAlignment('center');
      row += itens.length;
      sh.getRange(row, 1, 1, 4).merge().setValue('TOTAL').setFontWeight('bold').setHorizontalAlignment('right');
      sh.getRange(row, 5).setValue(itens.length).setFontWeight('bold').setHorizontalAlignment('center');
      row++;
    } else {
      sh.getRange(row, 1, 1, 6).merge().setValue('Nenhum pedido pendente deste modelo.')
        .setFontStyle('italic').setFontColor('#666666');
      row++;
    }
    row += 1;
  };

  escreverProfissionais('PROFISSIONAIS — GOLA POLO', 'SERV-POLO');
  escreverProfissionais('PROFISSIONAIS — GOLA CARECA', 'SERV-CARECA');

  sh.getRange(row, 1, 1, 9).merge().setValue(
    'Relação gerada somente com pedidos PAGO/CONFIRMADO e ainda NÃO ENVIADOS à gráfica. ' +
    'Após enviar esta relação, use o menu “✅ Marcar relação atual como enviada à gráfica”.'
  ).setBackground('#F3F6F9').setFontColor('#4B5563').setFontStyle('italic').setWrap(true);
  row += 2;

  sh.getRange(1, 10, 1, 2).setValues([['Pedido (controle interno)', 'Linha Pedidos']]);
  const linhaPorPedido = {};
  dados.forEach((r, i) => {
    const id = String(r[0] || '').trim();
    if (id) linhaPorPedido[id.toUpperCase()] = i + 2;
  });
  const controles = pedidos.map(r => {
    const id = String(r[0] || '').trim();
    return [id, linhaPorPedido[id.toUpperCase()] || ''];
  });
  if (controles.length) sh.getRange(2, 10, controles.length, 2).setValues(controles);
  sh.hideColumns(10, 2);

  sh.setFrozenRows(2);
  sh.getRange(1, 1, Math.max(row, 10), 9).setVerticalAlignment('middle').setWrap(true);
  sh.getRange(1, 1, Math.max(row, 10), 9).setBorder(true, true, true, true, true, true, '#D5DEE7', SpreadsheetApp.BorderStyle.SOLID);
  sh.setColumnWidth(1, 190);
  sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 185);
  sh.setColumnWidth(4, 220);
  sh.setColumnWidth(5, 80);
  sh.setColumnWidth(6, 125);
  sh.setColumnWidth(7, 90);
  sh.setColumnWidth(8, 90);
  sh.setColumnWidth(9, 90);
  sh.setRowHeights(1, Math.max(row, 10), 24);
  sh.setHiddenGridlines(true);
  sh.activate();

  registrarLog_(
    'RELAÇÃO GRÁFICA GERADA',
    '',
    'SISTEMA',
    `${total} peça(s) • ${estudantes.length} estudante(s) • ${profissionais.length} profissional(is) • ${env}`
  );

  ss.toast(
    `Relação oficial gerada com ${total} peça(s), pronta para conferência e envio à gráfica.`,
    '🏭 Relação para gráfica',
    8
  );
}

/**
 * Marca como ENVIADO À GRÁFICA exatamente os pedidos que constam na relação
 * atualmente gerada. Isso evita marcar pedidos novos que chegaram depois.
 */
function marcarRelacaoGraficaComoEnviada() {
  const ss = SpreadsheetApp.getActive();
  const env = getAsaasConfig_().env;

  if (env !== 'PRODUCTION') {
    SpreadsheetApp.getUi().alert(
      'Pagamentos oficiais ainda não ativados',
      'Ative e valide o Asaas de Produção antes de enviar pedidos para fabricação.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const shRelacao = ss.getSheetByName(CFG.SHEET_RELACAO_GRAFICA);
  const shPedidos = ss.getSheetByName(CFG.SHEET_PEDIDOS);
  if (!shRelacao || !shPedidos) throw new Error('A relação ou a aba Pedidos não foi encontrada.');

  const lastRel = shRelacao.getLastRow();
  const controles = lastRel >= 2 ? shRelacao.getRange(2, 10, Math.max(lastRel - 1, 1), 2).getValues() : [];
  const ids = controles.map(r => String(r[0] || '').trim().toUpperCase()).filter(Boolean);

  if (!ids.length) {
    SpreadsheetApp.getUi().alert('Não há pedidos registrados na relação atual.');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const resposta = ui.alert(
    'Confirmar envio à gráfica',
    `Esta ação marcará até ${ids.length} pedido(s) como ENVIADO À GRÁFICA.\n\nUse somente depois de realmente enviar esta relação para produção.`,
    ui.ButtonSet.YES_NO
  );
  if (resposta !== ui.Button.YES) return;

  const cfg = getConfigMap_();
  const prazoDias = Number(cfg['Prazo padrão produção (dias)'] || 15);
  const agora = new Date();
  const prazo = new Date();
  prazo.setDate(prazo.getDate() + prazoDias);
  prazo.setHours(12, 0, 0, 0);

  const lastPedidos = shPedidos.getLastRow();
  if (lastPedidos < 2) return;
  const qtdRows = lastPedidos - 1;
  const colPedido = shPedidos.getRange(2, 1, qtdRows, 1).getDisplayValues();
  const colPagamento = shPedidos.getRange(2, 19, qtdRows, 1).getDisplayValues();
  const colProducao = shPedidos.getRange(2, 20, qtdRows, 1).getValues();
  const colPrazo = shPedidos.getRange(2, 26, qtdRows, 1).getValues();
  const colAtualizacao = shPedidos.getRange(2, 27, qtdRows, 1).getValues();
  const colAtualizadoPor = shPedidos.getRange(2, 28, qtdRows, 1).getValues();

  const idSet = new Set(ids);
  let alterados = 0;
  const idsAlterados = [];

  for (let i = 0; i < qtdRows; i++) {
    const id = String(colPedido[i][0] || '').trim().toUpperCase();
    if (!idSet.has(id)) continue;

    const statusPagamento = String(colPagamento[i][0] || '').trim().toUpperCase();
    const statusProducao = String(colProducao[i][0] || '').trim().toUpperCase();
    if (!['PAGO', 'CONFIRMADO'].includes(statusPagamento) || statusProducao !== 'NÃO ENVIADO') continue;

    colProducao[i][0] = 'ENVIADO À GRÁFICA';
    colPrazo[i][0] = prazo;
    colAtualizacao[i][0] = agora;
    colAtualizadoPor[i][0] = 'RELAÇÃO GRÁFICA';
    alterados++;
    idsAlterados.push(id);
  }

  if (alterados) {
    shPedidos.getRange(2, 20, qtdRows, 1).setValues(colProducao);
    shPedidos.getRange(2, 26, qtdRows, 1).setValues(colPrazo);
    shPedidos.getRange(2, 27, qtdRows, 1).setValues(colAtualizacao);
    shPedidos.getRange(2, 28, qtdRows, 1).setValues(colAtualizadoPor);
    registrarLog_(
      'LOTE ENVIADO À GRÁFICA',
      '',
      'RELAÇÃO GRÁFICA',
      `${alterados} pedido(s) • prazo previsto ${formatDateOnly_(prazo)} • ${idsAlterados.slice(0, 25).join(', ')}${idsAlterados.length > 25 ? '…' : ''}`
    );
  }

  ss.toast(
    `${alterados} pedido(s) marcados como ENVIADO À GRÁFICA.`,
    '✅ Produção atualizada',
    8
  );
}


function doGet(e) {
  const page = ((e && e.parameter && e.parameter.p) || CFG.PAGE_DEFAULT).toLowerCase();
  const allowed = ['estudantes', 'servidores', 'modelos', 'acompanhar', 'admin'];
  const safePage = allowed.includes(page) ? page : CFG.PAGE_DEFAULT;

  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.page = safePage;
  tpl.bootJson = JSON.stringify(getPublicData_()).replace(/</g, '\\u003c');

  return tpl.evaluate()
    .setTitle(safePage === 'admin' ? 'Área do Grêmio — Loja Digital' : (safePage === 'modelos' ? 'Conheça os Modelos — Nova Camisa Oficial' : 'Loja Digital do Grêmio — Nova Camisa Oficial'))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/**
 * API privada para o front-end hospedado na Vercel.
 *
 * A Vercel chama este endpoint de servidor para servidor usando um segredo
 * guardado nas Script Properties e nas Environment Variables da Vercel.
 * O navegador do aluno nunca recebe esse segredo.
 */
function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const payload = JSON.parse(raw || '{}');

    const expected = String(
      PropertiesService.getScriptProperties().getProperty('VERCEL_PROXY_SECRET') || ''
    ).trim();
    const received = String(payload.secret || '').trim();

    if (!expected) {
      throw new Error('A ponte Vercel ainda não foi configurada no Apps Script.');
    }
    if (!received || received !== expected) {
      throw new Error('Acesso não autorizado à API.');
    }

    const method = String(payload.method || '').trim();
    const args = Array.isArray(payload.args) ? payload.args : [];

    const result = executarRpcVercel_(method, args);
    return jsonOutput_({ ok: true, result: result });
  } catch (err) {
    return jsonOutput_({
      ok: false,
      error: {
        message: err && err.message ? String(err.message) : 'Erro interno da API.'
      }
    });
  }
}

function executarRpcVercel_(method, args) {
  const handlers = {
    getPublicData: function() {
      return getPublicData_();
    },
    getLogosPublicas: function() {
      return getLogosPublicas();
    },
    getMockupsPublicos: function() {
      return getMockupsPublicos();
    },
    getCommerceCapabilities: function() {
      return getCommerceCapabilities_();
    },
    simularOpcoesCartao: function() {
      return simularOpcoesCartao(args[0]);
    },
    processAsaasWebhookEvent: function() {
      return processAsaasWebhookEvent(args[0] || {});
    },
    criarPedido: function() {
      return criarPedido(args[0] || {});
    },
    consultarPedido: function() {
      return consultarPedido(args[0], args[1], Boolean(args[2]));
    },
    adminLogin: function() {
      return adminLogin(args[0], args[1]);
    },
    adminLogout: function() {
      return adminLogout(args[0]);
    },
    adminGetData: function() {
      return adminGetData(args[0], args[1] || {});
    },
    adminUpdateSettings: function() {
      return adminUpdateSettings(args[0], args[1] || {});
    },
    adminGetMockups: function() {
      return adminGetMockups(args[0]);
    },
    adminUpdateMockups: function() {
      return adminUpdateMockups(args[0], args[1] || {});
    },
    adminGetOrderHistory: function() {
      return adminGetOrderHistory(args[0], args[1]);
    },
    adminUpdateOrder: function() {
      return adminUpdateOrder(args[0], args[1] || {});
    },
    adminResendStatusEmail: function() {
      return adminResendStatusEmail(args[0], args[1]);
    }
  };

  if (!handlers[method]) {
    throw new Error('Método não permitido: ' + method);
  }
  return handlers[method]();
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Gera um segredo novo para a ponte Vercel e exibe junto com a URL do backend.
 * Não salva o segredo na planilha.
 */
function configurarPonteVercel() {
  const secret = (
    Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '')
  );

  PropertiesService.getScriptProperties().setProperty('VERCEL_PROXY_SECRET', secret);

  const backendUrl = ScriptApp.getService().getUrl() || '';
  const msg =
    'PONTE VERCEL CONFIGURADA\\n\\n' +
    'Na Vercel, cadastre estas Environment Variables:\\n\\n' +
    'APPS_SCRIPT_API_URL\\n' + backendUrl + '\\n\\n' +
    'APPS_SCRIPT_PROXY_SECRET\\n' + secret + '\\n\\n' +
    'IMPORTANTE: não coloque esse segredo no HTML nem envie em mensagens.';

  try {
    SpreadsheetApp.getUi().alert(
      '🌐 Ponte Vercel',
      msg,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (_) {}

  return {
    ok: true,
    backendUrl: backendUrl,
    secret: secret
  };
}

function mostrarDadosPonteVercel() {
  const backendUrl = ScriptApp.getService().getUrl() || '';
  const configured = Boolean(
    PropertiesService.getScriptProperties().getProperty('VERCEL_PROXY_SECRET')
  );

  const msg =
    'Backend Apps Script:\\n' + backendUrl + '\\n\\n' +
    'Segredo Vercel configurado: ' + (configured ? 'SIM' : 'NÃO') + '\\n\\n' +
    (configured
      ? 'Por segurança, o segredo existente não é exibido novamente. Se precisar de outro, execute “Configurar ponte Vercel” para gerar um novo.'
      : 'Execute “Configurar ponte Vercel” para gerar o segredo.');

  try {
    SpreadsheetApp.getUi().alert(
      '🌐 Dados da ponte Vercel',
      msg,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (_) {}

  return {
    ok: true,
    backendUrl: backendUrl,
    configured: configured
  };
}


/**
 * Ajuste seguro da V7.2.6. Pode ser executado mais de uma vez.
 * Não apaga pedidos e não recria a estrutura: apenas habilita a política final
 * de pagamentos e corrige o parcelamento padrão dos produtos atuais.
 */
function finalizarLojaDigital() {
  const ss = SpreadsheetApp.getActive();
  const cfgSh = ss.getSheetByName(CFG.SHEET_CONFIG);
  if (!cfgSh) throw new Error('A aba Configurações não foi encontrada.');

  upsertConfigValue_('Parcelamento cartão habilitado', 'SIM');
  upsertConfigValue_('Parcelamento máximo padrão', CFG.CARD_MAX_INSTALLMENTS);
  upsertConfigValue_('Parcela mínima cartão (R$)', CFG.CARD_MIN_INSTALLMENT_VALUE);
  upsertConfigValue_('Repassar taxa cartão ao comprador', 'SIM');

  const produtos = ss.getSheetByName(CFG.SHEET_PRODUTOS);
  if (produtos && produtos.getLastRow() > 1) {
    const count = produtos.getLastRow() - 1;
    const codes = produtos.getRange(2,1,count,1).getDisplayValues().flat();
    const parcelRange = produtos.getRange(2,21,count,1); // U
    const vals = parcelRange.getValues();
    vals.forEach((r,i) => {
      const code = String(codes[i] || '').trim().toUpperCase();
      if (!code) return;
      const atual = Number(r[0] || 0);
      if (['EST-POLO','SERV-POLO','SERV-CARECA'].includes(code) && atual <= 1) {
        r[0] = CFG.CARD_MAX_INSTALLMENTS;
      }
    });
    parcelRange.setValues(vals);
  }

  instalarTriggerSincronizacaoAsaas_();
  CacheService.getScriptCache().remove('CAMISAS_PUBLIC_BOOT_V722');
  try { ss.toast('Loja Digital finalizada: parcelamento e repasse de taxa do cartão habilitados.', '✅ Loja do Grêmio', 7); } catch (_) {}
  return { ok:true, versao:CFG.COMMERCE_VERSION, parcelasMaximas:CFG.CARD_MAX_INSTALLMENTS };
}

function upsertConfigValue_(campo, valor) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_CONFIG);
  if (!sh) return false;
  const last = Math.max(sh.getLastRow(), 1);
  const vals = sh.getRange(1,1,last,1).getDisplayValues().flat();
  const target = String(campo || '').trim().toUpperCase();
  const idx = vals.findIndex(v => String(v || '').trim().toUpperCase() === target);
  if (idx >= 0) sh.getRange(idx+1,2).setValue(valor);
  else sh.appendRow([campo, valor]);
  return true;
}

function configurarSistema() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('America/Fortaleza');

  const required = ['Dashboard', 'Pedidos', 'Produtos', 'Tamanhos', 'Lotes', 'Configurações', 'Log', 'Variantes', 'ItensPedido', 'Clientes', 'Pagamentos', 'Estoque', 'Cupons', 'Movimentações Financeiras', 'Entregas', 'Categorias', 'Campanhas', 'Eventos Asaas'];
  required.forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  validarCabecalhoPedidos_();
  aplicarFormatacaoPedidos_();
  aplicarValidacoesPedidos_();
  configurarCommerceCore_();
  finalizarLojaDigital();
  garantirSheetEventosAsaas_();
  instalarTriggerSincronizacaoAsaas_();

  // Não usa getUi() aqui, porque essa função pode ser executada pelo editor
  // do Apps Script, onde a interface da planilha nem sempre está disponível.
  try {
    const ss = SpreadsheetApp.getActive();
    if (ss) ss.toast(
      'Sistema validado. Integração Asaas preparada e sincronização automática instalada.',
      '✅ Loja do Grêmio',
      6
    );
  } catch (_) {
    // A configuração já foi concluída; apenas ignora a notificação visual.
  }

  return {
    ok: true,
    mensagem: 'Sistema validado, Commerce Core preparado e integração Asaas mantida.'
  };
}


/**
 * COMMERCE CORE V7
 * ----------------
 * Fundação genérica da Loja do Grêmio. O fluxo público da Nova Camisa continua
 * exatamente como antes; estas tabelas funcionam como camada de expansão.
 *
 * Objetivo: permitir que novos produtos (moletom, kits, canecas, eventos etc.)
 * sejam adicionados sem reconstruir pedidos, clientes, pagamentos e financeiro.
 */
function configurarCommerceCore() {
  configurarCommerceCore_();
  try {
    SpreadsheetApp.getActive().toast(
      'Estrutura futura da Loja preparada sem alterar as vendas da Nova Camisa.',
      '🛍️ Commerce Core V' + CFG.COMMERCE_VERSION,
      7
    );
  } catch (_) {}
  return { ok: true, versao: CFG.COMMERCE_VERSION };
}

function configurarCommerceCore_() {
  const ss = SpreadsheetApp.getActive();
  const defs = {
    'Variantes': ['Variante ID','Produto ID','SKU','Nome da variante','Tamanho','Cor','Modelo','Preço override','Custo override','Controla estoque','Ativo','Criado em','Atualizado em'],
    'ItensPedido': ['Pedido','Item ID','Produto ID','Variante ID','SKU','Produto','Quantidade','Valor unitário','Valor total','Custo unitário','Nome personalizado','Cargo/Área/Função','Lote','Criado em'],
    'Clientes': ['Cliente ID','Nome','WhatsApp','E-mail','Público','Turma/Setor','ID Asaas','Criado em','Última compra','Total de pedidos'],
    'Pagamentos': ['Pagamento ID','Pedido','Provedor','ID externo','Tipo/Meio','Valor bruto','Taxa','Valor líquido','Status','Link','Criado em','Pago em','Atualizado em'],
    'Estoque': ['Variante ID','SKU','Produto','Estoque físico','Reservado','Disponível','Estoque mínimo','Atualizado em','Observações'],
    'Cupons': ['Código','Tipo','Valor','Campanha','Produto ID','Início','Fim','Limite total','Usos','Limite por cliente','Ativo','Observações'],
    'Movimentações Financeiras': ['Referência','Data/Hora','Pedido','Tipo','Categoria','Descrição','Valor bruto','Taxa','Valor líquido','Custo','Resultado','Provedor','Meio de pagamento','Responsável'],
    'Entregas': ['Entrega ID','Pedido','Modalidade','Local','Status','Previsão','Entregue em','Recebido por','Responsável','Observações','Atualizado em'],
    'Categorias': ['Categoria ID','Nome','Slug','Descrição','Ordem','Ativo'],
    'Campanhas': ['Campanha ID','Nome','Slug','Início','Fim','Status','Descrição','Imagem ID','Ordem']
  };

  Object.keys(defs).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    ensureHeadersCommerce_(sh, defs[name]);
    formatarSheetCommerce_(sh, defs[name].length);
  });

  // HOTFIX V7.0.1: planilhas que já existiam podem carregar validações antigas
  // em colunas que agora têm outro significado (ex.: Campanhas!I = Ordem).
  // Normalizamos somente as abas pertencentes ao Commerce Core, sem apagar dados.
  normalizarValidacoesCommerceCore_(defs);

  ampliarProdutosParaLoja_();
  semearCommerceCore_();
  migrarProdutosAtuaisParaVariantes_();
}

function ensureHeadersCommerce_(sh, headers) {
  if (sh.getMaxColumns() < headers.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  }
  const current = sh.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  // Em tabelas novas, a linha 1 pode ser escrita integralmente. Se já houver dados,
  // só preenche cabeçalhos vazios para não destruir personalizações futuras.
  const emptySheet = sh.getLastRow() <= 1 && current.every(v => !String(v || '').trim());
  if (emptySheet) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const merged = current.map((v, i) => String(v || '').trim() || headers[i]);
    sh.getRange(1, 1, 1, headers.length).setValues([merged]);
  }
}

function formatarSheetCommerce_(sh, cols) {
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, cols)
    .setBackground('#063667')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setWrap(true);
}

/**
 * Remove validações herdadas/incompatíveis somente das tabelas do Commerce Core
 * e reaplica as regras corretas. Não apaga conteúdo nem formatação.
 *
 * Isso torna configurarSistema() idempotente: pode ser executado novamente mesmo
 * depois de uma configuração parcial ou de uma versão anterior ter deixado regras
 * de validação em colunas que mudaram de finalidade.
 */
function aplicarValidacaoCommerceColuna_(sh, column, validation) {
  if (!sh) return;
  const rows = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, column, rows, 1).setDataValidation(validation);
}

function normalizarValidacoesCommerceCore_(defs) {
  const ss = SpreadsheetApp.getActive();
  const simNao = SpreadsheetApp.newDataValidation()
    .requireValueInList(['SIM', 'NÃO'], true)
    .setAllowInvalid(false)
    .build();
  const statusCampanha = SpreadsheetApp.newDataValidation()
    .requireValueInList(['PLANEJADA', 'ATIVA', 'PAUSADA', 'ENCERRADA'], true)
    .setAllowInvalid(false)
    .build();

  Object.keys(defs).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const cols = defs[name].length;
    const rows = Math.max(sh.getMaxRows() - 1, 1);
    sh.getRange(2, 1, rows, cols).clearDataValidations();
  });

  const variantes = ss.getSheetByName(CFG.SHEET_VARIANTES);
  if (variantes) {
    aplicarValidacaoCommerceColuna_(variantes, 10, simNao); // Controla estoque
    aplicarValidacaoCommerceColuna_(variantes, 11, simNao); // Ativo
  }

  const cupons = ss.getSheetByName(CFG.SHEET_CUPONS);
  if (cupons) aplicarValidacaoCommerceColuna_(cupons, 11, simNao); // Ativo

  const categorias = ss.getSheetByName(CFG.SHEET_CATEGORIAS);
  if (categorias) aplicarValidacaoCommerceColuna_(categorias, 6, simNao); // Ativo

  const campanhas = ss.getSheetByName(CFG.SHEET_CAMPANHAS);
  if (campanhas) aplicarValidacaoCommerceColuna_(campanhas, 6, statusCampanha); // Status
}

function ampliarProdutosParaLoja_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PRODUTOS);
  if (!sh) return;

  const base = ['Código','Público','Nome','Personalização','Preço','Custo','Observações','Ativo'];
  const future = [
    'Categoria','Campanha','Tipo de venda','Descrição pública','Destaque','Ordem',
    'Imagem principal ID','Galeria IDs','Controla estoque','Permite cupom',
    'Preço Pix','Preço cartão','Parcelamento máximo','Criado em','Atualizado em'
  ];
  const headers = base.concat(future);

  if (sh.getMaxColumns() < headers.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  }

  // Mantém A:H (estrutura já usada pela Nova Camisa) e força somente os novos
  // cabeçalhos I:W. Isso corrige qualquer tentativa parcial da V7 anterior.
  const baseAtual = sh.getRange(1,1,1,base.length).getDisplayValues()[0];
  const baseMesclada = baseAtual.map((v,i) => String(v || '').trim() || base[i]);
  sh.getRange(1,1,1,base.length).setValues([baseMesclada]);
  sh.getRange(1,9,1,future.length).setValues([future]);
  sh.getRange(1,1,1,headers.length)
    .setBackground('#063667')
    .setFontColor('#fff')
    .setFontWeight('bold')
    .setWrap(true);

  // HOTFIX V7.0.2
  // V7.0/V7.0.1 podia deixar em Produtos!I:W validações herdadas de colunas
  // antigas (por exemplo SIM/NÃO em I2). Antes de escrever Categoria/Campanha,
  // removemos apenas as validações das NOVAS colunas. A:H não é tocado.
  const maxDataRows = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, 9, maxDataRows, future.length).clearDataValidations();

  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    const count = lastRow - 1;
    const codes = sh.getRange(2,1,count,1).getDisplayValues().flat();
    const rangeFuture = sh.getRange(2,9,count,future.length);
    const rows = rangeFuture.getValues();
    const agora = new Date();

    rows.forEach((r, idx) => {
      const code = String(codes[idx] || '').trim();
      if (!code) return;

      // Índices relativos à coluna I (0 = I, 14 = W).
      if (!r[0])  r[0]  = 'UNIFORMES';
      if (!r[1])  r[1]  = 'NOSSA ESCOLA, NOSSA IDENTIDADE';
      if (!r[2])  r[2]  = 'PRÉ-VENDA';
      if (!r[4])  r[4]  = code.toUpperCase() === 'EST-POLO' ? 'SIM' : 'NÃO';
      if (!r[5])  r[5]  = idx + 2;
      if (!r[8])  r[8]  = 'NÃO';
      if (!r[9])  r[9]  = 'SIM';
      if (!r[12]) r[12] = CFG.CARD_MAX_INSTALLMENTS;
      if (!r[13]) r[13] = agora;
      r[14] = agora;
    });

    rangeFuture.setValues(rows);
  }

  // Reaplica SOMENTE as validações corretas para as novas colunas booleanas.
  const simNao = SpreadsheetApp.newDataValidation()
    .requireValueInList(['SIM','NÃO'], true)
    .setAllowInvalid(false)
    .build();
  aplicarValidacaoCommerceColuna_(sh, 13, simNao); // Destaque
  aplicarValidacaoCommerceColuna_(sh, 17, simNao); // Controla estoque
  aplicarValidacaoCommerceColuna_(sh, 18, simNao); // Permite cupom
}

/**
 * Reparo manual de emergência para planilhas que passaram pela V7.0/V7.0.1.
 * Pode ser executado sozinho; não apaga pedidos nem dados A:H de Produtos.
 */
function repararValidacoesLoja() {
  const ss = SpreadsheetApp.getActive();
  const produtos = ss.getSheetByName(CFG.SHEET_PRODUTOS);
  if (produtos && produtos.getMaxColumns() >= 9) {
    const cols = Math.min(15, produtos.getMaxColumns() - 8);
    if (cols > 0) produtos.getRange(2, 9, Math.max(produtos.getMaxRows() - 1, 1), cols).clearDataValidations();
  }

  const campanhas = ss.getSheetByName(CFG.SHEET_CAMPANHAS);
  if (campanhas && campanhas.getMaxColumns() >= 9) {
    campanhas.getRange(2, 1, Math.max(campanhas.getMaxRows() - 1, 1), 9).clearDataValidations();
  }

  const categorias = ss.getSheetByName(CFG.SHEET_CATEGORIAS);
  if (categorias && categorias.getMaxColumns() >= 6) {
    categorias.getRange(2, 1, Math.max(categorias.getMaxRows() - 1, 1), 6).clearDataValidations();
  }

  configurarCommerceCore_();
  return { ok: true, versao: CFG.COMMERCE_VERSION, mensagem: 'Validações reparadas e Commerce Core reconfigurado.' };
}

function semearCommerceCore_() {
  const ss = SpreadsheetApp.getActive();
  const cat = ss.getSheetByName(CFG.SHEET_CATEGORIAS);
  if (cat && cat.getLastRow() < 2) {
    // setValues é usado de propósito: evita que uma validação antiga esquecida na
    // linha vazia bloqueie a semeadura inicial.
    cat.getRange(2, 1, 2, 6).clearDataValidations();
    cat.getRange(2, 1, 2, 6).setValues([
      ['CAT-UNIFORMES','Uniformes','uniformes','Camisas, moletons e peças de identidade escolar.',1,'SIM'],
      ['CAT-KITS','Kits e produtos','kits-e-produtos','Espaço preparado para futuros produtos e kits do Grêmio.',2,'SIM']
    ]);
    const simNao = SpreadsheetApp.newDataValidation().requireValueInList(['SIM','NÃO'], true).setAllowInvalid(false).build();
    aplicarValidacaoCommerceColuna_(cat, 6, simNao);
  }
  const camp = ss.getSheetByName(CFG.SHEET_CAMPANHAS);
  if (camp && camp.getLastRow() < 2) {
    // I = Ordem. Em uma execução anterior essa célula podia ter herdado a
    // validação SIM/NÃO, que gerava exatamente o erro em Campanhas!I2.
    camp.getRange(2, 1, 1, 9).clearDataValidations();
    camp.getRange(2, 1, 1, 9).setValues([[
      'CAMP-CAMISA-2026','Nossa Escola, Nossa Identidade','nossa-escola-nossa-identidade','','','ATIVA','Campanha da Nova Camisa Oficial.','',1
    ]]);
    const statusCampanha = SpreadsheetApp.newDataValidation().requireValueInList(['PLANEJADA','ATIVA','PAUSADA','ENCERRADA'], true).setAllowInvalid(false).build();
    aplicarValidacaoCommerceColuna_(camp, 6, statusCampanha);
  }
}

function migrarProdutosAtuaisParaVariantes_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_VARIANTES);
  if (!sh) return;
  const produtos = getProdutosAtivos_();
  const tamanhos = getTamanhosAtivos_();
  const existing = sh.getLastRow() > 1
    ? new Set(sh.getRange(2,3,sh.getLastRow()-1,1).getDisplayValues().flat().map(v => String(v||'').trim().toUpperCase()))
    : new Set();

  produtos.forEach(p => tamanhos.forEach(t => {
    const sku = `${p.codigo}-${t}`.toUpperCase();
    if (existing.has(sku)) return;
    const id = 'VAR-' + sku.replace(/[^A-Z0-9]+/g,'-');
    sh.appendRow([id,p.codigo,sku,t,t,'',p.nome,p.preco,p.custo,'NÃO','SIM',new Date(),new Date()]);
    existing.add(sku);
  }));
}

function registrarPedidoCommerceCore_(d) {
  const ss=SpreadsheetApp.getActive();
  const itens=ss.getSheetByName(CFG.SHEET_ITENS_PEDIDO),clientes=ss.getSheetByName(CFG.SHEET_CLIENTES),pagamentos=ss.getSheetByName(CFG.SHEET_PAGAMENTOS);
  if(!itens || !clientes || !pagamentos) return false;
  const clienteId=upsertClienteCommerce_(d);
  const sku=`${d.codigoProduto}-${d.tamanho}`.toUpperCase();
  const varianteId='VAR-'+sku.replace(/[^A-Z0-9]+/g,'-');
  const itemId=d.pedidoId+'-01';
  if(!findRowByValue_(itens,2,itemId)) itens.appendRow([d.pedidoId,itemId,d.codigoProduto,varianteId,sku,d.produtoNome,Number(d.quantidade||1),Number(d.valorUnitario||0),Number(d.valorUnitario||0)*Number(d.quantidade||1),Number(d.custoUnitario||0),d.nomeCamisa||'',d.cargoFuncao||'',d.lote||'',d.data||new Date()]);
  if(!findRowByValue_(pagamentos,2,d.pedidoId)) pagamentos.appendRow(['PAY-'+d.pedidoId,d.pedidoId,'ASAAS','',d.metodoPagamento||'PIX',Number(d.valorCobranca||d.valorUnitario||0),0,Number(d.valorUnitario||0),'CRIANDO COBRANÇA','',new Date(),' ',new Date()]);
  return {clienteId,itemId};
}


function upsertClienteCommerce_(d) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_CLIENTES);
  if (!sh) return '';
  const phone = somenteDigitos_(d.whatsapp);
  let row = 0;
  if (sh.getLastRow() > 1 && phone) {
    const vals = sh.getRange(2,3,sh.getLastRow()-1,1).getDisplayValues().flat();
    const idx = vals.findIndex(v => somenteDigitos_(v) === phone);
    if (idx >= 0) row = idx + 2;
  }
  const id = row ? String(sh.getRange(row,1).getValue()||'') : 'CLI-' + Utilities.getUuid().slice(0,8).toUpperCase();
  if (!row) {
    sh.appendRow([id,d.nome||'',phone,d.email||'',d.publico||'',d.turmaSetor||'','',new Date(),d.data||new Date(),1]);
  } else {
    const count = Number(sh.getRange(row,10).getValue()||0) + 1;
    sh.getRange(row,2,1,9).setValues([[
      d.nome||sh.getRange(row,2).getValue(), phone, d.email||sh.getRange(row,4).getValue(),
      d.publico||sh.getRange(row,5).getValue(), d.turmaSetor||sh.getRange(row,6).getValue(),
      sh.getRange(row,7).getValue(), sh.getRange(row,8).getValue()||new Date(), d.data||new Date(), count
    ]]);
  }
  return id;
}

function atualizarClienteCommerce_(pedidoId, asaasCustomerId) {
  const ss = SpreadsheetApp.getActive();
  const pedidos = ss.getSheetByName(CFG.SHEET_PEDIDOS);
  const clientes = ss.getSheetByName(CFG.SHEET_CLIENTES);
  if (!pedidos || !clientes) return;
  const pr = findRowByValue_(pedidos,1,pedidoId); if (!pr) return;
  const phone = somenteDigitos_(pedidos.getRange(pr,6).getValue()); if (!phone) return;
  const vals = clientes.getLastRow()>1 ? clientes.getRange(2,3,clientes.getLastRow()-1,1).getDisplayValues().flat() : [];
  const idx = vals.findIndex(v => somenteDigitos_(v)===phone);
  if (idx>=0) clientes.getRange(idx+2,7).setValue(asaasCustomerId||'');
}

function atualizarPagamentoCommerce_(pedidoId, p) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PAGAMENTOS);
  if (!sh) return false;
  let row = findRowByValue_(sh,2,pedidoId);
  if (!row) {
    sh.appendRow(['PAY-'+pedidoId,pedidoId,p.provedor||'ASAAS',p.idExterno||'',p.tipo||'UNDEFINED',Number(p.valor||0),Number(p.taxa||0),Number(p.liquido||0),p.status||'',p.link||'',new Date(),p.pagoEm||'',new Date()]);
    return true;
  }
  const created = sh.getRange(row,11).getValue() || new Date();
  sh.getRange(row,1,1,13).setValues([[
    sh.getRange(row,1).getValue() || 'PAY-'+pedidoId,
    pedidoId,p.provedor||'ASAAS',p.idExterno||sh.getRange(row,4).getValue(),p.tipo||sh.getRange(row,5).getValue()||'UNDEFINED',
    Number(p.valor||0),Number(p.taxa||0),Number(p.liquido||0),p.status||'',p.link||'',created,p.pagoEm||sh.getRange(row,12).getValue()||'',new Date()
  ]]);
  return true;
}

function registrarMovimentacaoFinanceiraCommerce_(m) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_FINANCEIRO);
  if (!sh || !m.referencia) return false;
  if (findRowByValue_(sh,1,m.referencia)) return false; // idempotência
  sh.appendRow([
    m.referencia,new Date(),m.pedidoId||'',m.tipo||'VENDA',m.categoria||'RECEITA DE VENDAS',m.descricao||'',
    Number(m.valorBruto||0),Number(m.taxa||0),Number(m.valorLiquido||0),Number(m.custo||0),Number(m.resultado||0),
    m.provedor||'',m.meioPagamento||'',m.responsavel||'SISTEMA'
  ]);
  return true;
}

function findRowByValue_(sh, col, value) {
  if (!sh || sh.getLastRow() < 2) return 0;
  const target = String(value==null?'':value).trim().toUpperCase();
  if (!target) return 0;
  const vals = sh.getRange(2,col,sh.getLastRow()-1,1).getDisplayValues().flat();
  const idx = vals.findIndex(v => String(v||'').trim().toUpperCase() === target);
  return idx >= 0 ? idx + 2 : 0;
}

function getCommerceCapabilities_() {
  return {
    versao: CFG.COMMERCE_VERSION,
    atual: 'Nova Camisa Oficial',
    preparadoPara: ['novos produtos','variantes','estoque','cupons','clientes','múltiplos pagamentos','entregas','financeiro','campanhas'],
    observacao: 'Recursos futuros são ativados gradualmente sem alterar o fluxo estável da camisa.'
  };
}

function abrirDashboard() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);
}

function mostrarLinksWebApp() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert(
      'Web App ainda não implantado',
      'Vá em Implantar → Nova implantação → Aplicativo da Web. Depois execute esta opção novamente.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const msg =
    'ESTUDANTES:\n' + url + '?p=estudantes\n\n' +
    'SERVIDORES:\n' + url + '?p=servidores\n\n' +
    'CONHEÇA OS MODELOS:\n' + url + '?p=modelos\n\n' +
    'ACOMPANHAR PEDIDO:\n' + url + '?p=acompanhar\n\n' +
    'ÁREA DO GRÊMIO:\n' + url + '?p=admin';

  SpreadsheetApp.getUi().alert('Links do sistema', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function criarPedido(payload) {
  const data = normalizarPayload_(payload);
  validarPedido_(data);
  const produto = getProduto_(data.codigoProduto);
  if (!produto) throw new Error('Produto indisponível.');
  if (data.publico === 'ESTUDANTE' && data.codigoProduto !== 'EST-POLO') throw new Error('Estudantes podem encomendar somente a camisa polo oficial.');
  if (data.publico === 'SERVIDOR' && !['SERV-POLO','SERV-CARECA'].includes(data.codigoProduto)) throw new Error('Modelo de servidor inválido.');
  if (!getTamanhosAtivos_().includes(data.tamanho)) throw new Error('Tamanho inválido ou indisponível.');

  const cfg = getConfigMap_();
  const vendasStatus = data.publico === 'ESTUDANTE' ? String(cfg['Vendas estudantes'] || 'ABERTAS').toUpperCase() : String(cfg['Vendas servidores'] || 'ABERTAS').toUpperCase();
  if (vendasStatus !== 'ABERTAS') throw new Error('As encomendas para este público estão temporariamente encerradas.');
  const prefixo = data.publico === 'ESTUDANTE' ? (cfg['Prefixo pedido estudante'] || 'PA-A') : (cfg['Prefixo pedido servidor'] || 'PA-S');
  const lote = data.publico === 'ESTUDANTE' ? (cfg['Lote atual estudantes'] || 'LOTE-EST-01') : (cfg['Lote atual servidores'] || 'LOTE-SERV-01');
  const asaas = getAsaasConfig_();
  if (!asaas.configured) throw new Error('A integração com o Asaas ainda não está configurada nas Propriedades do script.');
  if (asaas.env !== 'PRODUCTION') throw new Error('Os pagamentos oficiais ainda não foram ativados. Tente novamente em instantes.');

  const pedidoId = nextOrderId_(prefixo);
  const valorBase = roundMoney_(data.metodoPagamento === 'PIX' ? Number(produto.precoPix || produto.preco) : data.metodoPagamento === 'CREDIT_CARD' ? Number(produto.precoCartao || produto.preco) : Number(produto.preco));
  let valorCobrado = valorBase;
  let acrescimoPagamento = 0;
  let valorParcela = valorBase;
  let planoCartao = null;
  if (data.metodoPagamento === 'CREDIT_CARD') {
    planoCartao = calcularPlanoCartaoComRepasse_(valorBase, data.parcelas, true);
    valorCobrado = planoCartao.valorTotal;
    acrescimoPagamento = planoCartao.acrescimo;
    valorParcela = planoCartao.valorParcela;
  } else {
    data.parcelas = 1;
  }

  const taxaEstimada = Math.max(0, roundMoney_(valorCobrado - valorBase));
  const custo = Number(produto.custo);
  const liquidoEstimado = roundMoney_(valorCobrado - taxaEstimada);
  const margemEstimada = roundMoney_(liquidoEstimado - custo);
  const now = new Date();

  const row = [
    pedidoId,now,data.publico,data.nome,data.turmaSetor,data.whatsapp,data.email,
    data.codigoProduto,produto.nome,data.tamanho,data.nomeCamisa,data.cargoFuncao,
    lote,valorCobrado,taxaEstimada,liquidoEstimado,custo,margemEstimada,
    'CRIANDO COBRANÇA','NÃO ENVIADO','NÃO ENTREGUE','','','',data.observacoes,
    '',now,'SISTEMA','','','','',data.metodoPagamento,data.parcelas,valorBase,acrescimoPagamento,valorParcela,''
  ];
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  sh.appendRow(row);
  const rowIndex = sh.getLastRow();

  try {
    registrarPedidoCommerceCore_({
      pedidoId,data:now,publico:data.publico,nome:data.nome,turmaSetor:data.turmaSetor,whatsapp:data.whatsapp,email:data.email,
      codigoProduto:data.codigoProduto,produtoNome:produto.nome,tamanho:data.tamanho,quantidade:1,valorUnitario:valorBase,
      valorCobranca:valorCobrado,custoUnitario:custo,lote,nomeCamisa:data.nomeCamisa,cargoFuncao:data.cargoFuncao,
      metodoPagamento:data.metodoPagamento,parcelas:data.parcelas
    });
  } catch (commerceErr) { registrarLog_('AVISO COMMERCE CORE',pedidoId,'SISTEMA',commerceErr.message); }
  registrarLog_('PEDIDO CRIADO',pedidoId,'WEB APP',`${data.publico} • ${produto.nome} • ${data.tamanho} • ${data.metodoPagamento}${data.parcelas>1?' '+data.parcelas+'x':''}`);

  try {
    if (asaas.env === 'SANDBOX' && data.metodoPagamento === 'PIX') {
      try { const k=garantirChavePixSandbox_(); if(k && k.created) Utilities.sleep(800); } catch(keyErr) { registrarLog_('AVISO CHAVE PIX',pedidoId,'ASAAS',keyErr.message); }
    }
    const cpfCnpj = asaas.env === 'SANDBOX' ? '24971563792' : data.cpf;
    const notificacoesAtivas = asaas.env === 'PRODUCTION' ? sim_(cfg['Notificações Asaas ao cliente']) : sim_(cfg['E-mails de teste no Sandbox']);
    const cliente = obterOuCriarClienteAsaas_({nome:data.nome,cpfCnpj,whatsapp:data.whatsapp,email:data.email,pedidoId,notificacoesAtivas,ambiente:asaas.env});
    sh.getRange(rowIndex,31).setValue(cliente.id || '');
    try { atualizarClienteCommerce_(pedidoId,cliente.id || ''); } catch (_) {}
    if (notificacoesAtivas) { try { configurarNotificacoesClienteAsaas_(cliente.id,cfg); } catch(notifErr){ registrarLog_('AVISO NOTIFICAÇÕES ASAAS',pedidoId,'ASAAS',notifErr.message); } }

    const diasValidade = Math.max(0,Number(cfg['Validade cobrança (dias)'] || cfg['Validade cobrança Pix (dias)'] || 2));
    // No Sandbox não enviamos callback: o Asaas valida o domínio e isso provocava
    // uma tentativa com erro + retry em TODO checkout, adicionando vários segundos.
    // Em Produção o callback volta automaticamente, usando a URL pública cadastrada.
    const trackingUrl = asaas.env === 'PRODUCTION' ? buildTrackingUrl_(pedidoId,data.whatsapp) : '';
    const cobranca = criarOuRecuperarCobrancaAsaas_({
      pedidoId,customerId:cliente.id,billingType:data.metodoPagamento,value:valorCobrado,parcelas:data.parcelas,
      dueDate:datePlusDays_(diasValidade),description:`Loja Digital do Grêmio • ${pedidoId} • ${produto.nome} • ${data.tamanho}`,callbackUrl:trackingUrl
    });

    let pix=null,pixError='';
    if (data.metodoPagamento === 'PIX') {
      try { pix=asaasRequest_(`/payments/${encodeURIComponent(cobranca.id)}/pixQrCode`,'get'); }
      catch(pixErr){ pixError=pixErr.message || 'Não foi possível obter o QR Code Pix.'; registrarLog_('AVISO PIX',pedidoId,'ASAAS',pixError); }
    }
    const paymentUrl = cobranca.invoiceUrl || '';
    sh.getRange(rowIndex,19).setValue('AGUARDANDO PAGAMENTO');
    sh.getRange(rowIndex,22).setValue(cobranca.id || '');
    sh.getRange(rowIndex,23).setValue(paymentUrl);
    sh.getRange(rowIndex,27).setValue(new Date());
    sh.getRange(rowIndex,28).setValue('SISTEMA');
    sh.getRange(rowIndex,38).setValue(cobranca.installment || '');

    try {
      atualizarPagamentoCommerce_(pedidoId,{provedor:'ASAAS',idExterno:cobranca.id || '',tipo:data.metodoPagamento,valor:valorCobrado,taxa:taxaEstimada,liquido:liquidoEstimado,status:'AGUARDANDO PAGAMENTO',link:paymentUrl});
    } catch(commerceErr){ registrarLog_('AVISO PAGAMENTO CORE',pedidoId,'SISTEMA',commerceErr.message); }
    registrarLog_('COBRANÇA ASAAS CRIADA',pedidoId,asaas.env,`${cobranca.id || ''} • ${data.metodoPagamento}${data.parcelas>1?' '+data.parcelas+'x':''}`);
    try { enviarEmailPedidoCriado_(rowIndex); } catch(mailErr){ registrarLog_('AVISO E-MAIL',pedidoId,'SISTEMA',mailErr.message); }

    return {
      ok:true,pedido:pedidoId,publico:data.publico,produto:produto.nome,tamanho:data.tamanho,
      valor:valorCobrado,valorBase,valorCobrado,acrescimoPagamento,valorParcela,
      metodoPagamento:data.metodoPagamento,parcelas:data.parcelas,status:'AGUARDANDO PAGAMENTO',paymentUrl,
      pix:pix ? {encodedImage:pix.encodedImage || '',payload:pix.payload || '',expirationDate:pix.expirationDate || ''} : null,
      pixError,asaasEnv:asaas.env
    };
  } catch(err) {
    sh.getRange(rowIndex,19).setValue('ERRO NA COBRANÇA');
    sh.getRange(rowIndex,27).setValue(new Date());
    const obsAtual=String(sh.getRange(rowIndex,25).getValue() || '');
    sh.getRange(rowIndex,25).setValue([obsAtual,`ERRO ASAAS: ${limparTexto_(err.message,220)}`].filter(Boolean).join(' | '));
    registrarLog_('ERRO ASAAS',pedidoId,'API',err.message);
    throw new Error(`O pedido ${pedidoId} foi registrado, mas a cobrança não pôde ser criada. Avise o Grêmio informando esse número.`);
  }
}

function consultarPedido(pedido, whatsapp, incluirPix) {
  const id=String(pedido || '').trim().toUpperCase();
  const phone=somenteDigitos_(whatsapp);
  if(!id || phone.length<8) throw new Error('Informe o número do pedido e o WhatsApp usado na compra.');
  const sh=SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  const values=sh.getDataRange().getValues();
  for(let i=1;i<values.length;i++){
    if(String(values[i][0]).trim().toUpperCase()!==id) continue;
    if(somenteDigitos_(values[i][5])!==phone) throw new Error('Os dados informados não conferem.');
    const rowIndex=i+1;
    let paymentInfo=null;
    try{ if(values[i][21]) paymentInfo=sincronizarPagamentoLinha_(rowIndex); }catch(err){ registrarLog_('AVISO CONSULTA ASAAS',id,'API',err.message); }
    const row=sh.getRange(rowIndex,1,1,38).getValues()[0];
    let pix=null,pixError='';
    const statusPagamento=String(row[18] || '').toUpperCase();
    const metodoEscolhido=String(row[32] || (paymentInfo && paymentInfo.billingType) || '').toUpperCase();
    if(Boolean(incluirPix) && metodoEscolhido==='PIX' && row[21] && !['PAGO','CONFIRMADO','ESTORNADO','CANCELADO'].includes(statusPagamento)){
      try{ const qr=asaasRequest_(`/payments/${encodeURIComponent(row[21])}/pixQrCode`,'get'); pix={encodedImage:qr.encodedImage || '',payload:qr.payload || '',expirationDate:qr.expirationDate || ''}; }
      catch(err){ pixError=err.message || 'Não foi possível recuperar o QR Code Pix.'; }
    }
    return {
      pedido:row[0],data:formatDate_(row[1]),publico:row[2],nome:row[3],modelo:row[8],tamanho:row[9],lote:row[12],
      valor:Number(row[13] || 0),valorBase:Number(row[34] || row[13] || 0),acrescimoPagamento:Number(row[35] || 0),
      parcelas:Number(row[33] || 1),valorParcela:Number(row[36] || row[13] || 0),
      statusPagamento:row[18] || '',statusProducao:row[19] || '',statusEntrega:row[20] || '',
      meioPagamento:humanizarMeioPagamentoAsaas_(metodoEscolhido || (paymentInfo && paymentInfo.billingType) || ''),
      linkPagamento:row[22] || '',dataPagamento:formatDate_(row[23]),prazoPrevisto:formatDateOnly_(row[25]),ultimaAtualizacao:formatDate_(row[26]),
      mensagemCliente:row[28] || '',localRetirada:String(getConfigMap_()['Local padrão de retirada'] || ''),pix,pixError
    };
  }
  throw new Error('Pedido não encontrado.');
}


function roundMoney_(v) { return Math.round((Number(v || 0) + Number.EPSILON) * 100) / 100; }
function ceilMoney_(v) { return Math.ceil((Number(v || 0) - 1e-9) * 100) / 100; }

function getPaymentPolicy_() {
  const cfg = getConfigMap_();
  const maxCfg = Number(cfg['Parcelamento máximo padrão'] || CFG.CARD_MAX_INSTALLMENTS);
  const minCfg = Number(cfg['Parcela mínima cartão (R$)'] || CFG.CARD_MIN_INSTALLMENT_VALUE);
  return {
    cartaoHabilitado: sim_(cfg['Parcelamento cartão habilitado'] || 'SIM'),
    maxParcelas: Math.max(1, Math.min(CFG.CARD_MAX_INSTALLMENTS, Math.floor(maxCfg || CFG.CARD_MAX_INSTALLMENTS))),
    parcelaMinima: Math.max(5, roundMoney_(minCfg || CFG.CARD_MIN_INSTALLMENT_VALUE)),
    repassarTaxaCartao: sim_(cfg['Repassar taxa cartão ao comprador'] || 'SIM')
  };
}

function extrairSimulacaoCartao_(resp, valorInformado) {
  const r = resp || {};
  const cc = r.creditCard || r.credit_card || r.creditcard || r.CREDIT_CARD || r.card || {};
  const inst = cc.installment || r.installment || {};
  const num = (...vals) => {
    for (const v of vals) { const n = Number(v); if (Number.isFinite(n)) return n; }
    return 0;
  };
  return {
    netValue: num(cc.netValue, inst.paymentNetValue, r.netValue),
    feePercentage: num(cc.feePercentage, r.feePercentage),
    operationFee: num(cc.operationFee, r.operationFee),
    paymentValue: num(inst.paymentValue, r.paymentValue, valorInformado)
  };
}

function simularCartaoAsaas_(valor, parcelas) {
  const body = {
    value: roundMoney_(valor),
    installmentCount: Math.max(1, Math.floor(Number(parcelas || 1))),
    billingTypes: ['CREDIT_CARD']
  };
  const raw = asaasRequest_('/payments/simulate', 'post', body);
  const parsed = extrairSimulacaoCartao_(raw, body.value);
  if (!(parsed.netValue > 0)) throw new Error('O Asaas não retornou o valor líquido da simulação do cartão.');
  return Object.assign({ raw:raw }, parsed);
}

function calcularPlanoCartaoComRepasse_(valorBase, parcelas, verificar) {
  const base = roundMoney_(valorBase);
  const n = Math.max(1, Math.floor(Number(parcelas || 1)));
  const policy = getPaymentPolicy_();
  if (!policy.cartaoHabilitado) throw new Error('Pagamento por cartão está temporariamente desabilitado.');
  if (n > policy.maxParcelas) throw new Error('Quantidade de parcelas acima do limite configurado.');

  const initial = simularCartaoAsaas_(base, n);
  let total = base;
  if (policy.repassarTaxaCartao) {
    const pct = Number(initial.feePercentage || 0) / 100;
    const fixed = Number(initial.operationFee || 0);
    if (pct > 0 && pct < 0.5) total = ceilMoney_((base + fixed) / (1 - pct));
    else if (initial.netValue > 0) total = ceilMoney_(base * (base / initial.netValue));
  }

  let checked = initial;
  if (verificar && total !== base) {
    checked = simularCartaoAsaas_(total, n);
    if (policy.repassarTaxaCartao && checked.netValue + 0.005 < base) {
      const pct2 = Number(checked.feePercentage || 0) / 100;
      const ajuste = base - checked.netValue;
      total = ceilMoney_(total + (pct2 > 0 && pct2 < 0.5 ? ajuste / (1-pct2) : ajuste));
      checked = simularCartaoAsaas_(total, n);
    }
  }

  const parcela = roundMoney_(total / n);
  return {
    parcelas: n,
    valorBase: base,
    valorTotal: total,
    valorParcela: parcela,
    acrescimo: roundMoney_(total - base),
    liquidoEstimado: roundMoney_(checked.netValue || base),
    taxaPercentual: Number(checked.feePercentage || initial.feePercentage || 0),
    taxaOperacao: Number(checked.operationFee || initial.operationFee || 0)
  };
}

function simularOpcoesCartao(codigoProduto) {
  const produto = getProduto_(codigoProduto);
  if (!produto) throw new Error('Produto não encontrado.');
  const asaas = getAsaasConfig_();
  if (!asaas.configured) throw new Error('A integração com o Asaas ainda não está configurada.');
  if (asaas.env !== 'PRODUCTION') throw new Error('O pagamento oficial por cartão ainda não foi ativado.');
  const policy = getPaymentPolicy_();
  if (!policy.cartaoHabilitado) throw new Error('Pagamento por cartão está temporariamente desabilitado.');
  const base = roundMoney_(Number(produto.precoCartao || produto.preco || 0));
  if (!(base > 0)) throw new Error('Preço do produto inválido.');
  const byMinimum = Math.max(1, Math.floor(base / policy.parcelaMinima));
  const max = Math.max(1, Math.min(CFG.CARD_MAX_INSTALLMENTS, policy.maxParcelas, Number(produto.parcelamentoMaximo || policy.maxParcelas), byMinimum));

  const cache = CacheService.getScriptCache();
  const cacheKey = ['CARD_SIM_V721', asaas.env, produto.codigo, base.toFixed(2), max, policy.repassarTaxaCartao ? 'R' : 'N'].join('_');
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (_) {}
  }

  const opcoes = [];
  let ultimoErro = '';
  for (let n=1;n<=max;n++) {
    try { opcoes.push(calcularPlanoCartaoComRepasse_(base, n, false)); }
    catch (err) { ultimoErro = err.message || String(err); }
  }
  if (!opcoes.length) throw new Error(ultimoErro || 'Não foi possível simular o cartão agora.');
  const result = { produto:produto.codigo, nome:produto.nome, ambiente:asaas.env, valorBase:base, opcoes:opcoes };
  try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (_) {}
  return result;
}

function testarParcelamentoCartaoSelecionado() {
  const sh = SpreadsheetApp.getActiveSheet();
  let code = 'EST-POLO';
  if (sh && sh.getName() === CFG.SHEET_PRODUTOS && sh.getActiveRange().getRow() >= 2) {
    code = String(sh.getRange(sh.getActiveRange().getRow(),1).getValue() || code);
  }
  try {
    const r = simularOpcoesCartao(code);
    const lines = r.opcoes.map(o => `${o.parcelas}x de R$ ${o.valorParcela.toFixed(2)} • total R$ ${o.valorTotal.toFixed(2)} • acréscimo R$ ${o.acrescimo.toFixed(2)}`);
    notificar_('💳 Parcelamento Asaas • ' + code, lines.join('\n'));
    return r;
  } catch (err) {
    notificar_('❌ Erro no parcelamento', err.message || String(err));
    return {ok:false,mensagem:err.message || String(err)};
  }
}

function getPublicData_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('CAMISAS_PUBLIC_BOOT_V722');
  if (cached) { try { return JSON.parse(cached); } catch (_) {} }
  const cfg = getConfigMap_();
  const produtos = getProdutosAtivos_();
  const tamanhos = getTamanhosAtivos_();
  const asaas = getAsaasConfig_();
  const legacyLogoId = cfg['ID do arquivo da logo no Google Drive'] || '';
  const paymentCfg = getPaymentPolicy_();
  const result = {
    escola: cfg['Escola'] || 'EEMTI Professor Paulo Ayrton de Araújo',
    organizacao: cfg['Organização'] || 'Grêmio Estudantil',
    projeto: cfg['Projeto'] || 'Nossa Escola, Nossa Identidade',
    instagram: cfg['Instagram'] || '@gremiopaoficial',
    whatsappGremio: cfg['WhatsApp do Grêmio'] || '',
    emailGremio: cfg['E-mail / Drive do Grêmio'] || 'gremioestudantilpaa@gmail.com',
    colors: { navy: cfg['Cor azul-marinho'] || '#063667', green: cfg['Cor verde'] || '#67B82E', orange: cfg['Cor laranja'] || '#F36F00' },
    logos: {
      escolaId: String(cfg['ID da logo da escola'] || legacyLogoId || ''),
      gremioId: String(cfg['ID da logo do Grêmio'] || ''),
      projetoId: String(cfg['ID da logo do projeto'] || '')
    },
    mockups: (() => {
      const ids = getMockupIds_(cfg);
      return {
        estudanteId: ids.estudante,
        servidorPoloId: ids.servidorPolo,
        servidorCarecaId: ids.servidorCareca
      };
    })(),
    asaasEnv: asaas.env,
    asaasConfigured: asaas.configured,
    meiosPagamento: ['PIX','BOLETO','CREDIT_CARD'],
    pagamento: {
      pix: true, boleto: true, cartao: true, parcelamento: paymentCfg.cartaoHabilitado,
      parcelasMaximas: paymentCfg.maxParcelas,
      parcelaMinima: paymentCfg.parcelaMinima,
      repasseTaxaCartao: paymentCfg.repassarTaxaCartao,
      cartaoProcessadoNoAsaas: true
    },
    commerce: {
      versao: CFG.COMMERCE_VERSION, modoAtual: 'NOVA_CAMISA', lojaPreparada: true,
      recursosPreparados: ['CATALOGO','VARIANTES','CARRINHO_FUTURO','LOTES','ESTOQUE','CUPONS','CLIENTES','PAGAMENTOS','ENTREGAS','FINANCEIRO']
    },
    baseUrl: ScriptApp.getService().getUrl() || '',
    vendas: {
      estudantes: String(cfg['Vendas estudantes'] || 'ABERTAS').toUpperCase(),
      servidores: String(cfg['Vendas servidores'] || 'ABERTAS').toUpperCase()
    },
    adminConfigured: Boolean(PropertiesService.getScriptProperties().getProperty('ADMIN_CODE_HASH') || PropertiesService.getScriptProperties().getProperty('ADMIN_ACCESS_CODE')),
    produtos, tamanhos
  };
  try { cache.put('CAMISAS_PUBLIC_BOOT_V722', JSON.stringify(result), 30); } catch (_) {}
  return result;
}


function getLogosPublicas() {
  const cfg = getConfigMap_();
  const legacyLogoId = cfg['ID do arquivo da logo no Google Drive'] || '';
  return {
    escola: getLogoDataUri_(cfg['ID da logo da escola'] || legacyLogoId),
    gremio: getLogoDataUri_(cfg['ID da logo do Grêmio']),
    projeto: getLogoDataUri_(cfg['ID da logo do projeto'])
  };
}


/**
 * Imagens públicas da página "Conheça os Modelos".
 * Os IDs oficiais são gerenciados pela Área do Grêmio e salvos na aba
 * Configurações. O bloco MOCKUP_IDS permanece apenas como fallback legado.
 */
function getMockupIds_(cfg) {
  cfg = cfg || getConfigMap_();
  return {
    estudante: String(cfg[MOCKUP_CONFIG_KEYS.estudante] || MOCKUP_IDS.estudante || '').trim(),
    servidorPolo: String(cfg[MOCKUP_CONFIG_KEYS.servidorPolo] || MOCKUP_IDS.servidorPolo || '').trim(),
    servidorCareca: String(cfg[MOCKUP_CONFIG_KEYS.servidorCareca] || MOCKUP_IDS.servidorCareca || '').trim()
  };
}

function getMockupsPublicos() {
  const ids = getMockupIds_();
  return {
    estudante: getLogoDataUri_(ids.estudante),
    servidorPolo: getLogoDataUri_(ids.servidorPolo),
    servidorCareca: getLogoDataUri_(ids.servidorCareca)
  };
}

function extrairDriveFileId_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  if (/^[A-Za-z0-9_-]{20,}$/.test(raw)) return raw;

  const patterns = [
    /\/d\/([A-Za-z0-9_-]{20,})/i,
    /[?&]id=([A-Za-z0-9_-]{20,})/i,
    /\/file\/d\/([A-Za-z0-9_-]{20,})/i
  ];
  for (const re of patterns) {
    const m = raw.match(re);
    if (m) return m[1];
  }
  throw new Error('Use o ID do arquivo ou um link válido do Google Drive.');
}

function validarMockupDriveId_(value, label) {
  const id = extrairDriveFileId_(value);
  if (!id) return '';
  try {
    const file = DriveApp.getFileById(id);
    const mime = String(file.getMimeType() || '').toLowerCase();
    if (!mime.startsWith('image/')) {
      throw new Error('O arquivo não é uma imagem.');
    }
    return id;
  } catch (err) {
    throw new Error(`${label}: não foi possível acessar essa imagem no Google Drive. Verifique o link/ID e se a conta do sistema tem acesso ao arquivo.`);
  }
}

function getConfigMap_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_CONFIG);
  if (!sh) return {};
  const last = Math.max(sh.getLastRow(), 1);
  const rows = sh.getRange(1, 1, last, 2).getValues();
  const obj = {};
  rows.forEach(r => {
    const key = String(r[0] || '').trim();
    if (key && key !== 'Campo') obj[key] = r[1];
  });
  return obj;
}

function getProdutosAtivos_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PRODUTOS);
  const values = sh.getDataRange().getValues();
  const out = [];
  const policy = getPaymentPolicy_();
  for (let i=1;i<values.length;i++) {
    if (!values[i][0]) continue;
    if (String(values[i][7]).trim().toUpperCase() !== 'SIM') continue;
    const explicitMax = Number(values[i][20] || 0);
    out.push({
      codigo: String(values[i][0]).trim(),
      publico: String(values[i][1]).trim().toUpperCase(),
      nome: String(values[i][2]).trim(),
      personalizacao: String(values[i][3]).trim().toUpperCase(),
      preco: Number(values[i][4] || 0),
      custo: Number(values[i][5] || 0),
      precoPix: Number(values[i][18] || values[i][4] || 0),
      precoCartao: Number(values[i][19] || values[i][4] || 0),
      parcelamentoMaximo: Math.max(1, Math.min(CFG.CARD_MAX_INSTALLMENTS, explicitMax > 0 ? explicitMax : policy.maxParcelas))
    });
  }
  return out;
}


function getProduto_(codigo) {
  const code = String(codigo || '').trim().toUpperCase();
  return getProdutosAtivos_().find(p => p.codigo.toUpperCase() === code) || null;
}

function getTamanhosAtivos_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_TAMANHOS);
  const values = sh.getDataRange().getValues();
  return values.slice(1)
    .filter(r => r[1] && String(r[2]).trim().toUpperCase() === 'SIM')
    .sort((a, b) => Number(a[0] || 999) - Number(b[0] || 999))
    .map(r => String(r[1]).trim().toUpperCase());
}

function nextOrderId_(prefixo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const props = PropertiesService.getScriptProperties();
    const key = 'COUNTER_' + prefixo.replace(/[^A-Z0-9]/gi, '_').toUpperCase();
    let n = Number(props.getProperty(key) || 0);

    if (!n) {
      const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
      const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().flat() : [];
      ids.forEach(id => {
        const m = String(id).match(new RegExp('^' + escapeRegExp_(prefixo) + '-(\\d+)$', 'i'));
        if (m) n = Math.max(n, Number(m[1]));
      });
    }

    n += 1;
    props.setProperty(key, String(n));
    return `${prefixo}-${String(n).padStart(4, '0')}`;
  } finally {
    lock.releaseLock();
  }
}

function normalizarPayload_(p) {
  p = p || {};
  const metodo = String(p.metodoPagamento || 'PIX').trim().toUpperCase();
  return {
    publico: String(p.publico || '').trim().toUpperCase(),
    nome: limparTexto_(p.nome, 100),
    turmaSetor: limparTexto_(p.turmaSetor, 80),
    whatsapp: somenteDigitos_(p.whatsapp),
    email: limparTexto_(p.email, 120).toLowerCase(),
    cpf: somenteDigitos_(p.cpf),
    codigoProduto: String(p.codigoProduto || '').trim().toUpperCase(),
    tamanho: String(p.tamanho || '').trim().toUpperCase(),
    nomeCamisa: limparTexto_(p.nomeCamisa, 50).toUpperCase(),
    cargoFuncao: limparTexto_(p.cargoFuncao, 80).toUpperCase(),
    observacoes: limparTexto_(p.observacoes, 300),
    metodoPagamento: ['PIX','BOLETO','CREDIT_CARD'].includes(metodo) ? metodo : 'PIX',
    parcelas: Math.max(1, Math.floor(Number(p.parcelas || 1))),
    aceitou: Boolean(p.aceitou)
  };
}


function validarPedido_(d) {
  if (!['ESTUDANTE','SERVIDOR'].includes(d.publico)) throw new Error('Público inválido.');
  if (d.nome.length < 4) throw new Error('Informe seu nome completo.');
  if (d.whatsapp.length < 10 || d.whatsapp.length > 13) throw new Error('Informe um WhatsApp válido.');
  const asaas = getAsaasConfig_();
  if (asaas.env === 'PRODUCTION' && !validarCpf_(d.cpf)) throw new Error('Informe um CPF válido para gerar a cobrança no Asaas.');
  if (!d.codigoProduto) throw new Error('Selecione o modelo.');
  if (!d.tamanho) throw new Error('Selecione o tamanho.');
  if (!d.aceitou) throw new Error('É necessário aceitar as condições da encomenda.');
  if (!['PIX','BOLETO','CREDIT_CARD'].includes(d.metodoPagamento)) throw new Error('Forma de pagamento inválida.');
  if (d.metodoPagamento !== 'CREDIT_CARD' && d.parcelas !== 1) d.parcelas = 1;
  if (d.parcelas < 1 || d.parcelas > CFG.CARD_MAX_INSTALLMENTS) throw new Error('Quantidade de parcelas inválida.');
  if (d.publico === 'ESTUDANTE' && !d.turmaSetor) throw new Error('Informe sua turma.');
  if (d.publico === 'SERVIDOR' && !d.turmaSetor) throw new Error('Informe seu setor/área.');
  if (d.publico === 'SERVIDOR' && (d.nomeCamisa || d.cargoFuncao) && (!d.nomeCamisa || !d.cargoFuncao)) {
    throw new Error('Para personalizar, informe tanto o nome quanto o cargo/área/função.');
  }
}




function getPublicSiteUrl_() {
  const cfg = getConfigMap_();
  const configured = String(cfg['URL pública da loja'] || '').trim().replace(/\/+$/, '');
  if (configured) return configured;

  const scriptUrl = String(ScriptApp.getService().getUrl() || '').trim();
  return scriptUrl.replace(/\/+$/, '');
}

function buildTrackingUrl_(pedidoId, whatsapp) {
  const base = getPublicSiteUrl_();
  if (!base) return '';
  const sep = base.indexOf('?') >= 0 ? '&' : '?';
  return base + sep
    + 'p=acompanhar'
    + '&pedido=' + encodeURIComponent(String(pedidoId || ''))
    + '&whatsapp=' + encodeURIComponent(String(whatsapp || ''));
}

function obterOuCriarClienteAsaas_(dados) {
  const d = dados || {};
  const payload = {
    name: String(d.nome || '').trim(),
    cpfCnpj: String(d.cpfCnpj || '').trim(),
    mobilePhone: String(d.whatsapp || '').trim(),
    notificationDisabled: !Boolean(d.notificacoesAtivas)
  };
  if (d.email) payload.email = String(d.email).trim();

  // No Sandbox usamos um CPF de teste compartilhado. Reutilizá-lo faria todos os
  // pedidos de teste parecerem pertencer à mesma pessoa, então criamos um cliente
  // separado por pedido. Em Produção, procuramos antes pelo CPF/CNPJ para evitar
  // cadastros duplicados e manter o histórico do pagador.
  if (String(d.ambiente || '').toUpperCase() === 'PRODUCTION') {
    const lista = asaasRequest_(
      '/customers?cpfCnpj=' + encodeURIComponent(payload.cpfCnpj) + '&limit=10&offset=0',
      'get'
    );
    const existentes = Array.isArray(lista.data) ? lista.data : [];
    const existente = existentes.find(c => !c.deleted) || existentes[0];

    if (existente && existente.id) {
      const update = {
        name: payload.name,
        mobilePhone: payload.mobilePhone,
        notificationDisabled: payload.notificationDisabled
      };
      if (payload.email) update.email = payload.email;
      try {
        return asaasRequest_('/customers/' + encodeURIComponent(existente.id), 'put', update);
      } catch (err) {
        registrarLog_('AVISO CLIENTE ASAAS', d.pedidoId || '', 'ASAAS', 'Falha ao atualizar cliente existente: ' + err.message);
        return existente;
      }
    }
  }

  payload.externalReference = 'LOJA-GREMIO-' + String(d.pedidoId || Utilities.getUuid());
  return asaasRequest_('/customers', 'post', payload);
}

function recuperarCobrancaPorPedidoAsaas_(pedidoId) {
  if (!pedidoId) return null;
  const lista = asaasRequest_(
    '/payments?externalReference=' + encodeURIComponent(String(pedidoId)) + '&limit=10&offset=0',
    'get'
  );
  const data = Array.isArray(lista.data) ? lista.data : [];
  return data.find(p => String(p.externalReference || '') === String(pedidoId) && String(p.status || '').toUpperCase() !== 'DELETED')
    || data.find(p => String(p.externalReference || '') === String(pedidoId))
    || null;
}

function criarOuRecuperarCobrancaAsaas_(dados) {
  const d = dados || {};
  // O pedido acabou de receber um ID único; consultar o Asaas ANTES da primeira
  // tentativa era uma chamada de rede desnecessária no caminho crítico. Em caso
  // de timeout/erro continuamos fazendo a recuperação por externalReference abaixo,
  // preservando a proteção contra duplicidade sem atrasar toda compra normal.
  const billingType = ['PIX','BOLETO','CREDIT_CARD'].includes(String(d.billingType || '').toUpperCase()) ? String(d.billingType).toUpperCase() : 'PIX';
  const parcelas = Math.max(1,Math.floor(Number(d.parcelas || 1)));
  const body = {
    customer:d.customerId,
    billingType,
    dueDate:d.dueDate,
    description:String(d.description || '').slice(0,500),
    externalReference:String(d.pedidoId || '')
  };
  if (billingType === 'CREDIT_CARD' && parcelas >= 2) {
    body.installmentCount = parcelas;
    body.totalValue = roundMoney_(d.value);
  } else {
    body.value = roundMoney_(d.value);
  }
  if (d.callbackUrl) body.callback = {successUrl:String(d.callbackUrl),autoRedirect:true};

  try { return asaasRequest_('/payments','post',body); }
  catch(createErr){
    try {
      const recuperada=recuperarCobrancaPorPedidoAsaas_(d.pedidoId);
      if (recuperada && recuperada.id) {
        registrarLog_('COBRANÇA ASAAS RECUPERADA APÓS ERRO',d.pedidoId,'ASAAS',recuperada.id);
        return recuperada;
      }
    } catch (_) {}

    // O Asaas exige que callback.successUrl pertença a um domínio cadastrado
    // em Configurações da conta → Informações. A ausência desse cadastro não
    // pode bloquear a venda: repetimos a criação sem callback e mantemos
    // consulta, sincronização e webhook como fontes de verdade do pagamento.
    if (body.callback && erroCallbackDominioAsaas_(createErr)) {
      registrarLog_(
        'AVISO CALLBACK ASAAS',
        d.pedidoId,
        'ASAAS',
        'Domínio de redirecionamento ainda não cadastrado. Cobrança criada sem callback; acompanhamento e sincronização permanecem ativos.'
      );
      const semCallback = Object.assign({}, body);
      delete semCallback.callback;
      try {
        return asaasRequest_('/payments','post',semCallback);
      } catch(retryErr) {
        try {
          const recuperada=recuperarCobrancaPorPedidoAsaas_(d.pedidoId);
          if (recuperada && recuperada.id) {
            registrarLog_('COBRANÇA ASAAS RECUPERADA APÓS RETRY',d.pedidoId,'ASAAS',recuperada.id);
            return recuperada;
          }
        } catch (_) {}
        throw retryErr;
      }
    }
    throw createErr;
  }
}

function erroCallbackDominioAsaas_(err) {
  const msg = String(err && err.message ? err.message : err || '').toLowerCase();
  const falaDeDominio = msg.includes('domínio') || msg.includes('dominio');
  const falaDeCadastro = msg.includes('configurad') || msg.includes('cadastr') || msg.includes('site') || msg.includes('successurl') || msg.includes('callback');
  return falaDeDominio && falaDeCadastro;
}


function humanizarMeioPagamentoAsaas_(billingType) {
  switch (String(billingType || '').toUpperCase()) {
    case 'PIX': return 'Pix';
    case 'BOLETO': return 'Boleto';
    case 'CREDIT_CARD': return 'Cartão';
    case 'DEBIT_CARD': return 'Cartão de débito';
    case 'UNDEFINED': return 'A escolher';
    default: return '';
  }
}

function garantirSheetEventosAsaas_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CFG.SHEET_ASAAS_EVENTOS);
  if (!sh) sh = ss.insertSheet(CFG.SHEET_ASAAS_EVENTOS);

  const headers = ['Evento ID','Evento','Pagamento Asaas','Pedido','Recebido em','Processado em','Status','Resultado'];
  if (sh.getMaxColumns() < headers.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  }
  sh.getRange(1,1,1,headers.length).setValues([headers])
    .setBackground('#063667').setFontColor('#fff').setFontWeight('bold').setWrap(true);
  sh.setFrozenRows(1);
  return sh;
}

function processAsaasWebhookEvent(payload) {
  const evt = payload || {};
  const eventId = String(evt.id || '').trim();
  const eventName = String(evt.event || '').trim().toUpperCase();
  const payment = evt.payment || {};

  if (!eventId || !eventName) throw new Error('Webhook Asaas inválido: evento sem identificação.');
  if (!payment || !payment.id) {
    return { ok: true, ignored: true, eventId: eventId, reason: 'Evento sem objeto payment.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let shEvents;
  let eventRow = 0;
  try {
    shEvents = garantirSheetEventosAsaas_();

    if (shEvents.getLastRow() > 1) {
      const found = shEvents.getRange(2,1,shEvents.getLastRow()-1,1)
        .createTextFinder(eventId).matchEntireCell(true).findNext();
      if (found) {
        return { ok: true, duplicate: true, eventId: eventId };
      }
    }

    const pedidoRef = String(payment.externalReference || '').trim().toUpperCase();
    const pedidoRow = pedidoRef ? localizarPedidoRow_(pedidoRef) : localizarPedidoPorPagamentoAsaas_(payment.id);

    shEvents.appendRow([
      eventId,
      eventName,
      payment.id || '',
      pedidoRef || '',
      new Date(),
      '',
      'RECEBIDO',
      ''
    ]);
    eventRow = shEvents.getLastRow();

    if (!pedidoRow) {
      shEvents.getRange(eventRow,6,1,3).setValues([[new Date(),'IGNORADO','Pedido não localizado']]);
      return { ok: true, ignored: true, eventId: eventId, reason: 'Pedido não localizado.' };
    }

    // O lock já está ativo para garantir idempotência do evento.
    aplicarPagamentoAsaasNaLinha_(pedidoRow, payment, 'WEBHOOK ASAAS', true);

    const pedidoId = String(SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS).getRange(pedidoRow,1).getValue() || '');
    shEvents.getRange(eventRow,4).setValue(pedidoId);
    shEvents.getRange(eventRow,6,1,3).setValues([[new Date(),'PROCESSADO','OK']]);

    return { ok: true, duplicate: false, eventId: eventId, pedido: pedidoId };
  } catch (err) {
    if (shEvents && eventRow) {
      try {
        shEvents.getRange(eventRow,6,1,3).setValues([[new Date(),'ERRO',String(err.message || err).slice(0,500)]]);
      } catch (_) {}
    }
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function localizarPedidoPorPagamentoAsaas_(paymentId) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  if (!sh || sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2,22,sh.getLastRow()-1,1).getDisplayValues().flat(); // V
  const target = String(paymentId || '').trim();
  const idx = ids.findIndex(v => String(v || '').trim() === target);
  return idx >= 0 ? idx + 2 : 0;
}

function diagnosticarAsaas() {
  const asaas=getAsaasConfig_();
  const linhas=[];
  linhas.push('Versão da loja: '+CFG.COMMERCE_VERSION);
  linhas.push('Ambiente: '+asaas.env);
  linhas.push('API Key: '+(asaas.configured?'CONFIGURADA':'NÃO CONFIGURADA'));
  let apiOk=false,simulateOk=false;
  if(asaas.configured){
    try{ asaasRequest_('/payments?limit=1&offset=0','get'); apiOk=true; linhas.push('API: OK'); }catch(err){ linhas.push('API: ERRO — '+err.message); }
  }
  if(apiOk){
    try{ const keys=asaasRequest_('/pix/addressKeys?limit=20&offset=0','get'); const total=Array.isArray(keys.data)?keys.data.length:Number(keys.totalCount||0); linhas.push('Chave Pix: '+(total>0?'OK':'NÃO CADASTRADA')); }catch(_){ linhas.push('Chave Pix: não foi possível verificar'); }
    try{ const sim=simularOpcoesCartao('EST-POLO'); simulateOk=Boolean(sim && sim.opcoes && sim.opcoes.length); linhas.push('Simulador cartão: '+(simulateOk?'OK ('+sim.opcoes.length+' opção(ões))':'SEM OPÇÕES')); }catch(err){ linhas.push('Simulador cartão: ERRO — '+err.message); }
  }
  const triggers=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='sincronizarPagamentosAsaas');
  linhas.push('Sincronização automática: '+(triggers.length?'ATIVA':'NÃO INSTALADA'));
  const pol=getPaymentPolicy_();
  linhas.push('Pix: ATIVO');
  linhas.push('Boleto: ATIVO');
  linhas.push('Cartão: ATIVO • até '+pol.maxParcelas+'x • parcela mínima R$ '+pol.parcelaMinima.toFixed(2));
  linhas.push('Repasse da taxa do cartão ao comprador: '+(pol.repassarTaxaCartao?'SIM':'NÃO'));
  linhas.push('Redirecionamento pós-pagamento: '+(getPublicSiteUrl_()?'CONFIGURADO':'SEM URL'));
  linhas.push('Webhook: código preparado para endpoint Vercel.');
  const msg=linhas.join('\n');
  notificar_(apiOk && simulateOk ? '✅ Diagnóstico Asaas' : '⚠️ Diagnóstico Asaas',msg);
  return {ok:apiOk && simulateOk,ambiente:asaas.env,relatorio:linhas};
}


function validarAtivacaoOficialAsaas() {
  const props = PropertiesService.getScriptProperties();
  const asaas = getAsaasConfig_();
  const linhas = [];
  let ok = true;

  linhas.push('Loja Digital do Grêmio • Produção oficial');

  if (asaas.env !== 'PRODUCTION') {
    ok = false;
    linhas.push('Ambiente de pagamentos: PENDENTE');
    linhas.push('Defina ASAAS_ENV como PRODUCTION nas Propriedades do script.');
  } else {
    linhas.push('Ambiente de pagamentos: PRODUÇÃO');
  }

  const key = String(props.getProperty('ASAAS_API_KEY') || '').trim();
  if (!key) {
    ok = false;
    linhas.push('API Key: NÃO CONFIGURADA');
  } else if (!key.startsWith('$aact_prod_')) {
    ok = false;
    linhas.push('API Key: a chave configurada não parece ser uma chave de Produção do Asaas.');
  } else {
    linhas.push('API Key: PRODUÇÃO CONFIGURADA');
  }

  if (asaas.env === 'PRODUCTION' && key.startsWith('$aact_prod_')) {
    try {
      asaasRequest_('/payments?limit=1&offset=0', 'get');
      linhas.push('Conexão com API oficial: OK');
    } catch (err) {
      ok = false;
      linhas.push('Conexão com API oficial: ERRO — ' + err.message);
    }

    try {
      const keys = asaasRequest_('/pix/addressKeys?limit=20&offset=0', 'get');
      const total = Array.isArray(keys.data) ? keys.data.length : Number(keys.totalCount || 0);
      if (total > 0) linhas.push('Pix: chave cadastrada');
      else { ok = false; linhas.push('Pix: cadastre uma chave Pix na conta Asaas de Produção.'); }
    } catch (err) {
      ok = false;
      linhas.push('Pix: não foi possível validar a chave — ' + err.message);
    }
  }

  const publicUrl = getPublicSiteUrl_();
  if (publicUrl) linhas.push('URL pública da loja: ' + publicUrl);
  else { ok = false; linhas.push('URL pública da loja: NÃO CONFIGURADA'); }

  const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sincronizarPagamentosAsaas');
  if (triggers.length) linhas.push('Sincronização automática: ATIVA');
  else { ok = false; linhas.push('Sincronização automática: NÃO INSTALADA'); }

  linhas.push(ok ? 'STATUS: PRONTO PARA VENDAS REAIS' : 'STATUS: EXISTEM PENDÊNCIAS ANTES DE ABRIR AS VENDAS');
  const msg = linhas.join('\n');
  notificar_(ok ? '✅ Asaas Produção pronto' : '⚠️ Ativação oficial incompleta', msg);
  return { ok: ok, ambiente: asaas.env, relatorio: linhas };
}


function getAsaasConfig_() {
  const props = PropertiesService.getScriptProperties();
  const key = String(props.getProperty('ASAAS_API_KEY') || '').trim();
  const envRaw = String(props.getProperty('ASAAS_ENV') || '').trim().toUpperCase();
  const env = envRaw === 'PRODUCTION' ? 'PRODUCTION' : (envRaw === 'SANDBOX' ? 'SANDBOX' : 'UNCONFIGURED');
  return {
    key: key,
    env: env,
    configured: Boolean(key),
    baseUrl: env === 'PRODUCTION'
      ? 'https://api.asaas.com/v3'
      : (env === 'SANDBOX' ? 'https://api-sandbox.asaas.com/v3' : '')
  };
}

function asaasRequest_(path, method, body) {
  const cfg = getAsaasConfig_();
  if (!cfg.configured) throw new Error('ASAAS_API_KEY não configurada.');
  if (!cfg.baseUrl) throw new Error('ASAAS_ENV não configurado. Use PRODUCTION para a loja oficial.');

  const options = {
    method: String(method || 'get').toLowerCase(),
    muteHttpExceptions: true,
    headers: {
      'access_token': cfg.key,
      'User-Agent': 'GremioEstudantilPAA-Loja/7.3.0',
      'Accept': 'application/json'
    }
  };

  if (body !== undefined && body !== null && options.method !== 'get') {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(body);
  }

  const response = UrlFetchApp.fetch(cfg.baseUrl + path, options);
  const status = response.getResponseCode();
  const text = response.getContentText() || '';
  let json = {};

  if (text) {
    try {
      json = JSON.parse(text);
    } catch (_) {
      json = { raw: text };
    }
  }

  if (status < 200 || status >= 300) {
    const descriptions = Array.isArray(json.errors)
      ? json.errors.map(e => e.description || e.code).filter(Boolean).join(' | ')
      : '';
    throw new Error(`Asaas HTTP ${status}: ${descriptions || json.message || json.raw || 'erro não identificado'}`);
  }

  return json;
}

function testarConexaoAsaas() {
  const cfg = getAsaasConfig_();

  if (!cfg.configured) {
    notificar_('❌ Asaas não configurado', 'Defina ASAAS_API_KEY e ASAAS_ENV nas Propriedades do script.');
    return { ok: false, mensagem: 'Asaas não configurado.' };
  }

  try {
    asaasRequest_('/payments?limit=1&offset=0', 'get');

    let pixInfo = '';
    try {
      const keys = asaasRequest_('/pix/addressKeys?limit=20&offset=0', 'get');
      const total = Array.isArray(keys.data) ? keys.data.length : Number(keys.totalCount || 0);
      pixInfo = total > 0
        ? `\nChave Pix: OK (${total} cadastrada${total === 1 ? '' : 's'}).`
        : '\nChave Pix: ainda não cadastrada.';
    } catch (_) {}

    const msg = `Ambiente: ${cfg.env}\nA API respondeu corretamente.${pixInfo}\nJá pode testar o Web App.`;
    notificar_('✅ Conexão Asaas OK', msg);
    return { ok: true, ambiente: cfg.env, mensagem: msg };
  } catch (err) {
    notificar_('❌ Erro na conexão Asaas', err.message);
    return { ok: false, mensagem: err.message };
  }
}
function notificar_(titulo, mensagem) {
  // Preferência: alerta quando a função foi chamada pelo menu da planilha.
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(titulo, mensagem, ui.ButtonSet.OK);
    return;
  } catch (_) {}

  // Fallback: toast quando executado pelo editor do Apps Script.
  try {
    const ss = SpreadsheetApp.getActive();
    if (ss) {
      ss.toast(String(mensagem || ''), String(titulo || 'Sistema de Camisas'), 8);
      return;
    }
  } catch (_) {}

  console.log(`${titulo}: ${mensagem}`);
}


function verificarOuCriarChavePixSandbox() {
  const asaas = getAsaasConfig_();
  if (asaas.env !== 'SANDBOX') {
    notificar_('Chave Pix', 'Esta função automática está disponível somente no Sandbox. Em Produção, cadastre a chave Pix na conta Asaas aprovada.');
    return { ok: false, ambiente: asaas.env };
  }

  try {
    const result = garantirChavePixSandbox_();
    const msg = result.created
      ? `Chave Pix aleatória criada com sucesso no Sandbox.\n${result.key || ''}\nCrie um NOVO pedido de teste para gerar o QR Code.`
      : `Já existe uma chave Pix cadastrada no Sandbox.\n${result.key || ''}`;
    notificar_('✅ Chave Pix pronta', msg);
    return { ok: true, created: result.created, key: result.key || '' };
  } catch (err) {
    notificar_('❌ Erro na chave Pix', err.message);
    return { ok: false, mensagem: err.message };
  }
}

function garantirChavePixSandbox_() {
  const asaas = getAsaasConfig_();
  if (asaas.env !== 'SANDBOX') return { created: false, key: '' };

  // A existência da chave não muda a cada pedido. Cachear essa verificação evita
  // uma chamada extra à API em cada checkout Pix do Sandbox.
  const cache = CacheService.getScriptCache();
  const cacheKey = 'ASAAS_SANDBOX_PIX_READY_V724';
  if (cache.get(cacheKey) === '1') return { created: false, key: '' };

  const list = asaasRequest_('/pix/addressKeys?limit=20&offset=0', 'get');
  const data = Array.isArray(list.data) ? list.data : [];
  const ativa = data.find(k => String(k.status || 'ACTIVE').toUpperCase() !== 'DELETED') || data[0];

  if (ativa) {
    cache.put(cacheKey, '1', 21600);
    return { created: false, key: ativa.key || ativa.addressKey || ativa.id || '' };
  }

  const created = asaasRequest_('/pix/addressKeys', 'post', { type: 'EVP' });
  cache.put(cacheKey, '1', 21600);
  return {
    created: true,
    key: created.key || created.addressKey || created.id || ''
  };
}

function configurarNotificacoesClienteAsaas_(customerId, cfg) {
  if (!customerId) return;

  const lista = asaasRequest_(`/customers/${encodeURIComponent(customerId)}/notifications`, 'get');
  const notifs = Array.isArray(lista.data) ? lista.data : [];
  const sms = sim_(cfg['SMS Asaas ao cliente']);
  const whatsapp = sim_(cfg['WhatsApp Asaas ao cliente']);

  const eventos = new Set([
    'PAYMENT_CREATED',
    'PAYMENT_RECEIVED',
    'PAYMENT_UPDATED',
    'PAYMENT_DUEDATE_WARNING',
    'PAYMENT_OVERDUE'
  ]);

  notifs.forEach(n => {
    if (!n.id || !eventos.has(String(n.event || '').toUpperCase())) return;

    asaasRequest_(`/notifications/${encodeURIComponent(n.id)}`, 'put', {
      enabled: true,
      emailEnabledForProvider: false,
      smsEnabledForProvider: false,
      emailEnabledForCustomer: true,
      smsEnabledForCustomer: sms,
      phoneCallEnabledForCustomer: false,
      whatsappEnabledForCustomer: whatsapp
    });
  });
}

function sim_(value) {
  return ['SIM', 'S', 'TRUE', '1', 'YES'].includes(String(value || '').trim().toUpperCase());
}

function instalarTriggerSincronizacaoAsaas_() {
  const cfgSheet = getConfigMap_();
  let minutos = Number(cfgSheet['Sincronização automática Asaas (min)'] || 5);
  const permitidos = [1, 5, 10, 15, 30];
  if (!permitidos.includes(minutos)) minutos = 5;

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sincronizarPagamentosAsaas')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('sincronizarPagamentosAsaas')
    .timeBased()
    .everyMinutes(minutos)
    .create();
}

function sincronizarPagamentosAsaas() {
  const asaas = getAsaasConfig_();
  if (!asaas.configured) return;

  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  if (!sh || sh.getLastRow() < 2) return;

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 38).getValues();
  const finais = new Set(['PAGO', 'CONFIRMADO', 'CANCELADO', 'ESTORNADO']);

  rows.forEach((row, idx) => {
    const statusLocal = String(row[18] || '').toUpperCase();
    const paymentId = String(row[21] || '').trim();
    if (!paymentId || finais.has(statusLocal)) return;

    try {
      sincronizarPagamentoLinha_(idx + 2);
    } catch (err) {
      registrarLog_('ERRO SINCRONIZAÇÃO', row[0], 'ASAAS', err.message);
    }
  });
}

function sincronizarPagamentoLinha_(rowIndex) {
  const sh=SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  const row=sh.getRange(rowIndex,1,1,38).getValues()[0];
  const paymentId=String(row[21] || '').trim();
  if(!paymentId) return null;
  const payment=asaasRequest_(`/payments/${encodeURIComponent(paymentId)}`,'get');
  aplicarPagamentoAsaasNaLinha_(rowIndex,payment,'POLLING');
  return payment;
}



function consolidarParcelamentoAsaas_(payment) {
  const result={bruto:Number(payment && payment.value || 0),liquido:Number(payment && payment.netValue || 0),parcelas:1,installmentId:String(payment && payment.installment || '')};
  if(!result.installmentId) return result;
  try{
    const list=asaasRequest_(`/installments/${encodeURIComponent(result.installmentId)}/payments`,'get');
    const items=Array.isArray(list.data) ? list.data : Array.isArray(list) ? list : [];
    if(items.length){
      result.bruto=roundMoney_(items.reduce((a,p)=>a+Number(p.value || 0),0));
      result.liquido=roundMoney_(items.reduce((a,p)=>a+Number(p.netValue || 0),0));
      result.parcelas=items.length;
    }
  }catch(err){
    registrarLog_('AVISO CONSOLIDAÇÃO PARCELAMENTO',String(payment.externalReference || ''),'ASAAS',err.message);
  }
  return result;
}

function aplicarPagamentoAsaasNaLinha_(rowIndex, payment, origem, skipLock) {
  if(!payment || !payment.id) return null;
  const lock=skipLock ? null : LockService.getScriptLock();
  if(lock) lock.waitLock(15000);
  try{
    const sh=SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
    const row=sh.getRange(rowIndex,1,1,38).getValues()[0];
    const pedidoId=String(row[0] || '').trim();
    const novoStatus=mapAsaasStatus_(payment.status);
    const statusAtual=String(row[18] || '');
    const consolidado=consolidarParcelamentoAsaas_(payment);
    const brutoPedido=Number(row[13] || consolidado.bruto || payment.value || 0);
    const liquidoCore=Number(consolidado.liquido || payment.netValue || row[15] || 0);
    const metodo=String(row[32] || payment.billingType || '');
    const parcelas=Number(row[33] || consolidado.parcelas || 1);
    if(payment.installment && !row[37]) sh.getRange(rowIndex,38).setValue(payment.installment);

    try{
      atualizarPagamentoCommerce_(pedidoId,{provedor:'ASAAS',idExterno:payment.id || row[21] || '',tipo:metodo || payment.billingType || 'UNDEFINED',valor:brutoPedido,taxa:liquidoCore>0?Math.max(0,brutoPedido-liquidoCore):Number(row[14]||0),liquido:liquidoCore,status:novoStatus || statusAtual,link:payment.invoiceUrl || row[22] || '',pagoEm:payment.paymentDate || payment.clientPaymentDate || ''});
    }catch(commerceErr){ registrarLog_('AVISO SINCRONIZAÇÃO CORE',pedidoId,'SISTEMA',commerceErr.message); }
    if(!novoStatus) return payment;

    if(novoStatus!==statusAtual){
      sh.getRange(rowIndex,19).setValue(novoStatus);
      sh.getRange(rowIndex,27).setValue(new Date());
      sh.getRange(rowIndex,28).setValue(origem || 'ASAAS');
      if(novoStatus==='PAGO' || novoStatus==='CONFIRMADO'){
        const dataPagamento=payment.paymentDate || payment.clientPaymentDate || new Date();
        sh.getRange(rowIndex,24).setValue(parseAsaasDate_(dataPagamento));
        const custoPedido=Number(row[16] || 0);
        if(Number.isFinite(liquidoCore) && liquidoCore>0 && brutoPedido>0){
          const taxaReal=Math.max(0,roundMoney_(brutoPedido-liquidoCore));
          sh.getRange(rowIndex,15).setValue(taxaReal);
          sh.getRange(rowIndex,16).setValue(liquidoCore);
          sh.getRange(rowIndex,18).setValue(roundMoney_(liquidoCore-custoPedido));
          try{
            registrarMovimentacaoFinanceiraCommerce_({referencia:'ASAAS:'+(payment.installment || payment.id),pedidoId,tipo:'VENDA',categoria:'RECEITA DE VENDAS',descricao:'Pagamento confirmado • '+String(row[8] || ''),valorBruto:brutoPedido,taxa:taxaReal,valorLiquido:liquidoCore,custo:custoPedido,resultado:roundMoney_(liquidoCore-custoPedido),provedor:'ASAAS',meioPagamento:(metodo || payment.billingType || 'NÃO INFORMADO')+(parcelas>1?' '+parcelas+'x':'')});
          }catch(_){}
        }
      }
      registrarLog_('STATUS PAGAMENTO',pedidoId,origem || 'ASAAS',`${statusAtual || '(vazio)'} → ${novoStatus} (${payment.status || ''}) • Meio: ${metodo || payment.billingType || 'NÃO INFORMADO'}${parcelas>1?' • '+parcelas+'x':''}`);
      if(novoStatus==='PAGO' || novoStatus==='CONFIRMADO'){ try{ enviarEmailPagamentoConfirmado_(rowIndex); }catch(mailErr){ registrarLog_('AVISO E-MAIL',pedidoId,'SISTEMA',mailErr.message); } }
    }
    return payment;
  }finally{ if(lock) lock.releaseLock(); }
}


function mapAsaasStatus_(status) {
  switch (String(status || '').toUpperCase()) {
    case 'RECEIVED':
    case 'RECEIVED_IN_CASH':
      return 'PAGO';
    case 'CONFIRMED':
      return 'CONFIRMADO';
    case 'PENDING':
      return 'AGUARDANDO PAGAMENTO';
    case 'AWAITING_RISK_ANALYSIS':
    case 'AUTHORIZED':
      return 'EM ANÁLISE';
    case 'OVERDUE':
      return 'VENCIDO';
    case 'REFUNDED':
    case 'REFUND_REQUESTED':
    case 'REFUND_IN_PROGRESS':
      return 'ESTORNADO';
    case 'CHARGEBACK_REQUESTED':
    case 'CHARGEBACK_DISPUTE':
    case 'AWAITING_CHARGEBACK_REVERSAL':
      return 'CONTESTADO';
    case 'DELETED':
      return 'CANCELADO';
    default:
      return '';
  }
}

function simularPagamentoSandboxSelecionado() {
  const asaas = getAsaasConfig_();

  if (asaas.env !== 'SANDBOX') {
    notificar_('Simulação indisponível', 'Função disponível somente no Sandbox.');
    return;
  }

  const sh = SpreadsheetApp.getActiveSheet();
  if (!sh || sh.getName() !== CFG.SHEET_PEDIDOS) {
    notificar_('Selecione um pedido', 'Abra a aba “Pedidos” e selecione uma célula da linha que deseja testar.');
    return;
  }

  const rowIndex = sh.getActiveRange().getRow();
  if (rowIndex < 2) {
    notificar_('Selecione um pedido', 'Selecione uma linha de pedido.');
    return;
  }

  const pedido = String(sh.getRange(rowIndex, 1).getValue() || '');
  const paymentId = String(sh.getRange(rowIndex, 22).getValue() || '');

  if (!paymentId) {
    notificar_('Cobrança não encontrada', 'Essa linha ainda não possui ID de cobrança Asaas.');
    return;
  }

  try {
    asaasRequest_(`/sandbox/payment/${encodeURIComponent(paymentId)}/confirm`, 'post', {});
    Utilities.sleep(700);
    sincronizarPagamentoLinha_(rowIndex);
    notificar_('✅ Pagamento de teste confirmado', `${pedido} foi simulado no Sandbox e o status foi sincronizado.`);
  } catch (err) {
    notificar_('Erro ao simular pagamento', err.message);
  }
}

function datePlusDays_(days) {
  const d = new Date();
  d.setDate(d.getDate() + Number(days || 0));
  return Utilities.formatDate(d, 'America/Fortaleza', 'yyyy-MM-dd');
}

function parseAsaasDate_(value) {
  if (value instanceof Date) return value;
  const text = String(value || '').trim();
  if (!text) return new Date();

  const iso = new Date(text);
  if (!isNaN(iso.getTime())) return iso;

  const m = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);

  return new Date();
}

function validarCpf_(cpf) {
  const c = somenteDigitos_(cpf);
  if (c.length !== 11 || /^(\d)\1{10}$/.test(c)) return false;

  let soma = 0;
  for (let i = 0; i < 9; i++) soma += Number(c[i]) * (10 - i);
  let d1 = (soma * 10) % 11;
  if (d1 === 10) d1 = 0;
  if (d1 !== Number(c[9])) return false;

  soma = 0;
  for (let i = 0; i < 10; i++) soma += Number(c[i]) * (11 - i);
  let d2 = (soma * 10) % 11;
  if (d2 === 10) d2 = 0;

  return d2 === Number(c[10]);
}



function configurarCodigoAdminPeloMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.prompt(
      '🔐 Área do Grêmio',
      'Crie um código de acesso com pelo menos 6 caracteres. Ele será usado para entrar na área administrativa do sistema.',
      ui.ButtonSet.OK_CANCEL
    );

    if (resp.getSelectedButton() !== ui.Button.OK) return;

    const codigo = String(resp.getResponseText() || '').trim();
    if (codigo.length < 6) {
      ui.alert('Código muito curto. Use pelo menos 6 caracteres.');
      return;
    }

    PropertiesService.getScriptProperties().setProperty('ADMIN_CODE_HASH', hashAdminCode_(codigo));
    PropertiesService.getScriptProperties().deleteProperty('ADMIN_ACCESS_CODE');

    ui.alert('✅ Acesso configurado', 'O código foi salvo de forma resumida (hash). A Área do Grêmio já pode ser acessada pelo Web App.', ui.ButtonSet.OK);
  } catch (err) {
    notificar_('Área do Grêmio', 'Abra a planilha e use o menu “🛍️ Loja do Grêmio” para configurar o código de acesso.');
  }
}

function hashAdminCode_(codigo) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(codigo || '').trim(),
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => (b + 256) % 256).map(b => b.toString(16).padStart(2, '0')).join('');
}

function adminLogin(codigo, nome) {
  const props = PropertiesService.getScriptProperties();
  const storedHash = String(props.getProperty('ADMIN_CODE_HASH') || '').trim();
  const legacy = String(props.getProperty('ADMIN_ACCESS_CODE') || '').trim();
  const c = String(codigo || '').trim();
  const operador = limparTexto_(nome, 80);

  if (!storedHash && !legacy) {
    throw new Error('A Área do Grêmio ainda não foi configurada. Use o menu da planilha para criar o código de acesso.');
  }
  if (!operador || operador.length < 2) {
    throw new Error('Informe seu nome para registrar quem está fazendo as alterações.');
  }

  const ok = storedHash
    ? hashAdminCode_(c) === storedHash
    : c === legacy;

  if (!ok) throw new Error('Código de acesso incorreto.');

  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put(
    CFG.ADMIN_CACHE_PREFIX + token,
    JSON.stringify({ nome: operador, criadoEm: new Date().toISOString() }),
    CFG.ADMIN_SESSION_SECONDS
  );

  registrarLog_('LOGIN ADMIN', '', operador, 'Área do Grêmio');
  return { ok: true, token: token, nome: operador, expiresIn: CFG.ADMIN_SESSION_SECONDS };
}

function adminLogout(token) {
  if (token) CacheService.getScriptCache().remove(CFG.ADMIN_CACHE_PREFIX + token);
  return { ok: true };
}

function getAdminSession_(token) {
  const t = String(token || '').trim();
  if (!t) throw new Error('Sessão administrativa não encontrada. Entre novamente.');

  const raw = CacheService.getScriptCache().get(CFG.ADMIN_CACHE_PREFIX + t);
  if (!raw) throw new Error('Sua sessão expirou. Entre novamente na Área do Grêmio.');

  let data = {};
  try { data = JSON.parse(raw); } catch (_) {}
  if (!data.nome) data.nome = 'Grêmio Estudantil';

  // Renova a sessão enquanto estiver em uso.
  CacheService.getScriptCache().put(
    CFG.ADMIN_CACHE_PREFIX + t,
    JSON.stringify(data),
    CFG.ADMIN_SESSION_SECONDS
  );
  return data;
}

function adminGetData(token, filtros) {
  const session = getAdminSession_(token);
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  const cfg = getConfigMap_();
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, 38).getValues() : [];
  const f = filtros || {};

  const resumo = {
    total: rows.length,
    pagos: 0,
    aguardando: 0,
    emProducao: 0,
    prontos: 0,
    entregues: 0,
    receitaPaga: 0
  };

  rows.forEach(r => {
    const pag = String(r[18] || '').toUpperCase();
    const prod = String(r[19] || '').toUpperCase();
    const ent = String(r[20] || '').toUpperCase();
    if (['PAGO', 'CONFIRMADO'].includes(pag)) {
      resumo.pagos++;
      resumo.receitaPaga += Number(r[13] || 0);
    }
    if (['AGUARDANDO PAGAMENTO', 'VENCIDO', 'EM ANÁLISE'].includes(pag)) resumo.aguardando++;
    if (prod === 'EM PRODUÇÃO') resumo.emProducao++;
    if (prod === 'PRONTO') resumo.prontos++;
    if (ent === 'ENTREGUE') resumo.entregues++;
  });

  const query = limparTexto_(f.query, 100).toLowerCase();
  const publico = String(f.publico || '').toUpperCase();
  const pagamento = String(f.pagamento || '').toUpperCase();
  const producao = String(f.producao || '').toUpperCase();
  const entrega = String(f.entrega || '').toUpperCase();

  let filtered = rows.filter(r => {
    if (publico && String(r[2] || '').toUpperCase() !== publico) return false;
    if (pagamento && String(r[18] || '').toUpperCase() !== pagamento) return false;
    if (producao && String(r[19] || '').toUpperCase() !== producao) return false;
    if (entrega && String(r[20] || '').toUpperCase() !== entrega) return false;

    if (query) {
      const hay = [
        r[0], r[3], r[4], r[5], r[6], r[8], r[9], r[10], r[11], r[12]
      ].join(' ').toLowerCase();
      if (!hay.includes(query)) return false;
    }
    return true;
  });

  filtered = filtered.slice().reverse().slice(0, 300);

  const pedidos = filtered.map(r => ({
    pedido: r[0],
    data: formatDate_(r[1]),
    publico: r[2],
    nome: r[3],
    turmaSetor: r[4],
    whatsapp: r[5],
    email: r[6],
    modelo: r[8],
    tamanho: r[9],
    nomeCamisa: r[10],
    cargoFuncao: r[11],
    lote: r[12],
    valor: Number(r[13] || 0),
    valorBase: Number(r[34] || r[13] || 0),
    acrescimoPagamento: Number(r[35] || 0),
    metodoPagamento: humanizarMeioPagamentoAsaas_(r[32] || ''),
    parcelas: Number(r[33] || 1),
    valorParcela: Number(r[36] || r[13] || 0),
    statusPagamento: r[18],
    statusProducao: r[19],
    statusEntrega: r[20],
    dataPagamento: formatDate_(r[23]),
    observacoes: r[24],
    prazoPrevisto: dateInputValue_(r[25]),
    prazoPrevistoFormatado: formatDateOnly_(r[25]),
    ultimaAtualizacao: formatDate_(r[26]),
    atualizadoPor: r[27],
    mensagemCliente: r[28],
    observacaoInterna: r[29],
    idClienteAsaas: r[30]
  }));

  return {
    ok: true,
    operador: session.nome,
    resumo: resumo,
    pedidos: pedidos,
    settings: {
      vendasEstudantes: String(cfg['Vendas estudantes'] || 'ABERTAS').toUpperCase(),
      vendasServidores: String(cfg['Vendas servidores'] || 'ABERTAS').toUpperCase(),
      loteEstudantes: String(cfg['Lote atual estudantes'] || ''),
      loteServidores: String(cfg['Lote atual servidores'] || ''),
      prazoPadrao: Number(cfg['Prazo padrão produção (dias)'] || 15),
      emailsAutomaticos: sim_(cfg['E-mails automáticos do sistema']),
      notificacoesAsaas: sim_(cfg['Notificações Asaas ao cliente']),
      smsAsaas: sim_(cfg['SMS Asaas ao cliente']),
      whatsappAsaas: sim_(cfg['WhatsApp Asaas ao cliente'])
    }
  };
}

function adminUpdateOrder(token, payload) {
  const session = getAdminSession_(token);
  const p = payload || {};
  const pedido = String(p.pedido || '').trim().toUpperCase();
  if (!pedido) throw new Error('Pedido inválido.');

  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  const rowIndex = localizarPedidoRow_(pedido);
  if (!rowIndex) throw new Error('Pedido não encontrado.');

  const row = sh.getRange(rowIndex, 1, 1, 38).getValues()[0];

  const producoes = ['NÃO ENVIADO', 'ENVIADO À GRÁFICA', 'EM PRODUÇÃO', 'PRONTO'];
  const entregas = ['NÃO ENTREGUE', 'DISPONÍVEL PARA RETIRADA', 'ENTREGUE'];

  const novaProducao = String(p.statusProducao || row[19] || '').trim().toUpperCase();
  const novaEntrega = String(p.statusEntrega || row[20] || '').trim().toUpperCase();
  if (!producoes.includes(novaProducao)) throw new Error('Status de produção inválido.');
  if (!entregas.includes(novaEntrega)) throw new Error('Status de entrega inválido.');

  let prazo = parseDateInput_(p.prazoPrevisto);
  const cfg = getConfigMap_();

  if (!prazo && novaProducao === 'ENVIADO À GRÁFICA' && String(row[19] || '').toUpperCase() !== 'ENVIADO À GRÁFICA') {
    prazo = new Date();
    prazo.setDate(prazo.getDate() + Number(cfg['Prazo padrão produção (dias)'] || 15));
    prazo.setHours(12, 0, 0, 0);
  }

  const mensagem = limparTexto_(p.mensagemCliente, 500);
  const interna = limparTexto_(p.observacaoInterna, 700);

  const alteracoes = [];
  if (String(row[19] || '') !== novaProducao) alteracoes.push(`Produção: ${row[19] || '-'} → ${novaProducao}`);
  if (String(row[20] || '') !== novaEntrega) alteracoes.push(`Entrega: ${row[20] || '-'} → ${novaEntrega}`);

  const prazoAnterior = formatDateOnly_(row[25]);
  const prazoNovo = formatDateOnly_(prazo);
  if (prazoAnterior !== prazoNovo) alteracoes.push(`Previsão: ${prazoAnterior || '-'} → ${prazoNovo || '-'}`);
  if (String(row[28] || '') !== mensagem && mensagem) alteracoes.push('Nova mensagem do Grêmio');

  sh.getRange(rowIndex, 20).setValue(novaProducao); // T
  sh.getRange(rowIndex, 21).setValue(novaEntrega);  // U
  sh.getRange(rowIndex, 26).setValue(prazo || ''); // Z
  sh.getRange(rowIndex, 27).setValue(new Date());  // AA
  sh.getRange(rowIndex, 28).setValue(session.nome); // AB
  sh.getRange(rowIndex, 29).setValue(mensagem);     // AC
  sh.getRange(rowIndex, 30).setValue(interna);      // AD

  if (alteracoes.length) {
    registrarLog_('PEDIDO ATUALIZADO', pedido, session.nome, alteracoes.join(' | '));
  }

  let emailEnviado = false;
  if (Boolean(p.notificarCliente) && alteracoes.length) {
    try {
      emailEnviado = enviarEmailAtualizacaoPedido_(rowIndex, alteracoes.join('. '));
    } catch (err) {
      registrarLog_('AVISO E-MAIL', pedido, session.nome, err.message);
    }
  }

  const updated = sh.getRange(rowIndex, 1, 1, 38).getValues()[0];
  const mensagemWhats = montarMensagemWhatsAppStatus_(updated);
  const phone = formatWhatsappLink_(updated[5]);

  return {
    ok: true,
    pedido: pedido,
    emailEnviado: emailEnviado,
    whatsappUrl: phone ? `https://wa.me/${phone}?text=${encodeURIComponent(mensagemWhats)}` : '',
    mensagem: alteracoes.length ? 'Pedido atualizado.' : 'Nenhuma alteração detectada.'
  };
}

function adminUpdateSettings(token, payload) {
  const session = getAdminSession_(token);
  const p = payload || {};

  const vendasEst = String(p.vendasEstudantes || 'ABERTAS').toUpperCase();
  const vendasServ = String(p.vendasServidores || 'ABERTAS').toUpperCase();
  if (!['ABERTAS', 'FECHADAS'].includes(vendasEst) || !['ABERTAS', 'FECHADAS'].includes(vendasServ)) {
    throw new Error('Status de vendas inválido.');
  }

  setConfigValue_('Vendas estudantes', vendasEst);
  setConfigValue_('Vendas servidores', vendasServ);
  setConfigValue_('Lote atual estudantes', limparTexto_(p.loteEstudantes, 60));
  setConfigValue_('Lote atual servidores', limparTexto_(p.loteServidores, 60));
  setConfigValue_('Prazo padrão produção (dias)', Math.max(1, Math.min(120, Number(p.prazoPadrao || 15))));
  setConfigValue_('E-mails automáticos do sistema', Boolean(p.emailsAutomaticos) ? 'SIM' : 'NÃO');
  setConfigValue_('Notificações Asaas ao cliente', Boolean(p.notificacoesAsaas) ? 'SIM' : 'NÃO');
  setConfigValue_('SMS Asaas ao cliente', Boolean(p.smsAsaas) ? 'SIM' : 'NÃO');
  setConfigValue_('WhatsApp Asaas ao cliente', Boolean(p.whatsappAsaas) ? 'SIM' : 'NÃO');

  registrarLog_('CONFIGURAÇÕES ATUALIZADAS', '', session.nome, `Vendas estudantes: ${vendasEst} • Vendas servidores: ${vendasServ}`);

  return { ok: true };
}

function adminGetMockups(token) {
  getAdminSession_(token);
  const ids = getMockupIds_();
  return {
    ok: true,
    estudante: { id: ids.estudante, image: getLogoDataUri_(ids.estudante) },
    servidorPolo: { id: ids.servidorPolo, image: getLogoDataUri_(ids.servidorPolo) },
    servidorCareca: { id: ids.servidorCareca, image: getLogoDataUri_(ids.servidorCareca) }
  };
}

function adminUpdateMockups(token, payload) {
  const session = getAdminSession_(token);
  const p = payload || {};

  const estudante = validarMockupDriveId_(p.estudante, 'Estudantes — Gola Polo');
  const servidorPolo = validarMockupDriveId_(p.servidorPolo, 'Profissionais — Gola Polo');
  const servidorCareca = validarMockupDriveId_(p.servidorCareca, 'Profissionais — Gola Careca');

  setConfigValue_(MOCKUP_CONFIG_KEYS.estudante, estudante);
  setConfigValue_(MOCKUP_CONFIG_KEYS.servidorPolo, servidorPolo);
  setConfigValue_(MOCKUP_CONFIG_KEYS.servidorCareca, servidorCareca);

  CacheService.getScriptCache().remove('CAMISAS_PUBLIC_BOOT_V722');
  registrarLog_(
    'MODELOS ATUALIZADOS',
    '',
    session.nome,
    `Estudantes: ${estudante ? 'configurado' : 'removido'} • Profissionais polo: ${servidorPolo ? 'configurado' : 'removido'} • Profissionais careca: ${servidorCareca ? 'configurado' : 'removido'}`
  );

  return adminGetMockups(token);
}

function adminResendStatusEmail(token, pedido) {
  const session = getAdminSession_(token);
  const rowIndex = localizarPedidoRow_(pedido);
  if (!rowIndex) throw new Error('Pedido não encontrado.');

  const enviado = enviarEmailAtualizacaoPedido_(rowIndex, 'Reenvio manual de atualização solicitado pelo Grêmio.');
  registrarLog_('REENVIO E-MAIL', pedido, session.nome, enviado ? 'Enviado' : 'Não enviado');
  return { ok: true, enviado: enviado };
}


function adminGetOrderHistory(token, pedido) {
  getAdminSession_(token);
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_LOG);
  if (!sh || sh.getLastRow() < 2) return [];

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  const target = String(pedido || '').trim().toUpperCase();

  return rows
    .filter(r => String(r[2] || '').trim().toUpperCase() === target)
    .slice()
    .reverse()
    .slice(0, 30)
    .map(r => ({
      data: formatDate_(r[0]),
      acao: r[1] || '',
      origem: r[3] || '',
      detalhes: r[4] || ''
    }));
}

function localizarPedidoRow_(pedido) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  if (!sh || sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().flat();
  const target = String(pedido || '').trim().toUpperCase();
  const idx = ids.findIndex(v => String(v || '').trim().toUpperCase() === target);
  return idx >= 0 ? idx + 2 : 0;
}

function setConfigValue_(key, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_CONFIG);
  const last = Math.max(sh.getLastRow(), 1);
  const keys = sh.getRange(1, 1, last, 1).getDisplayValues().flat();
  const idx = keys.findIndex(k => String(k || '').trim() === key);

  if (idx >= 0) {
    sh.getRange(idx + 1, 2).setValue(value);
  } else {
    sh.appendRow([key, value]);
  }
}

function parseDateInput_(value) {
  const s = String(value || '').trim();
  if (!s) return null;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
  return isNaN(d.getTime()) ? null : d;
}

function dateInputValue_(value) {
  if (!value) return '';
  try {
    const d = value instanceof Date ? value : new Date(value);
    if (isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, 'America/Fortaleza', 'yyyy-MM-dd');
  } catch (_) {
    return '';
  }
}

function formatWhatsappLink_(phone) {
  let p = somenteDigitos_(phone);
  if (!p) return '';
  if ((p.length === 10 || p.length === 11) && !p.startsWith('55')) p = '55' + p;
  return p;
}

function montarMensagemWhatsAppStatus_(row) {
  const prazo = formatDateOnly_(row[25]);
  const msg = String(row[28] || '');
  return [
    `Olá, ${row[3]}! 👋`,
    `Aqui é o Grêmio Estudantil. Temos uma atualização do pedido ${row[0]} da Nova Camisa Oficial:`,
    `Pagamento: ${row[18]}`,
    `Produção: ${row[19]}`,
    `Entrega: ${row[20]}`,
    prazo ? `Previsão: ${prazo}` : '',
    msg ? `Mensagem: ${msg}` : '',
    'Você pode acompanhar o pedido pelo nosso sistema.'
  ].filter(Boolean).join('\n');
}

function podeEnviarEmailAutomatico_() {
  const cfg = getConfigMap_();
  if (!sim_(cfg['E-mails automáticos do sistema'])) return false;

  const asaas = getAsaasConfig_();
  if (asaas.env === 'SANDBOX' && !sim_(cfg['E-mails de teste no Sandbox'])) return false;

  return true;
}

function enviarEmailPedidoCriado_(rowIndex) {
  if (!podeEnviarEmailAutomatico_()) return false;

  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  const row = sh.getRange(rowIndex, 1, 1, 38).getValues()[0];
  const email = String(row[6] || '').trim();
  if (!email || temFlagNotificacao_(row, 'PEDIDO_CRIADO_EMAIL')) return false;

  const cfg = getConfigMap_();
  const baseUrl = ScriptApp.getService().getUrl() || '';
  const tracking = baseUrl ? `${baseUrl}?p=acompanhar` : '';
  const nomeRemetente = String(cfg['Nome exibido nos e-mails'] || 'Grêmio Estudantil');

  const subject = `Pedido ${row[0]} recebido — Nova Camisa Oficial`;
  const htmlBody = `
    <div style="font-family:Arial,sans-serif;color:#143B60;line-height:1.6">
      <h2 style="color:#063667">Pedido recebido! 👕</h2>
      <p>Olá, <strong>${htmlEscape_(row[3])}</strong>.</p>
      <p>Seu pedido <strong>${htmlEscape_(row[0])}</strong> foi registrado no sistema do Grêmio Estudantil.</p>
      <p>
        <strong>${htmlEscape_(row[8])}</strong><br>
        Tamanho: <strong>${htmlEscape_(row[9])}</strong><br>
        Valor: <strong>${formatMoney_(row[13])}</strong><br>
        Status: <strong>${htmlEscape_(row[18])}</strong>
      </p>
      ${row[22] ? `<p><a href="${htmlEscape_(row[22])}" style="background:#063667;color:#fff;text-decoration:none;padding:12px 16px;border-radius:10px;display:inline-block">ABRIR PAGAMENTO</a></p>` : ''}
      ${tracking ? `<p>Depois, acompanhe o pedido em:<br><a href="${htmlEscape_(tracking)}">${htmlEscape_(tracking)}</a></p>` : ''}
      <p style="color:#6A7680;font-size:12px">Guarde o número do pedido e use o mesmo WhatsApp informado na encomenda para consultar o andamento.</p>
    </div>`;

  enviarEmailSistema_(email, subject, htmlBody, nomeRemetente);
  adicionarFlagNotificacao_(rowIndex, 'PEDIDO_CRIADO_EMAIL');
  registrarLog_('E-MAIL ENVIADO', row[0], 'SISTEMA', 'Confirmação de pedido');
  return true;
}

function enviarEmailPagamentoConfirmado_(rowIndex) {
  if (!podeEnviarEmailAutomatico_()) return false;

  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  const row = sh.getRange(rowIndex, 1, 1, 38).getValues()[0];
  const email = String(row[6] || '').trim();
  if (!email || temFlagNotificacao_(row, 'PAGO_EMAIL')) return false;

  const cfg = getConfigMap_();
  const baseUrl = ScriptApp.getService().getUrl() || '';
  const tracking = baseUrl ? `${baseUrl}?p=acompanhar` : '';
  const nomeRemetente = String(cfg['Nome exibido nos e-mails'] || 'Grêmio Estudantil');

  const subject = `Pagamento confirmado — ${row[0]}`;
  const htmlBody = `
    <div style="font-family:Arial,sans-serif;color:#143B60;line-height:1.6">
      <h2 style="color:#2b6e24">Pagamento confirmado! ✅</h2>
      <p>Olá, <strong>${htmlEscape_(row[3])}</strong>.</p>
      <p>Recebemos o pagamento do pedido <strong>${htmlEscape_(row[0])}</strong>.</p>
      <p>Agora ele seguirá para as próximas etapas de produção. Quando houver uma atualização importante, você poderá receber um novo aviso.</p>
      ${tracking ? `<p><a href="${htmlEscape_(tracking)}" style="background:#063667;color:#fff;text-decoration:none;padding:12px 16px;border-radius:10px;display:inline-block">ACOMPANHAR PEDIDO</a></p>` : ''}
      <p style="color:#6A7680;font-size:12px">Grêmio Estudantil • ${htmlEscape_(cfg['Escola'] || '')}</p>
    </div>`;

  enviarEmailSistema_(email, subject, htmlBody, nomeRemetente);
  adicionarFlagNotificacao_(rowIndex, 'PAGO_EMAIL');
  registrarLog_('E-MAIL ENVIADO', row[0], 'SISTEMA', 'Pagamento confirmado');
  return true;
}

function enviarEmailAtualizacaoPedido_(rowIndex, resumoMudancas) {
  if (!podeEnviarEmailAutomatico_()) return false;

  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  const row = sh.getRange(rowIndex, 1, 1, 38).getValues()[0];
  const email = String(row[6] || '').trim();
  if (!email) return false;

  const cfg = getConfigMap_();
  const baseUrl = ScriptApp.getService().getUrl() || '';
  const tracking = baseUrl ? `${baseUrl}?p=acompanhar` : '';
  const nomeRemetente = String(cfg['Nome exibido nos e-mails'] || 'Grêmio Estudantil');

  const subject = `Atualização do pedido ${row[0]} — Nova Camisa Oficial`;
  const prazo = formatDateOnly_(row[25]);
  const mensagem = String(row[28] || '');

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;color:#143B60;line-height:1.6">
      <h2 style="color:#063667">Seu pedido foi atualizado 📦</h2>
      <p>Olá, <strong>${htmlEscape_(row[3])}</strong>.</p>
      <p>Há uma nova atualização no pedido <strong>${htmlEscape_(row[0])}</strong>.</p>
      <p>
        Pagamento: <strong>${htmlEscape_(row[18])}</strong><br>
        Produção: <strong>${htmlEscape_(row[19])}</strong><br>
        Entrega: <strong>${htmlEscape_(row[20])}</strong>
        ${prazo ? `<br>Previsão: <strong>${htmlEscape_(prazo)}</strong>` : ''}
      </p>
      ${resumoMudancas ? `<p>${htmlEscape_(resumoMudancas)}</p>` : ''}
      ${mensagem ? `<div style="background:#F1F6FA;padding:12px 14px;border-radius:10px"><strong>Mensagem do Grêmio:</strong><br>${htmlEscape_(mensagem)}</div>` : ''}
      ${tracking ? `<p><a href="${htmlEscape_(tracking)}" style="background:#063667;color:#fff;text-decoration:none;padding:12px 16px;border-radius:10px;display:inline-block">ACOMPANHAR PEDIDO</a></p>` : ''}
    </div>`;

  enviarEmailSistema_(email, subject, htmlBody, nomeRemetente);
  registrarLog_('E-MAIL ENVIADO', row[0], 'SISTEMA', 'Atualização de produção/entrega');
  return true;
}

function enviarEmailSistema_(to, subject, htmlBody, name) {
  const cfg = getConfigMap_();
  const replyTo = String(cfg['E-mail / Drive do Grêmio'] || '').trim();

  const options = {
    to: to,
    subject: subject,
    body: String(htmlBody || '').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(),
    htmlBody: htmlBody,
    name: name || 'Grêmio Estudantil'
  };
  if (replyTo) options.replyTo = replyTo;

  MailApp.sendEmail(options);
}

function temFlagNotificacao_(row, flag) {
  const flags = String(row[31] || '').split('|').map(s => s.trim()).filter(Boolean);
  return flags.includes(flag);
}

function adicionarFlagNotificacao_(rowIndex, flag) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  const cell = sh.getRange(rowIndex, 32);
  const flags = String(cell.getValue() || '').split('|').map(s => s.trim()).filter(Boolean);
  if (!flags.includes(flag)) flags.push(flag);
  cell.setValue(flags.join('|'));
}

function htmlEscape_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatMoney_(value) {
  return 'R$ ' + Number(value || 0).toFixed(2).replace('.', ',');
}

function getLogoDataUri_(fileId) {
  const id = String(fileId || '').trim();
  if (!id) return '';
  try {
    const blob = DriveApp.getFileById(id).getBlob();
    const mime = blob.getContentType() || 'image/png';
    return `data:${mime};base64,${Utilities.base64Encode(blob.getBytes())}`;
  } catch (err) {
    return '';
  }
}

function registrarLog_(acao, pedido, origem, detalhes) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_LOG);
  if (!sh) return;
  sh.appendRow([new Date(), acao, pedido || '', origem || '', detalhes || '']);
}

function limparTexto_(value, maxLen) {
  return String(value || '')
    .replace(/[<>]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, maxLen || 200);
}

function somenteDigitos_(value) {
  return String(value || '').replace(/\D/g, '');
}

function formatDate_(value) {
  if (!value) return '';
  try {
    const d = value instanceof Date ? value : new Date(value);
    return Utilities.formatDate(d, 'America/Fortaleza', 'dd/MM/yyyy HH:mm');
  } catch (_) {
    return '';
  }
}

function formatDateOnly_(value) {
  if (!value) return '';
  try {
    const d = value instanceof Date ? value : new Date(value);
    if (isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, 'America/Fortaleza', 'dd/MM/yyyy');
  } catch (_) {
    return '';
  }
}

function escapeRegExp_(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function validarCabecalhoPedidos_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  const headers = [
    'Pedido','Data/Hora','Público','Nome completo','Turma/Setor','WhatsApp','E-mail',
    'Código produto','Modelo','Tamanho','Nome na camisa','Cargo/Área/Função','Lote',
    'Valor bruto','Taxa plataforma','Valor líquido','Custo gráfica','Margem Grêmio',
    'Status pagamento','Status produção','Status entrega','ID cobrança Asaas',
    'Link pagamento','Data pagamento','Observações','Prazo previsto','Última atualização',
    'Atualizado por','Mensagem ao cliente','Observação interna','ID cliente Asaas',
    'Notificações enviadas','Método escolhido','Parcelas','Valor-base','Acréscimo pagamento',
    'Valor parcela','ID parcelamento Asaas'
  ];

  if (sh.getMaxColumns() < headers.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  }
  const current = sh.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  if (current.join('|') !== headers.join('|')) sh.getRange(1,1,1,headers.length).setValues([headers]);
}


function aplicarFormatacaoPedidos_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, 38)
    .setBackground('#063667').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  const widths = [
    125,145,110,220,150,130,180,120,205,85,170,205,120,105,105,105,
    105,110,170,170,185,175,250,145,240,130,145,140,260,260,175,220,
    150,85,110,125,110,190
  ];
  widths.forEach((px,i)=>sh.setColumnWidth(i+1,px));
  sh.getRange('N2:R3000').setNumberFormat('R$ #,##0.00');
  sh.getRange('AI2:AK3000').setNumberFormat('R$ #,##0.00');
  sh.getRange('B2:B3000').setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange('X2:X3000').setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange('Z2:Z3000').setNumberFormat('dd/MM/yyyy');
  sh.getRange('AA2:AA3000').setNumberFormat('dd/MM/yyyy HH:mm');
}


function aplicarValidacoesPedidos_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_PEDIDOS);
  const list = values => SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();

  sh.getRange('C2:C3000').setDataValidation(list(['ESTUDANTE', 'SERVIDOR']));
  sh.getRange('S2:S3000').setDataValidation(list([
    'CRIANDO COBRANÇA',
    'AGUARDANDO PAGAMENTO',
    'CONFIRMADO',
    'PAGO',
    'VENCIDO',
    'CANCELADO',
    'ESTORNADO',
    'EM ANÁLISE',
    'CONTESTADO',
    'ERRO NA COBRANÇA'
  ]));
  sh.getRange('T2:T3000').setDataValidation(list(['NÃO ENVIADO', 'ENVIADO À GRÁFICA', 'EM PRODUÇÃO', 'PRONTO']));
  sh.getRange('U2:U3000').setDataValidation(list(['NÃO ENTREGUE', 'DISPONÍVEL PARA RETIRADA', 'ENTREGUE']));
}