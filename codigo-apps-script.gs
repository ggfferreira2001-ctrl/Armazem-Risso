/**
 * CONTROLE DE ARMAZÉM — BACKEND (Google Apps Script)
 * ---------------------------------------------------
 * Atualizar: Extensões > Apps Script > cole > salve >
 * Implantar > Gerenciar implantações > lápis > Nova versão > Implantar
 */

const PASTA_FOTOS_NOME = 'Fotos Armazem Risso';

const COLUNAS_NOTAS = [
  'id','notaFiscal','status','descricao','conferente','dataHora',
  'localizacao','separacaoId','fotos','setor','peso','volumes',
  'conferidoPor','conferidoEm','dataUltimaMovimentacao',
  'dataAgendada','volumeTotal','volumesFaltando','volRecuperados','extravioLocalizado'
];
const COLUNAS_COMENTARIOS = ['id','notaId','notaFiscal','texto','conferente','dataHora','fotos'];
const COLUNAS_SEPARACOES  = ['id','titulo','conferente','dataHora'];
const COLUNAS_CHEGADAS    = ['id','codigo','conferente','dataHora'];
const COLUNAS_MOVIMENTOS  = ['id','tipo','notaFiscal','conferente','detalhe','dataHora'];
const COLUNAS_HISTORICO   = COLUNAS_NOTAS.concat(['separacaoTitulo','dataBaixa']);
const COLUNAS_INVENTARIOS     = ['id','titulo','conferente','dataHora','encerrado'];
const COLUNAS_INVENTARIO_NOTAS = ['id','inventarioId','notaFiscal','setor','volumes','fotos','status','conferente','dataHora'];
const COLUNAS_CONFERENCIAS    = ['id','inventarioId','setor','conferente','dataHora'];
const COLUNAS_CONF_ITENS      = ['id','conferenciaId','inventarioId','notaFiscal','resultado','setorRegistrado','setorEncontrado','conferente','dataHora'];

function getSheet(nome, colunas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(nome);
  if (!sheet) {
    sheet = ss.insertSheet(nome);
    sheet.appendRow(colunas);
    sheet.setFrozenRows(1);
    sheet.getRange('A:D').setNumberFormat('@');
  }
  return sheet;
}

function getPastaFotos() {
  const pastas = DriveApp.getFoldersByName(PASTA_FOTOS_NOME);
  if (pastas.hasNext()) return pastas.next();
  return DriveApp.createFolder(PASTA_FOTOS_NOME);
}

function doGet(e) {
  migrarColunas();
  var params = (e && e.parameter) ? e.parameter : {};

  // Busca de Histórico sob demanda (não entra na sincronização normal)
  if (params.historico === '1') {
    return responder({ historico: buscarHistorico(params) });
  }
  // Comentários de UMA nota específica (usado ao abrir detalhe no Histórico)
  if (params.comentariosDe) {
    return responder({ comentarios: buscarComentariosPorNota(params.comentariosDe) });
  }

  // notas de inventário específico sob demanda
  if (params.inventarioNotas) {
    var todas = lerLinhas(getSheet('InventarioNotas', COLUNAS_INVENTARIO_NOTAS), ['fotos']);
    return responder({ notas: todas.filter(function(n){ return String(n.inventarioId)===String(params.inventarioNotas); }) });
  }

  return responder({
    notas:        lerLinhas(getSheet('Notas',           COLUNAS_NOTAS),            ['fotos']),
    comentarios:  lerLinhas(getSheet('Comentarios',     COLUNAS_COMENTARIOS),      ['fotos']),
    separacoes:   lerLinhas(getSheet('Separacoes',      COLUNAS_SEPARACOES),       []),
    chegadas:     lerLinhas(getSheet('Chegadas',        COLUNAS_CHEGADAS),         []),
    movimentos:   lerLinhas(getSheet('Movimentos',      COLUNAS_MOVIMENTOS),       []),
    inventarios:  lerLinhas(getSheet('Inventarios',     COLUNAS_INVENTARIOS),      []),
    invNotas:     lerLinhas(getSheet('InventarioNotas', COLUNAS_INVENTARIO_NOTAS), ['fotos']),
    conferencias: lerLinhas(getSheet('Conferencias',    COLUNAS_CONFERENCIAS),     []),
    confItens:    lerLinhas(getSheet('ConfItens',       COLUNAS_CONF_ITENS),       [])
  });
}

// Filtra o Histórico no servidor (não traz a planilha inteira pro app)
function buscarHistorico(params) {
  var linhas = lerLinhas(getSheet('Historico', COLUNAS_HISTORICO), ['fotos']);
  var busca  = (params.busca || '').toLowerCase();
  var setor  = params.setor || '';
  var status = params.status || '';
  var de     = params.de || '';
  var ate    = params.ate || '';
  return linhas.filter(function(n) {
    if (busca && String(n.notaFiscal).toLowerCase().indexOf(busca) === -1) return false;
    if (setor && n.setor !== setor) return false;
    if (status && n.status !== status) return false;
    if (de && new Date(n.dataBaixa) < new Date(de)) return false;
    if (ate && new Date(n.dataBaixa) > new Date(ate + 'T23:59:59')) return false;
    return true;
  }).sort(function(a,b){ return new Date(b.dataBaixa) - new Date(a.dataBaixa); });
}

function buscarComentariosPorNota(notaId) {
  return lerLinhas(getSheet('Comentarios', COLUNAS_COMENTARIOS), ['fotos'])
    .filter(function(c){ return String(c.notaId) === String(notaId); });
}

// Garante que todas as colunas necessárias existem na planilha.
// Se uma coluna nova foi adicionada no código mas não existe ainda na aba,
// ela é criada no final sem apagar os dados existentes.
function migrarColunas() {
  var mapa = {
    'Notas':       COLUNAS_NOTAS,
    'Comentarios': COLUNAS_COMENTARIOS,
    'Separacoes':  COLUNAS_SEPARACOES,
    'Chegadas':    COLUNAS_CHEGADAS,
    'Movimentos':  COLUNAS_MOVIMENTOS,
    'Historico':   COLUNAS_HISTORICO,
    'Inventarios':     COLUNAS_INVENTARIOS,
    'InventarioNotas': COLUNAS_INVENTARIO_NOTAS,
    'Conferencias':    COLUNAS_CONFERENCIAS,
    'ConfItens':       COLUNAS_CONF_ITENS
  };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(mapa).forEach(function(nome) {
    var colunasEsperadas = mapa[nome];
    var sheet = ss.getSheetByName(nome);
    if (!sheet) return; // será criada quando necessário
    var cabecalhoAtual = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    colunasEsperadas.forEach(function(col) {
      if (!cabecalhoAtual.includes(col)) {
        var novaCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, novaCol).setValue(col);
        cabecalhoAtual.push(col);
      }
    });
  });
}

function lerLinhas(sheet, colsLista) {
  const dados = sheet.getDataRange().getValues();
  if (dados.length < 2) return [];
  const cab = dados[0];
  return dados.slice(1).filter(l => l[0]).map(linha => {
    const obj = {};
    cab.forEach((k,i) => obj[k] = linha[i]);
    colsLista.forEach(col => {
      if (obj[col] !== undefined)
        obj[col] = obj[col] ? String(obj[col]).split('|').filter(Boolean) : [];
    });
    ['dataHora','conferidoEm','dataUltimaMovimentacao'].forEach(k => {
      if (obj[k] instanceof Date) obj[k] = obj[k].toISOString();
    });
    return obj;
  });
}

function registrarMovimento(tipo, notaFiscal, conferente, detalhe) {
  getSheet('Movimentos', COLUNAS_MOVIMENTOS).appendRow([
    'mv_' + Date.now(), tipo, String(notaFiscal), conferente || '', detalhe || '', new Date().toISOString()
  ]);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const acao = body.acao;

    if (acao === 'adicionarNota') {
      const d = body.dados;
      const fotos = subirFotos(d.fotos || [], d.notaFiscal);
      const agora = d.dataHora || new Date().toISOString();
      appendRowPorNome(getSheet('Notas', COLUNAS_NOTAS), {
        id: d.id, notaFiscal: String(d.notaFiscal), status: d.status || '', descricao: d.descricao || '',
        conferente: d.conferente || '', dataHora: agora, localizacao: d.localizacao || 'notas',
        separacaoId: d.separacaoId || '', fotos: fotos.join('|'), setor: d.setor || '',
        peso: d.peso || '', volumes: d.volumes || '', conferidoPor: '', conferidoEm: '',
        dataUltimaMovimentacao: agora
      });
      registrarMovimento('nota_adicionada', d.notaFiscal, d.conferente, 'Nota adicionada' + (d.status ? ' como ' + d.status : ''));
      return responder({ ok: true });
    }

    if (acao === 'atualizarNota') {
      const campos = body.campos || {};
      if (!campos.dataUltimaMovimentacao) campos.dataUltimaMovimentacao = new Date().toISOString();
      atualizarCampos('Notas', body.id, campos);
      if (body.movTipo) registrarMovimento(body.movTipo, body.notaFiscal || '', body.conferente || '', body.movDetalhe || '');
      return responder({ ok: true });
    }

    if (acao === 'adicionarComentario') {
      const d = body.dados;
      const fotos = subirFotos(d.fotos || [], d.notaFiscal);
      appendRowPorNome(getSheet('Comentarios', COLUNAS_COMENTARIOS), {
        id: d.id, notaId: String(d.notaId), notaFiscal: String(d.notaFiscal),
        texto: d.texto || '', conferente: d.conferente || '', dataHora: d.dataHora, fotos: fotos.join('|')
      });
      atualizarCampos('Notas', d.notaId, { dataUltimaMovimentacao: d.dataHora });
      registrarMovimento('comentario', d.notaFiscal, d.conferente, 'Comentário: ' + (d.texto || '').slice(0,60));
      return responder({ ok: true });
    }

    if (acao === 'criarSeparacao') {
      const d = body.dados;
      appendRowPorNome(getSheet('Separacoes', COLUNAS_SEPARACOES), {
        id: d.id, titulo: d.titulo, conferente: d.conferente || '', dataHora: d.dataHora
      });
      registrarMovimento('separacao_criada', '', d.conferente, 'Separação criada: ' + d.titulo);
      return responder({ ok: true });
    }

    if (acao === 'moverParaSeparacao') {
      atualizarCampos('Notas', body.notaId, {
        localizacao: 'separacao', separacaoId: body.separacaoId,
        dataUltimaMovimentacao: new Date().toISOString()
      });
      registrarMovimento('nota_movida', body.notaFiscal || '', body.conferente || '', 'Movida para separação ' + (body.sepTitulo || ''));
      return responder({ ok: true });
    }

    if (acao === 'removerDeSeparacao') {
      atualizarCampos('Notas', body.notaId, {
        localizacao: 'notas', separacaoId: '',
        dataUltimaMovimentacao: new Date().toISOString()
      });
      registrarMovimento('nota_removida_sep', body.notaFiscal || '', body.conferente || '', 'Removida de separação');
      return responder({ ok: true });
    }

    if (acao === 'conferirNota') {
      atualizarCampos('Notas', body.notaId, {
        conferidoPor: body.conferente, conferidoEm: body.dataHora,
        dataUltimaMovimentacao: body.dataHora
      });
      registrarMovimento('conferida', body.notaFiscal || '', body.conferente || '', 'Nota conferida');
      return responder({ ok: true });
    }

    if (acao === 'excluirNota') {
      excluirPorId('Notas', body.id);
      excluirPorColuna('Comentarios', 'notaId', body.id);
      registrarMovimento('nota_excluida', body.notaFiscal || '', body.conferente || '', 'Nota excluída');
      return responder({ ok: true });
    }

    if (acao === 'baixarSeparacao') {
      var sheetNotas = getSheet('Notas', COLUNAS_NOTAS);
      var sheetHist  = getSheet('Historico', COLUNAS_HISTORICO);
      var dadosN = sheetNotas.getDataRange().getValues();
      var cabN = dadosN[0];
      var colSepId = cabN.indexOf('separacaoId');
      var agora = new Date().toISOString();
      var linhasParaExcluir = [];
      var qtd = 0;

      for (var i = 1; i < dadosN.length; i++) {
        if (String(dadosN[i][colSepId]) === String(body.id)) {
          // monta objeto {nomeColuna: valor} lendo pelo cabeçalho REAL da planilha de Notas
          var objLinha = {};
          cabN.forEach(function(col, idx) { objLinha[col] = dadosN[i][idx]; });
          objLinha.separacaoTitulo = body.sepTitulo || '';
          objLinha.dataBaixa = agora;
          appendRowPorNome(sheetHist, objLinha);
          linhasParaExcluir.push(i + 1);
          qtd++;
        }
      }
      linhasParaExcluir.sort(function(a,b){ return b - a; });
      linhasParaExcluir.forEach(function(linha) { sheetNotas.deleteRow(linha); });

      excluirPorId('Separacoes', body.id);
      registrarMovimento('separacao_baixada', '', body.conferente || '', (body.sepTitulo || '') + ' — ' + qtd + ' nota(s) para o Histórico');
      return responder({ ok: true, quantidade: qtd });
    }

    if (acao === 'excluirSeparacao') {
      const sheet = getSheet('Notas', COLUNAS_NOTAS);
      const dados = sheet.getDataRange().getValues();
      const cab = dados[0];
      const colSep = cab.indexOf('separacaoId');
      const colLoc = cab.indexOf('localizacao');
      for (let i = 1; i < dados.length; i++) {
        if (String(dados[i][colSep]) === String(body.id)) {
          sheet.getRange(i+1, colSep+1).setValue('');
          sheet.getRange(i+1, colLoc+1).setValue('notas');
        }
      }
      excluirPorId('Separacoes', body.id);
      registrarMovimento('separacao_excluida', '', body.conferente || '', 'Separação excluída: ' + (body.titulo || ''));
      return responder({ ok: true });
    }

    if (acao === 'registrarChegada') {
      const d = body.dados;
      appendRowPorNome(getSheet('Chegadas', COLUNAS_CHEGADAS), {
        id: d.id, codigo: String(d.codigo), conferente: d.conferente || '', dataHora: d.dataHora
      });
      return responder({ ok: true });
    }

    if (acao === 'excluirChegada') {
      excluirPorId('Chegadas', body.id);
      return responder({ ok: true });
    }

    if (acao === 'salvarConferencia') {
      var d = body.dados;
      // cabeçalho da conferência
      appendRowPorNome(getSheet('Conferencias', COLUNAS_CONFERENCIAS), {
        id: d.id, inventarioId: d.inventarioId, setor: d.setor,
        conferente: d.conferente || '', dataHora: d.dataHora
      });
      // itens da conferência
      var sheetItens = getSheet('ConfItens', COLUNAS_CONF_ITENS);
      (d.itens || []).forEach(function(item) {
        appendRowPorNome(sheetItens, {
          id: item.id, conferenciaId: d.id, inventarioId: d.inventarioId,
          notaFiscal: String(item.notaFiscal), resultado: item.resultado,
          setorRegistrado: item.setorRegistrado || '', setorEncontrado: item.setorEncontrado || '',
          conferente: d.conferente || '', dataHora: d.dataHora
        });
      });
      registrarMovimento('conferencia_setor', '', d.conferente, 'Conferência do ' + d.setor + ': ' + (d.itens||[]).length + ' notas');
      return responder({ ok: true });
    }

    if (acao === 'excluirConferencia') {
      excluirPorId('Conferencias', body.id);
      // remove todos os itens vinculados
      var sheetIt = getSheet('ConfItens', COLUNAS_CONF_ITENS);
      var dados = sheetIt.getDataRange().getValues();
      var cab = dados[0];
      var colConf = cab.indexOf('conferenciaId');
      var linhasRem = [];
      for (var i = 1; i < dados.length; i++) {
        if (String(dados[i][colConf]) === String(body.id)) linhasRem.push(i+1);
      }
      linhasRem.sort(function(a,b){return b-a;}).forEach(function(l){sheetIt.deleteRow(l);});
      return responder({ ok: true });
    }

    if (acao === 'criarInventario') {
      var d = body.dados;
      appendRowPorNome(getSheet('Inventarios', COLUNAS_INVENTARIOS), {
        id: d.id, titulo: d.titulo, conferente: d.conferente || '', dataHora: d.dataHora, encerrado: 'false'
      });
      registrarMovimento('inventario_criado', '', d.conferente, 'Inventário criado: ' + d.titulo);
      return responder({ ok: true });
    }

    if (acao === 'encerrarInventario') {
      atualizarCampos('Inventarios', body.id, { encerrado: 'true' });
      registrarMovimento('inventario_encerrado', '', body.conferente, 'Inventário encerrado');
      return responder({ ok: true });
    }

    if (acao === 'reabrirInventario') {
      atualizarCampos('Inventarios', body.id, { encerrado: 'false' });
      registrarMovimento('inventario_reaberto', '', body.conferente, 'Inventário reaberto (adm)');
      return responder({ ok: true });
    }

    if (acao === 'adicionarInvNota') {
      var d = body.dados;
      var fotos = subirFotos(d.fotos || [], d.notaFiscal);
      appendRowPorNome(getSheet('InventarioNotas', COLUNAS_INVENTARIO_NOTAS), {
        id: d.id, inventarioId: d.inventarioId, notaFiscal: String(d.notaFiscal),
        setor: d.setor || '', volumes: d.volumes || '', fotos: fotos.join('|'),
        status: d.status || '', conferente: d.conferente || '', dataHora: d.dataHora
      });
      registrarMovimento('inv_nota_adicionada', d.notaFiscal, d.conferente, 'Adicionada ao inventário');
      return responder({ ok: true });
    }

    if (acao === 'atualizarInvNota') {
      atualizarCampos('InventarioNotas', body.id, body.campos);
      return responder({ ok: true });
    }

    if (acao === 'excluirInvNota') {
      excluirPorId('InventarioNotas', body.id);
      return responder({ ok: true });
    }

    if (acao === 'restaurarDoHistorico') {
      var sheetHist  = getSheet('Historico', COLUNAS_HISTORICO);
      var sheetNotas = getSheet('Notas', COLUNAS_NOTAS);
      var dadosH = sheetHist.getDataRange().getValues();
      var cabH = dadosH[0];
      var colId = cabH.indexOf('id');
      for (var i = 1; i < dadosH.length; i++) {
        if (String(dadosH[i][colId]) === String(body.id)) {
          var obj = {};
          cabH.forEach(function(col,idx){ obj[col] = dadosH[i][idx]; });
          obj.localizacao = 'notas';
          obj.separacaoId = '';
          delete obj.separacaoTitulo;
          delete obj.dataBaixa;
          appendRowPorNome(sheetNotas, obj);
          sheetHist.deleteRow(i + 1);
          registrarMovimento('nota_restaurada', obj.notaFiscal, body.conferente || '', 'Restaurada do Histórico');
          return responder({ ok: true });
        }
      }
      return responder({ ok: false, erro: 'Nota não encontrada no Histórico' });
    }

    return responder({ ok: false, erro: 'Ação desconhecida: ' + acao });
  } catch(err) {
    return responder({ ok: false, erro: String(err) });
  }
}

function subirFotos(fotos, notaFiscal) {
  if (!fotos || fotos.length === 0) return [];
  const pasta = getPastaFotos();
  return fotos.map(base64 => {
    try {
      const partes = base64.split(',');
      const tipo = (partes[0].match(/data:(.*);base64/)||[])[1] || 'image/jpeg';
      const bytes = Utilities.base64Decode(partes[1]);
      const blob = Utilities.newBlob(bytes, tipo, 'nota_'+notaFiscal+'_'+Date.now()+'.jpg');
      const arq = pasta.createFile(blob);
      arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return 'https://drive.google.com/thumbnail?id='+arq.getId()+'&sz=w1000';
    } catch(e) { return ''; }
  }).filter(Boolean);
}

// Escreve uma linha nova casando cada valor pelo NOME da coluna (lido ao vivo
// do cabeçalho real da planilha), não por posição fixa. Isso torna a escrita
// imune a qualquer desalinhamento de colunas que possa ter acontecido no
// histórico da planilha (colunas adicionadas em momentos diferentes).
function appendRowPorNome(sheet, dadosObj) {
  var cab = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var linha = cab.map(function(col) {
    var v = dadosObj[col];
    return (v === undefined || v === null) ? '' : v;
  });
  sheet.appendRow(linha);
}

function atualizarCampos(sheetNome, id, campos) {
  const sheet = getSheet(sheetNome, []);
  const dados = sheet.getDataRange().getValues();
  const cab = dados[0];
  const colId = cab.indexOf('id');
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][colId]) === String(id)) {
      Object.entries(campos).forEach(([k,v]) => {
        const col = cab.indexOf(k);
        if (col >= 0) sheet.getRange(i+1, col+1).setValue(v);
      });
      break;
    }
  }
}

function excluirPorId(sheetNome, id) {
  const sheet = getSheet(sheetNome, []);
  const dados = sheet.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(id)) { sheet.deleteRow(i+1); return; }
  }
}

function excluirPorColuna(sheetNome, colNome, valor) {
  const sheet = getSheet(sheetNome, []);
  const dados = sheet.getDataRange().getValues();
  const col = dados[0].indexOf(colNome);
  if (col < 0) return;
  for (let i = dados.length-1; i >= 1; i--) {
    if (String(dados[i][col]) === String(valor)) sheet.deleteRow(i+1);
  }
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Diagnóstico opcional: rode manualmente (Executar > diagnosticarColunas) pra
// conferir se a ordem real das colunas da planilha bate com o que o código espera.
// Ver o resultado em: Ver > Registros de execução (Execution log).
function diagnosticarColunas() {
  var sheet = getSheet('Notas', COLUNAS_NOTAS);
  var cabReal = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('Colunas REAIS na planilha Notas: ' + cabReal.join(' | '));
  Logger.log('Colunas ESPERADAS pelo código:   ' + COLUNAS_NOTAS.join(' | '));
}
