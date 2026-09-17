/*******************************************************
 * ROTA PADEL — BACKEND DO TORNEIO
 * Adicionar este arquivo ao projeto Apps Script APP-RTC.
 *
 * Usa a aba existente "RotaPadel" e cria:
 *   - Torneio_Config
 *   - Torneio_Jogos
 *   - Torneio_Avisos
 *
 * Formato de 6 duplas por categoria:
 *   QF1: 3 x 6
 *   QF2: 4 x 5
 *   SF1: 1 x vencedor QF2
 *   SF2: 2 x vencedor QF1
 *   FINAL: vencedor SF1 x vencedor SF2
 *******************************************************/

const RP_RELEASE_DEFAULT = '2026-10-10T14:00:00-03:00';
const RP_CATEGORIAS = [
  'Misto A','Misto B','Misto C',
  'Feminino A','Feminino B','Feminino C',
  'Masculino A','Masculino B','Masculino C'
];

function obterAbaTorneio_(nome, cabecalhos) {
  const ss = obterPlanilha();
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) sh.appendRow(cabecalhos);
  return sh;
}

function prepararTabelasRotaPadel_() {
  const cfg = obterAbaTorneio_('Torneio_Config',
    ['Chave','Valor']);
  if (cfg.getLastRow() === 1) {
    cfg.getRange(2,1,4,2).setValues([
      ['release','2026-10-10T14:00:00-03:00'],
      ['evento','Rota Padel'],
      ['local','Cruz Alta'],
      ['ativo','SIM']
    ]);
  }

  obterAbaTorneio_('Torneio_Jogos', [
    'ID','Categoria','Fase','Ordem','DataHora','Quadra',
    'Dupla1','Dupla2','Status','Placar1','Placar2',
    'Vencedor','FinalizadoEm'
  ]);

  obterAbaTorneio_('Torneio_Avisos', [
    'DataHora','Mensagem','Ativo'
  ]);
}

/**
 * Lê inscrições aceitas da aba RotaPadel.
 * Colunas existentes:
 * 1 ID, 2 Data/Hora, 3 P1, 4 Tel1, 5 Email1,
 * 6 P2, 7 Tel2, 8 Email2, 9 Modalidade, 10 Categoria,
 * 11 Aceite, 12 Status, 13 Comprovante, 14 Arquivo Drive
 */
function listarInscricoesTorneio_() {
  const sh = obterAbaRotaPadel();
  const last = sh.getLastRow();
  if (last < 2) return [];

  const rows = sh.getRange(2,1,last-1,14).getValues();
  return rows.map(r => ({
    protocolo: String(r[0] || '').trim(),
    p1: String(r[2] || '').trim(),
    p2: String(r[5] || '').trim(),
    modalidade: String(r[8] || '').trim(),
    categoria: String(r[9] || '').trim(),
    aceite: String(r[10] || '').trim(),
    status: String(r[11] || '').trim()
  })).filter(x => x.protocolo && x.p1 && x.p2);
}

function categoriaTorneio_(r) {
  const m = r.modalidade.toLowerCase();
  const c = r.categoria.toUpperCase();
  let modalidade = '';
  if (m.includes('misto')) modalidade = 'Misto';
  else if (m.includes('fem')) modalidade = 'Feminino';
  else if (m.includes('masc')) modalidade = 'Masculino';
  if (!modalidade || !['A','B','C'].includes(c)) return '';
  return modalidade + ' ' + c;
}

function inscricaoAceitaTorneio_(r) {
  const aceite = r.aceite.toLowerCase();
  const status = r.status.toLowerCase();
  return ['sim','aceito','aprovado','confirmado','pago'].some(x => aceite === x || status === x)
    || (aceite === '' && status === '');
}

/**
 * Gera/regera somente o chaveamento estrutural.
 * As duplas são preenchidas quando houver exatamente 6 inscritas aceitas.
 */
function gerarChaveamentoRotaPadel() {
  prepararTabelasRotaPadel_();

  const sh = SpreadsheetApp.getActive().getSheetByName('Torneio_Jogos');
  const inscricoes = listarInscricoesTorneio_()
    .filter(inscricaoAceitaTorneio_)
    .map(r => ({...r, chave: categoriaTorneio_(r)}))
    .filter(r => RP_CATEGORIAS.includes(r.chave));

  // Limpa jogos anteriores para gerar novamente de forma determinística.
  if (sh.getLastRow() > 1) {
    sh.getRange(2,1,sh.getLastRow()-1,13).clearContent();
  }

  const out = [];
  RP_CATEGORIAS.forEach(cat => {
    const duplas = inscricoes.filter(x => x.chave === cat).slice(0,6);

    if (duplas.length < 6) return;

    const nomes = duplas.map((d,i) =>
      'D' + (i+1) + ' — ' + d.p1 + ' / ' + d.p2
    );

    const base = cat.replace(/[^A-Za-z0-9]/g,'-').toUpperCase();

    out.push([
      base+'-QF1',cat,'Quartas',1,'','',
      nomes[2],nomes[5],'Agendado','','','',''
    ]);
    out.push([
      base+'-QF2',cat,'Quartas',2,'','',
      nomes[3],nomes[4],'Agendado','','','',''
    ]);
    out.push([
      base+'-SF1',cat,'Semifinal',3,'','',
      nomes[0],base+'-QF2','Agendado','','','',''
    ]);
    out.push([
      base+'-SF2',cat,'Semifinal',4,'','',
      nomes[1],base+'-QF1','Agendado','','','',''
    ]);
    out.push([
      base+'-FINAL',cat,'Final',5,'','',
      base+'-SF1',base+'-SF2','Agendado','','','',''
    ]);
  });

  if (out.length) {
    sh.getRange(2,1,out.length,13).setValues(out);
  }

  return {
    categorias: RP_CATEGORIAS.map(cat => ({
      categoria: cat,
      duplas: inscricoes.filter(x => x.chave === cat).length,
      chaveGerada: inscricoes.filter(x => x.chave === cat).length === 6
    })),
    jogosGerados: out.length
  };
}

/**
 * Publica somente os dados da categoria do protocolo.
 * Nunca devolve telefone/e-mail.
 */
function obterTorneioPorProtocolo(protocolo, adminSolicitado) {
  prepararTabelasRotaPadel_();

  const cfg = obterAbaTorneio_('Torneio_Config',['Chave','Valor']);
  const vals = cfg.getRange(2,1,Math.max(cfg.getLastRow()-1,0),2).getValues();
  const mapa = {};
  vals.forEach(r => mapa[String(r[0])] = String(r[1]));

  const release = mapa.release || RP_RELEASE_DEFAULT;
  const agora = new Date();
  const liberado = agora >= new Date(release);

  if (!adminSolicitado && !liberado) {
    return {
      ok:true,
      liberado:false,
      release:release
    };
  }

  const p = String(protocolo || '').trim();
  if (!p && !adminSolicitado) {
    return {ok:false, erro:'Informe o protocolo.'};
  }

  const inscricoes = listarInscricoesTorneio_()
    .filter(inscricaoAceitaTorneio_)
    .map(r => ({...r, chave: categoriaTorneio_(r)}));

  const minha = adminSolicitado
    ? null
    : inscricoes.find(r => r.protocolo.toLowerCase() === p.toLowerCase());

  if (!adminSolicitado && !minha) {
    return {ok:false, erro:'Protocolo não encontrado.'};
  }

  const categoria = minha ? minha.chave : '';
  const sh = obterAbaTorneio_('Torneio_Jogos',[
    'ID','Categoria','Fase','Ordem','DataHora','Quadra',
    'Dupla1','Dupla2','Status','Placar1','Placar2',
    'Vencedor','FinalizadoEm'
  ]);
  const last = sh.getLastRow();
  const jogos = last < 2 ? [] : sh.getRange(2,1,last-1,13).getValues();

  const publicos = jogos.map(r => ({
    id:String(r[0]||''),
    categoria:String(r[1]||''),
    fase:String(r[2]||''),
    ordem:Number(r[3]||0),
    dataHora:r[4] instanceof Date ? r[4].toISOString() : String(r[4]||''),
    quadra:String(r[5]||''),
    dupla1:String(r[6]||''),
    dupla2:String(r[7]||''),
    status:String(r[8]||''),
    placar1:String(r[9]||''),
    placar2:String(r[10]||''),
    vencedor:String(r[11]||'')
  })).filter(j => adminSolicitado || j.categoria === categoria);

  const resultados = publicos
    .filter(j => j.status === 'Finalizado')
    .map(j => ({
      id:j.id,categoria:j.categoria,fase:j.fase,
      dupla1:j.dupla1,dupla2:j.dupla2,
      placar1:j.placar1,placar2:j.placar2,
      vencedor:j.vencedor
    }));

  // Em "Agora em quadra" não há placar.
  const aoVivo = publicos
    .filter(j => j.status === 'Em jogo')
    .map(j => ({
      id:j.id,categoria:j.categoria,fase:j.fase,
      quadra:j.quadra,dupla1:j.dupla1,dupla2:j.dupla2
    }));

  return {
    ok:true,
    liberado:true,
    release:release,
    categoria:categoria,
    minhaDupla:minha ? {
      protocolo:minha.protocolo,
      participante1:minha.p1,
      participante2:minha.p2
    } : null,
    jogos:publicos,
    resultados:resultados,
    aoVivo:aoVivo
  };
}

/**
 * Endpoint para o frontend.
 *
 * GET:
 *   ?acao=torneio&protocolo=RP-XXXX
 *   ?acao=torneio&admin=1
 *
 * POST:
 *   {acao:"torneio", protocolo:"RP-XXXX"}
 *
 * Para não quebrar o doGet/doPost atual, chamar estas funções
 * de dentro do doGet/doPost existentes.
 */
function endpointTorneio_(params) {
  const protocolo = String((params && params.protocolo) || '').trim();
  const adminSolicitado = String((params && params.admin) || '') === '1';

  // Admin real deve ser validado pelo e-mail da conta executora.
  const email = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const adminAutorizado = adminSolicitado &&
    email === 'rotaractclubcruzalta@gmail.com';

  return obterTorneioPorProtocolo(
    adminAutorizado ? '' : protocolo,
    adminAutorizado
  );
}
