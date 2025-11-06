/************** Code.gs — versão para index com modal LGPD (com patches) **************/

/// ---------- Configurações principais ----------
const SHEET_NAME = 'BASE DE DADOS CADASTRAIS';
const NOME_EMPRESA_DEFAULT = 'RC Beauty';
const COORD_DEFAULT = 'Darliany Kamila Oliveira';
const ASSUNTO_EMAIL = 'Inscrição recebida — Programa de Certificações RC Beauty';

// ID padrão da planilha (onde está o formulário)
const SID_DEFAULT = '11qKd4sc7laUyFdPWG89NgLagrqH9J6DcCyy3i4vYA_0';

// ---------- DESTINO ESPELHO (onde cai a cópia) ----------
const DEST_SPREADSHEET_ID = '1OmUrifj2_y-Kzk3sC3tg6SyzrrJLDCU0GtHg--4nUzU';
const DEST_SHEET_NAME     = 'BASE DE DADOS CADASTRAIS';

// ---------- Cabeçalho oficial ----------
const STANDARD_HEADER = [
  'CARIMBO DE DATA/HORA','CURSO','LOCAL DO EVENTO','CICLO','STATUS',
  'NOME COMPLETO SEM ABREVIAÇÕES','CPF','E-MAIL','TEL/WHATSAPP',
  'ENDEREÇO','NÚMERO','COMPLEMENTO','BAIRRO','CEP','CIDADE','ESTADO','PAÍS',
  'PROFISSÃO','ESCOLARIDADE','GRADUAÇÃO',
  // Novos campos LGPD/Consentimentos
  'LGPD_VERSION','LGPD_TS','LGPD_IP','OPT-IN','CONSENTIMENTO DE IMAGEM'
];

/// ---------- Helpers ----------
const sanitize_ = s => (s || '').toString().trim();
const onlyDigits_ = s => sanitize_(s).replace(/\D+/g, '');
const norm_ = s => sanitize_(s).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();

function isValidEmail_(email){ return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/i.test(email); }
function isValidCEP_(cep){ return /^[0-9]{8}$/.test(onlyDigits_(cep)); }
function isValidUF_(uf){ return /^(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)$/i.test(sanitize_(uf)); }
function isValidPhoneBR_(phone){ const d=onlyDigits_(phone); return d.length>=10 && d.length<=13; }
function isValidCPF_(cpf){
  const c=onlyDigits_(cpf);
  if(!/^\d{11}$/.test(c)) return false;
  if(/^(\d)\1{10}$/.test(c)) return false;
  let s=0; for(let i=0;i<9;i++) s+= +c[i]*(10-i);
  let dv1=11-(s%11); dv1=dv1>9?0:dv1; if(dv1!== +c[9]) return false;
  s=0; for(let i=0;i<10;i++) s+= +c[i]*(11-i);
  let dv2=11-(s%11); dv2=dv2>9?0:dv2; return dv2=== +c[10];
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function formatCPF_(c) {
  const d = onlyDigits_(c);
  if (d.length !== 11) return c || '';
  return d.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
}
function maskCPF_(c) {
  const d = onlyDigits_(c);
  if (d.length !== 11) return c || '';
  return `***.***.***-${d.slice(-2)}`;
}

/// ---------- Planilha base + cabeçalho (WRITE) ----------
function getSheet_(sid, sheetName = SHEET_NAME) {
  const id = String(sid || '').trim();
  if (!id) throw new Error('Faltou o parâmetro "sid".');
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName); // cria em operações de escrita
  ensureStandardHeader_(sh);
  return sh;
}

/// ---------- Planilha apenas leitura (READ, NÃO CRIA) ----------
function getSheetForRead_(sid, sheetName = SHEET_NAME) {
  const id = String(sid || '').trim();
  if (!id) throw new Error('Faltou o parâmetro "sid".');
  const ss = SpreadsheetApp.openById(id);
  return ss.getSheetByName(sheetName) || null; // não cria
}

function ensureStandardHeader_(sh) {
  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  let header = [];
  if (lastCol >= 1 && lastRow >= 1) {
    header = sh.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  }
  const hasValues = header.some(v => String(v || '').trim() !== '');
  if (!hasValues) {
    if (sh.getMaxColumns() < STANDARD_HEADER.length) {
      sh.insertColumnsAfter(sh.getMaxColumns() || 1, STANDARD_HEADER.length - (sh.getMaxColumns() || 1));
    }
    sh.getRange(1, 1, 1, STANDARD_HEADER.length).setValues([STANDARD_HEADER]);
    sh.setFrozenRows(1);
    return;
  }
  const existingNorm = header.map(h => norm_(h));
  STANDARD_HEADER.forEach(hWanted => {
    const wantedNorm = norm_(hWanted);
    if (!existingNorm.includes(wantedNorm)) {
      sh.insertColumnAfter(sh.getLastColumn());
      sh.getRange(1, sh.getLastColumn()).setValue(hWanted);
      existingNorm.push(wantedNorm);
    }
  });
  sh.setFrozenRows(1);
}

function runEnsureHeaderMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  ensureStandardHeader_(sh);
  SpreadsheetApp.getUi().alert('Cabeçalho verificado/ajustado.');
}

/// ---------- Map dinâmico ----------
function headerMap_(header){
  const clean = s => (s || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
  const H = header.map(clean);
  const T = H.map(h => h.split(' ').filter(Boolean));
  const hasPhrase = (h, phrase) => !!phrase && (` ${h} `).includes(` ${phrase} `);
  const findSmart = (...keys) => {
    const K = keys.map(k => clean(k)).filter(Boolean);
    return H.findIndex((h, i) => {
      const tokens = T[i];
      return K.some(k => hasPhrase(h,k) || (k.split(' ').every(w => tokens.includes(w))));
    });
  };
  const map = {};
  map.timestamp   = findSmart('carimbo de data hora','timestamp');
  map.curso       = findSmart('curso');
  map.local       = findSmart('local do evento','local evento','cidade do curso','cidade que fara o curso');
  map.ciclo       = findSmart('ciclo');
  map.status      = findSmart('status');
  map.nome        = findSmart('nome completo sem abreviacoes','nome completo');
  map.cpf         = findSmart('cpf');
  map.email       = findSmart('e mail','email','e-mail');
  map.whatsapp    = findSmart('tel whatsapp','whatsapp','telefone','tel');
  map.endereco    = findSmart('endereco','endereço','endereco logradouro');
  map.numero      = findSmart('numero');
  map.complemento = findSmart('complemento');
  map.bairro      = findSmart('bairro');
  map.cep         = findSmart('cep');
  map.cidade      = findSmart('cidade');
  map.estado      = findSmart('estado','uf','sigla do estado');
  map.pais        = findSmart('pais');
  map.profissao   = findSmart('profissao');
  map.escolaridade= findSmart('escolaridade');
  map.graduacao   = findSmart('graduacao','curso academico','curso superior area','area de formacao','formacao curso');

  // Novos campos LGPD/Consentimentos
  map.lgpdVersion = findSmart('lgpd version','lgpd_version');
  map.lgpdTs      = findSmart('lgpd ts','lgpd timestamp','lgpd data hora');
  map.lgpdIp      = findSmart('lgpd ip','ip');
  map.optin       = findSmart('opt in','opt-in','marketing');
  map.consentImg  = findSmart('consentimento de imagem','consent imagem','uso de imagem','imagem voz');

  return map;
}

/// ---------- doGet ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  const sidParam = (p.sid || '').trim();
  const sid = sidParam || SID_DEFAULT;
  const sheet = (p.sheet || '').trim() || SHEET_NAME;

  // --- página principal (index com modal) ---
  const tpl = HtmlService.createTemplateFromFile('index');

  const defaultCourses = [
    'FORMAÇÃO BÁSICA EM TERAPIA CAPILAR',
    'FORMAÇÃO AVANÇADA EM TERAPIA CAPILAR',
    'PÓS-GRADUAÇÃO EM TRICOLOGIA E CIÊNCIA COSMÉTICA',
  ];

  function parseCursosParam_(raw) {
    const s = String(raw || '').trim();
    if (!s) return null;
    if (s.startsWith('[') && s.endsWith(']')) {
      try {
        const jsonTxt = (s.includes("'") && !s.includes('"')) ? s.replace(/'/g,'"') : s;
        const arr = JSON.parse(jsonTxt);
        if (Array.isArray(arr)) return arr.map(x=>String(x).trim()).filter(Boolean);
      } catch(_) {}
    }
    if (s.includes('|') || s.includes(',')) return s.split(/[|,]/).map(x=>x.trim()).filter(Boolean);
    return [s];
  }

  const cursosParam = parseCursosParam_(p.cursos);
  const cursosList  = (Array.isArray(cursosParam) && cursosParam.length) ? cursosParam : defaultCourses;

  // Injeta variáveis do servidor
  // IMPORTANTE: aqui deixamos como array; o index.html fará JSON.stringify(cursos)
  tpl.cursos    = cursosList;
  tpl.sid       = sid;
  tpl.sheet     = sheet;
  tpl.pageTitle = p.title || `Cadastro de Aluno - ${NOME_EMPRESA_DEFAULT}`;
  tpl.empresa   = p.empresa || NOME_EMPRESA_DEFAULT;
  tpl.coord     = p.coord || COORD_DEFAULT;
  tpl.header    = p.header || '';

  return tpl.evaluate()
            .setTitle(tpl.pageTitle)
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/// ---------- CEP ----------
function cepLookup(cepRaw){
  const cep = onlyDigits_(cepRaw);
  if (!/^\d{8}$/.test(cep)) {
    return { ok:false, message:'CEP inválido (use 8 dígitos).' };
  }

  // 1) ViaCEP
  try {
    const url = `https://viacep.com.br/ws/${cep}/json/`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions:true, followRedirects: true });
    if (res.getResponseCode() === 200) {
      const data = JSON.parse(res.getContentText() || '{}');
      if (!data.erro) {
        return {
          ok: true,
          data: {
            logradouro: data.logradouro || '',
            bairro:     data.bairro     || '',
            cidade:     data.localidade || '',
            uf:         (data.uf || '').toUpperCase(),
            pais:       'Brasil'
          }
        };
      }
    }
  } catch(_) {}

  // 2) BrasilAPI (fallback)
  try {
    const url2 = `https://brasilapi.com.br/api/cep/v2/${cep}`;
    const res2 = UrlFetchApp.fetch(url2, { muteHttpExceptions:true, followRedirects: true });
    if (res2.getResponseCode() === 200) {
      const b = JSON.parse(res2.getContentText() || '{}');
      return {
        ok: true,
        data: {
          logradouro: b.street || '',
          bairro:     b.neighborhood || '',
          cidade:     b.city || '',
          uf:         (b.state || '').toUpperCase(),
          pais:       'Brasil'
        }
      };
    }
    if (res2.getResponseCode() === 404) {
      return { ok:false, message:'CEP não encontrado em nenhuma base.' };
    }
  } catch(_) {}

  return { ok:false, message:'Não foi possível obter o endereço no momento.' };
}

/// ---------- Salvar ----------
function salvarInscricao(dados) {
  try {
    if (!dados) throw new Error('Nenhum dado recebido.');

    const sid           = sanitize_(dados.sid);
    const sheetName     = sanitize_(dados.sheet) || SHEET_NAME;

    const curso         = sanitize_(dados.curso);
    const localEvento   = sanitize_(dados.localEvento);
    const ciclo         = sanitize_(dados.ciclo);
    const status        = sanitize_(dados.status) || '';

    const nomeCompleto  = sanitize_(dados.nome);
    const cpf           = onlyDigits_(dados.cpf);
    const email         = sanitize_(dados.email).toLowerCase();
    const whatsapp      = onlyDigits_(dados.whatsapp);

    const endereco      = sanitize_(dados.endereco);
    const numero        = sanitize_(dados.numero);
    const complemento   = sanitize_(dados.complemento);
    const bairro        = sanitize_(dados.bairro);
    const cep           = onlyDigits_(dados.cep);
    const cidade        = sanitize_(dados.cidade);
    const estado        = sanitize_(dados.estado).toUpperCase();
    const pais          = sanitize_(dados.pais) || 'Brasil';

    const profissao     = sanitize_(dados.profissao);
    const escolaridade  = sanitize_(dados.escolaridade);
    const graduacao     = sanitize_(dados.graduacao);

    // LGPD + consentimentos (novos campos)
    const lgpdVersion   = sanitize_(dados.lgpdVersion);
    const lgpdTs        = sanitize_(dados.lgpdTs);
    const lgpdIp        = sanitize_(dados.lgpdIp);
    const optin         = String(dados.optin) === 'on' ? 'SIM' : '';
    const consentImagem = String(dados.consentImagem) === 'on' ? 'SIM' : '';

    // LGPD obrigatório (checkbox principal)
    const lgpdOk = String(dados.lgpd) === 'on';
    if (!lgpdOk) throw new Error('Você precisa aceitar a Política de Privacidade (LGPD) para continuar.');

    // Validações
    if (!curso) throw new Error('Informe o curso.');
    if (!localEvento) throw new Error('Informe o local do evento.');
    if (!nomeCompleto) throw new Error('Informe seu nome completo.');
    if (!isValidCPF_(cpf)) throw new Error('CPF inválido.');
    if (!isValidEmail_(email)) throw new Error('E-mail inválido.');
    if (!isValidPhoneBR_(whatsapp)) throw new Error('WhatsApp inválido.');
    if (!isValidCEP_(cep)) throw new Error('CEP inválido.');
    if (!endereco) throw new Error('Informe o endereço (logradouro).');
    if (!numero) throw new Error('Informe o número do endereço.');
    if (!cidade) throw new Error('Informe a cidade.');
    if (!isValidUF_(estado)) throw new Error('Estado inválido.');
    if (!profissao) throw new Error('Informe sua profissão.');
    if (!escolaridade) throw new Error('Selecione sua escolaridade.');

    const exigeGraduacao = ['Ensino Superior Completo (Graduação)','Pós-graduação','Mestrado','Doutorado'].includes(escolaridade);
    if (exigeGraduacao && !graduacao) throw new Error('Informe sua Graduação (curso superior).');

    // Planilha e mapa
    const sh = getSheet_(sid, sheetName);
    const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] || [];
    const map = headerMap_(header);

    if (map.email === -1 || map.cpf === -1 || map.curso === -1) {
      throw new Error('Colunas essenciais ausentes: E-MAIL, CPF ou CURSO.');
    }

    let values = [];
    if (sh.getLastRow() > 1) {
      values = sh.getRange(2,1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
    }

    // Regra de duplicidade
    const eq = (a,b) => String(a||'').trim().toLowerCase() === String(b||'').trim().toLowerCase();
    const sameInscricaoExists = values.some(r => {
      const alunoIgual = eq(r[map.email], email) || (onlyDigits_(r[map.cpf]) === cpf);
      const cursoIgual = eq(r[map.curso], curso);
      const cicloIgual = (map.ciclo > -1) ? eq(r[map.ciclo], ciclo) : false;
      const localIgual = (map.local > -1) ? eq(r[map.local], localEvento) : false;
      const mesmoEvento = ciclo ? cicloIgual : localIgual;
      return alunoIgual && cursoIgual && (cicloIgual || localIgual || mesmoEvento);
    });

    if (sameInscricaoExists) {
      if (String(dados.atualizarDados) === 'on') {
        atualizarDadosAluno_(sh, header, values, {
          cpf, email, whatsapp, endereco, numero, complemento, bairro, cep, cidade, estado, pais,
          profissao, escolaridade, graduacao, nomeCompleto, localEvento, ciclo, status,
          // novos
          lgpdVersion, lgpdTs, lgpdIp, optin, consentImagem
        });
        return { ok: true, message: 'Dados atualizados para inscrições existentes deste aluno.' };
      }
      throw new Error('Já existe uma inscrição deste aluno para este evento (mesmo curso e mesmo ciclo/local).');
    }

    // Monta linha
    const row = new Array(sh.getLastColumn()).fill('');
    const put = (idx, val) => { if (idx !== -1 && idx !== undefined) row[idx] = val; };

    if (map.timestamp === 0 || /carimbo de data\/hora/i.test(String(header[0] || ''))) {
      row[0] = new Date();
    }

    put(map.curso, curso);
    put(map.local, localEvento);
    put(map.ciclo, ciclo);
    put(map.status, status);
    put(map.nome, nomeCompleto);
    put(map.cpf, cpf);
    put(map.email, email);
    put(map.whatsapp, whatsapp);
    put(map.endereco, endereco);
    put(map.numero, numero);
    put(map.complemento, complemento);
    put(map.bairro, bairro);
    put(map.cep, cep);
    put(map.cidade, cidade);
    put(map.estado, estado);
    put(map.pais, pais);
    put(map.profissao, profissao);
    put(map.escolaridade, escolaridade);
    put(map.graduacao, exigeGraduacao ? graduacao : '');

    // Novos campos LGPD/Consentimentos
    put(map.lgpdVersion, lgpdVersion);
    put(map.lgpdTs,      lgpdTs);
    put(map.lgpdIp,      lgpdIp);
    put(map.optin,       optin);
    put(map.consentImg,  consentImagem);

    sh.appendRow(row);

    // Espelho (não impede sucesso se falhar)
    try { mirrorToSecondary_(header, row); } catch(e) { Logger.log('Falha ao espelhar: ' + e); }

    // E-mail (opcional — não bloqueia)
    try {
      const templateEmail = HtmlService.createHtmlOutputFromFile('email').getContent();
      const cpfFormatado = formatCPF_(cpf);
      const cpfMask = maskCPF_(cpf);

      const corpo = templateEmail
        .replace(/{{NOME}}/g, escapeHtml_(nomeCompleto))
        .replace(/{{NOME_COMPLETO}}/g, escapeHtml_(nomeCompleto))
        .replace(/{{CPF_FORMATADO}}/g, escapeHtml_(cpfFormatado))
        .replace(/{{CPF_MASK}}/g, escapeHtml_(cpfMask))
        .replace(/{{CURSO}}/g, escapeHtml_(curso))
        .replace(/{{LOCAL_EVENTO}}/g, escapeHtml_(localEvento))
        .replace(/{{CICLO}}/g, escapeHtml_(ciclo || '—'))
        .replace(/{{STATUS}}/g, escapeHtml_(status || ''))
        .replace(/{{GRADUACAO}}/g, escapeHtml_(exigeGraduacao ? graduacao : ''))
        .replace(/{{NOME_EMPRESA}}/g, escapeHtml_(sanitize_(dados.empresa) || NOME_EMPRESA_DEFAULT));

      GmailApp.sendEmail(email, ASSUNTO_EMAIL, ' ', {
        name: sanitize_(dados.empresa) || NOME_EMPRESA_DEFAULT,
        htmlBody: corpo
      });
    } catch(e) { Logger.log('Falha ao enviar e-mail: ' + e); }

    return { ok: true, message: 'Inscrição registrada com sucesso!' };

  } catch (err) {
    Logger.log('ERRO em salvarInscricao: ' + (err.stack || err));
    return { ok: false, message: err && err.message ? err.message : 'Ocorreu um erro no servidor.' };
  }
}

/// ---------- Atualizar ----------
function atualizarDadosAluno_(sh, header, values, campos) {
  const map = headerMap_(header);
  const updates = {};
  if (map.nome        > -1) updates[map.nome]        = campos.nomeCompleto;
  if (map.local       > -1) updates[map.local]       = campos.localEvento;
  if (map.ciclo       > -1) updates[map.ciclo]       = campos.ciclo;
  if (map.status      > -1) updates[map.status]      = campos.status;

  if (map.endereco    > -1) updates[map.endereco]    = campos.endereco;
  if (map.numero      > -1) updates[map.numero]      = campos.numero;
  if (map.complemento > -1) updates[map.complemento] = campos.complemento;
  if (map.bairro      > -1) updates[map.bairro]      = campos.bairro;
  if (map.cidade      > -1) updates[map.cidade]      = campos.cidade;
  if (map.estado      > -1) updates[map.estado]      = campos.estado;
  if (map.pais        > -1) updates[map.pais]        = campos.pais;
  if (map.cep         > -1) updates[map.cep]         = campos.cep;

  if (map.profissao   > -1) updates[map.profissao]   = campos.profissao;
  if (map.escolaridade> -1) updates[map.escolaridade]= campos.escolaridade;
  if (map.graduacao   > -1 && typeof campos.graduacao !== 'undefined') updates[map.graduacao] = campos.graduacao;

  if (map.whatsapp    > -1) updates[map.whatsapp]    = campos.whatsapp;
  if (map.email       > -1) updates[map.email]       = campos.email;

  // Novos (LGPD/Consentimentos)
  if (map.lgpdVersion > -1) updates[map.lgpdVersion] = campos.lgpdVersion;
  if (map.lgpdTs      > -1) updates[map.lgpdTs]      = campos.lgpdTs;
  if (map.lgpdIp      > -1) updates[map.lgpdIp]      = campos.lgpdIp;
  if (map.optin       > -1) updates[map.optin]       = campos.optin;
  if (map.consentImg  > -1) updates[map.consentImg]  = campos.consentImagem;

  const startRow = 2;
  const lastCol = sh.getLastColumn();

  const isMatch = (row) =>
    (map.cpf   > -1 && onlyDigits_(row[map.cpf])   === campos.cpf) ||
    (map.email > -1 && String(row[map.email]).trim().toLowerCase() === campos.email);

  values.forEach((row, i) => {
    if (!isMatch(row)) return;
    let changed = false;
    Object.keys(updates).forEach(k => {
      const col = Number(k);
      const novo = updates[col];
      if (col > -1 && typeof novo !== 'undefined' && String(row[col]) !== String(novo)) {
        row[col] = novo;
        changed = true;
      }
    });
    if (changed) sh.getRange(startRow + i, 1, 1, lastCol).setValues([row]);
  });
}

/// ---------- Busca por CPF ----------
function findAlunoByCPFInSheet_(sid, sheetName, cpfDigits){
  const sh = getSheetForRead_(sid, sheetName);
  if (!sh) return { ok:false, message:`Aba "${sheetName}" não encontrada no SID ${sid}.` };

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { ok:false, message:'Nenhum dado na planilha.' };

  const header = sh.getRange(1,1,1,lastCol).getValues()[0] || [];
  const map = headerMap_(header);
  if (map.cpf === -1) return { ok:false, message:'Coluna CPF não encontrada.' };

  const data = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  for (let i = data.length - 1; i >= 0; i--){
    const row = data[i];
    const cpfRow = onlyDigits_(row[map.cpf]);
    if (cpfRow === cpfDigits){
      const get = idx => (idx > -1 ? row[idx] : '');
      const out = {
        curso:       get(map.curso),
        localEvento: get(map.local),
        ciclo:       get(map.ciclo),
        status:      get(map.status) || '',
        nome:        get(map.nome),
        email:       get(map.email),
        whatsapp:    onlyDigits_(get(map.whatsapp)),
        endereco:    get(map.endereco),
        numero:      get(map.numero),
        complemento: get(map.complemento),
        bairro:      get(map.bairro),
        cep:         onlyDigits_(get(map.cep)),
        cidade:      get(map.cidade),
        estado:      get(map.estado),
        pais:        get(map.pais),
        profissao:   get(map.profissao),
        escolaridade:get(map.escolaridade),
        graduacao:   get(map.graduacao),
      };
      if (!out.numero && out.endereco) {
        const m = String(out.endereco).match(/,\s*([\w\-\/]+)\s*$/);
        if (m) out.numero = m[1];
      }
      return { ok:true, data: out };
    }
  }
  return { ok:false, message:'Nenhum cadastro encontrado para este CPF.' };
}

function buscarAlunoPorCPF(opts) {
  try {
    const sid   = sanitize_(opts && opts.sid);
    const sheet = sanitize_(opts && opts.sheet) || SHEET_NAME;
    const cpfIn = onlyDigits_(opts && opts.cpf);

    if (!sid) throw new Error('Faltou o ID da planilha (sid).');
    if (!/^\d{11}$/.test(cpfIn)) throw new Error('Informe um CPF válido (11 dígitos).');

    let res = findAlunoByCPFInSheet_(sid, sheet, cpfIn);
    if (res.ok) return res;

    if (DEST_SPREADSHEET_ID) {
      res = findAlunoByCPFInSheet_(DEST_SPREADSHEET_ID, DEST_SHEET_NAME, cpfIn);
      if (res.ok) return res;
    }

    return res;
  } catch (err) {
    return { ok:false, message: err && err.message ? err.message : String(err) };
  }
}

/// ---------- Espelho ----------
function headerLogicalMap_(headerArr){
  const map = headerMap_(headerArr);
  return {
    timestamp: map.timestamp, curso: map.curso, local: map.local, ciclo: map.ciclo, status: map.status,
    nome: map.nome, cpf: map.cpf, email: map.email, whatsapp: map.whatsapp,
    endereco: map.endereco, numero: map.numero, complemento: map.complemento, bairro: map.bairro,
    cep: map.cep, cidade: map.cidade, estado: map.estado, pais: map.pais,
    profissao: map.profissao, escolaridade: map.escolaridade, graduacao: map.graduacao,
    // Novos
    lgpdVersion: map.lgpdVersion, lgpdTs: map.lgpdTs, lgpdIp: map.lgpdIp,
    optin: map.optin, consentImg: map.consentImg,
  };
}

function buildRowForDest_(destHeader, sourceHeader, sourceRow){
  const srcMap = headerLogicalMap_(sourceHeader);
  const dstMap = headerLogicalMap_(destHeader);
  const out = new Array(destHeader.length).fill('');
  const put = (key) => {
    const srcIdx = srcMap[key];
    const dstIdx = dstMap[key];
    if (dstIdx > -1 && srcIdx > -1) out[dstIdx] = sourceRow[srcIdx];
  };
  put('timestamp');
  if ((dstMap.timestamp === 0 || /carimbo de data\/hora/i.test(String(destHeader[0]||''))) && !out[0]) {
    out[0] = new Date();
  }
  put('curso'); put('local'); put('ciclo'); put('status');
  put('nome'); put('cpf'); put('email'); put('whatsapp');
  put('endereco'); put('numero'); put('complemento'); put('bairro');
  put('cep'); put('cidade'); put('estado'); put('pais');
  put('profissao'); put('escolaridade'); put('graduacao');

  // Novos campos
  put('lgpdVersion'); put('lgpdTs'); put('lgpdIp'); put('optin'); put('consentImg');

  return out;
}

function getDestSheet_() {
  const ss = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  let sh = ss.getSheetByName(DEST_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(DEST_SHEET_NAME);
  ensureStandardHeader_(sh);
  return sh;
}

function mirrorToSecondary_(sourceHeader, newRow){
  if (!DEST_SPREADSHEET_ID) return;
  const destSh = getDestSheet_();
  const destHeader = destSh.getRange(1,1,1,destSh.getLastColumn()).getValues()[0] || [];
  const out = buildRowForDest_(destHeader, sourceHeader, newRow);
  destSh.appendRow(out);
}
