/** Certificates.gs — isolado em namespace e compatível com Code.gs **/
(function (global) {
  'use strict';

  // ====== CONFIG usa as constantes do seu Code.gs ======
  const SHEET_NAME_G      = global.SHEET_NAME || 'BASE DE DADOS CADASTRAIS';
  const NOME_EMPRESA_G    = global.NOME_EMPRESA_DEFAULT || 'Grandha Alphaville';
  const DEST_SHEET_NAME_G = global.DEST_SHEET_NAME || SHEET_NAME_G;

  // (Opcional) ainda existe essa constante antiga, mas agora não é mais usada.
  // Pode apagar se quiser, não faz falta:
  const CERTS_FOLDER_ID = '1Ja_jnk8_0PFmse8aIQ66X1vqmaOBp1mH';

  // ====== Alvos: origem + espelho (Grandha Alphaville) ======
  const CERTIFICATE_TARGETS = [
    {
      sid: '1gCBRIGT1sFXlHQPCdWai0Mn88ATJtS7fjDyHqg32mrw', // ORIGEM – Grandha Alphaville
      sheet: SHEET_NAME_G,                                // 'BASE DE DADOS CADASTRAIS'
      rootFolderId: '1Ja_jnk8_0PFmse8aIQ66X1vqmaOBp1mH',  // PASTA RAIZ CERTIFICADOS – Alphaville
    },
    {
      sid: '1kVX0TH9_lM7e2nceMi6OktrGUZ0LootvVTLEx3AFOIo', // ESPELHO – coordenação Alphaville
      sheet: DEST_SHEET_NAME_G,                            // 'BASE DE DADOS CADASTRAIS'
      rootFolderId: '1Ja_jnk8_0PFmse8aIQ66X1vqmaOBp1mH',   // mesma pasta raiz
    },
  ].filter(t => String(t.sid || '').trim());

  
// Mapeia o texto do CURSO para a sigla da subpasta (FBTC, FATC, PEFTTC, POS)
const COURSE_CODE_RULES = [
  {
    code: 'FBTC',
    patterns: [
      'formacao basica em terapia capilar',
      'fbtc'
    ],
  },
  {
    code: 'FATC',
    patterns: [
      'formacao avancada em terapia capilar',
      'fatc'
    ],
  },
  {
    code: 'PEFTTC',
    patterns: [
      'pefttc',
      'formacao basica e avancada',
      'programa especial de formacao em terapia capilar',
      'formacao basica em terapia capilar e formacao avancada em terapia capilar'
    ],
  },
  {
    code: 'POS',
    patterns: [
      'pos graduacao em tricologia e ciencia cosmetica',
      'pos graduacao',
      'pos-graduacao',
      'pos'
    ],
  },
];
// Descobre a sigla do curso (FBTC/FATC/PEFTTC/POS) a partir da linha
function getCourseCodeForRow_(row, map) {
  if (map.curso === undefined || map.curso < 0) return '';
  var cursoNorm = norm_(row[map.curso]);
  if (!cursoNorm) return '';

  for (var i = 0; i < COURSE_CODE_RULES.length; i++) {
    var rule = COURSE_CODE_RULES[i];
    for (var j = 0; j < rule.patterns.length; j++) {
      var pattNorm = norm_(rule.patterns[j]);
      if (pattNorm && cursoNorm.indexOf(pattNorm) !== -1) {
        return rule.code;
      }
    }
  }
  return '';
}

// Pega (ou cria) uma subpasta pela sigla (FBTC, FATC, PEFTTC, POS)
function getCertificateFolderByCode_(target, code) {
  var root = getCertificateRootFolder_(target);
  if (!root || !code) return null;

  var name = code;
  var folders = root.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }

  // Se ainda não existir, cria a subpasta para essa sigla
  try {
    return root.createFolder(name);
  } catch (err) {
    Logger.log('Falha ao criar/acessar subpasta de certificados (' + name + '): ' + err);
    return null;
  }
}

// Pega a pasta raiz (configurada em rootFolderId para cada coordenação)
function getCertificateRootFolder_(target) {
  var rootId = String(target.rootFolderId || target.folderId || '').trim();
  if (!rootId) return null;
  try {
    return DriveApp.getFolderById(rootId);
  } catch (err) {
    Logger.log('Falha ao abrir pasta raiz de certificados: ' + err);
    return null;
  }
}

// Pega (ou cria) a subpasta de certificados da linha atual (FBTC/FATC/PEFTTC/POS)
function getCertificateFolderForRow_(target, row, map) {
  var root = getCertificateRootFolder_(target);
  if (!root) return null;

  var code = getCourseCodeForRow_(row, map);
  if (!code) {
    // Se não achou sigla, usa a raiz mesmo (compatibilidade)
    return root;
  }

  var codeName = code; // nome da subpasta = sigla
  var folders = root.getFoldersByName(codeName);
  if (folders.hasNext()) {
    return folders.next();
  }

  // Se não existe ainda, cria a subpasta na pasta raiz
  try {
    return root.createFolder(codeName);
  } catch (err) {
    Logger.log('Falha ao criar subpasta de certificados (' + codeName + '): ' + err);
    return root; // fallback para a raiz
  }
}

  const CERTIFICATE_FIRST_COLUMN_INDEX = 27; // AA

  const CERTIFICATE_COLUMNS = [
    { key: 'certLink',   name: 'CERTIFICADO (PDF)' },
    { key: 'certStatus', name: 'CERTIFICADO - STATUS' },
    { key: 'certSentAt', name: 'CERTIFICADO - ENVIADO EM' },
    { key: 'certError',  name: 'CERTIFICADO - ÚLTIMO ERRO' },
  ];

  const CERTIFICATE_STATUS_SENT    = 'ENVIADO';
  const CERTIFICATE_STATUS_PENDING = 'PENDENTE';
  const CERTIFICATE_STATUS_ERROR   = 'ERRO';

  const CERTIFICATE_EMAIL_TEMPLATE_FILE = 'certificate-email';
  const CERTIFICATE_EMAIL_SUBJECT_TEMPLATE = 'Certificado disponível — {{CURSO}}';

  // ====== MENUS (expostos globalmente) ======
  function onOpen() {
    try { maybeAddCertificateMenuForActiveSpreadsheet_(); }
    catch (err) { Logger.log('Falha ao construir menu (certificados): ' + err); }
  }
  function menuUpdateCertificateLinks() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const target = getCertificateTargetForSheet_(sheet);
    if (!target) return SpreadsheetApp.getUi().alert('Esta planilha/aba não está configurada para certificados.');
    const result = updateCertificateLinksForSheet_(sheet, target, { force:false });
    SpreadsheetApp.getUi().alert(
      'Atualização concluída!\n' +
      `Certificados localizados: ${result.matched}\n` +
      `Sem correspondência: ${result.notFound}\n` +
      `Linhas ignoradas: ${result.skipped}`
    );
  }
  function menuSendCertificateEmails() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const target = getCertificateTargetForSheet_(sheet);
    if (!target) return SpreadsheetApp.getUi().alert('Esta planilha/aba não está configurada para certificados.');
    const ui = SpreadsheetApp.getUi();
    if (ui.alert('Envio de certificados','Enviar os certificados pendentes?',ui.ButtonSet.YES_NO)!==ui.Button.YES) return;
    const result = sendCertificateEmailsForSheet_(sheet, target, { force:false, onlySelection:false });
    ui.alert(
      'Envio concluído!\n' +
      `Certificados enviados: ${result.sent}\n` +
      `Sem certificado: ${result.notFound}\n` +
      `Erros: ${result.errors}\n` +
      `Ignorados: ${result.skipped}`
    );
  }
  function menuResendSelectedCertificates() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const target = getCertificateTargetForSheet_(sheet);
    if (!target) return SpreadsheetApp.getUi().alert('Esta planilha/aba não está configurada para certificados.');
    const ui = SpreadsheetApp.getUi();
    if (ui.alert('Reenvio de certificados','Reenviar apenas para as linhas selecionadas?',ui.ButtonSet.YES_NO)!==ui.Button.YES) return;
    const result = sendCertificateEmailsForSheet_(sheet, target, { force:true, onlySelection:true });
    ui.alert(
      'Reenvio concluído!\n' +
      `Certificados enviados: ${result.sent}\n` +
      `Sem certificado: ${result.notFound}\n` +
      `Erros: ${result.errors}\n` +
      `Ignorados: ${result.skipped}`
    );
  }

  // ====== UI interno ======
  function maybeAddCertificateMenuForActiveSpreadsheet_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    const sid = ss.getId();
    const enabled = CERTIFICATE_TARGETS.some(t => t.sid === sid);
    if (!enabled) return;
    SpreadsheetApp.getUi()
      .createMenu('Certificados')
      .addItem('Atualizar links (planilha atual)', 'menuUpdateCertificateLinks')
      .addItem('Enviar certificados pendentes', 'menuSendCertificateEmails')
      .addSeparator()
      .addItem('Reenviar certificados selecionados', 'menuResendSelectedCertificates')
      .addToUi();
  }

  // ====== Núcleo ======
  function getCertificateTargetForSheet_(sheet) {
    if (!sheet) return null;
    try {
      const sid = sheet.getParent().getId();
      const sheetName = sheet.getName();
      return CERTIFICATE_TARGETS.find(t => t.sid === sid && (t.sheet || SHEET_NAME_G) === sheetName) || null;
    } catch (err) { Logger.log(err); return null; }
  }

  // >>>> FUNÇÃO EXPORTADA (o seu Code.gs chama isso no espelho)
  function ensureCertificateColumns_(sheet) {
    if (!sheet) return;
    const target = getCertificateTargetForSheet_(sheet);
    if (!target) return; // só cria nas planilhas configuradas

    const baseIndex = CERTIFICATE_FIRST_COLUMN_INDEX;
    const maxNeeded = baseIndex + CERTIFICATE_COLUMNS.length - 1;

    if (sheet.getMaxColumns() < maxNeeded) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), maxNeeded - sheet.getMaxColumns());
    }

    const maxAttempts = CERTIFICATE_COLUMNS.length * 4;
    for (let attempt = 0; attempt < maxAttempts; attempt++) {
      const width = Math.max(sheet.getLastColumn(), maxNeeded);
      const header = sheet.getRange(1, 1, 1, width).getValues()[0] || [];
      const normHeader = header.map(h => norm_(h));

      let updated = false;
      for (let i = 0; i < CERTIFICATE_COLUMNS.length; i++) {
        const col = CERTIFICATE_COLUMNS[i];
        const desiredIndex = baseIndex + i;
        const wantedNorm = norm_(col.name);
        const existingIdx0 = normHeader.indexOf(wantedNorm);
        const existingIndex = existingIdx0 === -1 ? -1 : existingIdx0 + 1;

        if (existingIndex === desiredIndex) {
          const cell = sheet.getRange(1, desiredIndex);
          if (cell.getValue() !== col.name) cell.setValue(col.name);
          continue;
        }
        if (existingIndex === -1) {
          ensureColumnSlot_(sheet, desiredIndex);
          sheet.getRange(1, desiredIndex).setValue(col.name);
        } else {
          moveColumnToIndex_(sheet, existingIndex, desiredIndex);
          sheet.getRange(1, desiredIndex).setValue(col.name);
        }
        updated = true;
        break; // realinha uma por vez
      }
      if (!updated) { sheet.setFrozenRows(1); return; }
    }
    sheet.setFrozenRows(1);
  }

  function updateCertificateLinksForSheet_(sheet, target, options) {
  ensureCertificateColumns_(sheet);

  var rootFolder = getCertificateRootFolder_(target);
  if (!rootFolder) throw new Error('Não foi possível acessar a pasta raiz de certificados.');

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { matched: 0, notFound: 0, skipped: 0 };

  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  var map = headerMap_(header);
  var linkIdx   = map.certLink;
  var statusIdx = map.certStatus;
  var sentAtIdx = map.certSentAt;
  var errorIdx  = map.certError;
  if (linkIdx === -1) throw new Error('Coluna "CERTIFICADO (PDF)" não localizada.');

  var range  = sheet.getRange(2, 1, lastRow - 1, lastCol);
  var values = range.getValues();

  var matched = 0, notFound = 0, skipped = 0;
  var force = !!(options && options.force);
  var rowsToSync = [];

  // cache de índices por subpasta: folderId -> fileIndex
  var indexCache = {};

  values.forEach(function(row, idx) {
    var hasCpf  = map.cpf  > -1 ? !!onlyDigits_(row[map.cpf])  : false;
    var hasNome = map.nome > -1 ? !!sanitize_(row[map.nome])   : false;
    if (!hasCpf && !hasNome) { skipped++; return; }

    var currentLink   = sanitize_(row[linkIdx]);
    var currentStatus = statusIdx > -1 ? sanitize_(row[statusIdx]) : '';

    if (!force && currentLink) { skipped++; return; }

    var originalLink   = linkIdx   > -1 ? row[linkIdx]   : '';
    var originalStatus = statusIdx > -1 ? row[statusIdx] : '';
    var originalSentAt = sentAtIdx > -1 ? row[sentAtIdx] : '';
    var originalError  = errorIdx  > -1 ? row[errorIdx]  : '';

    // *** Escolhe a subpasta (FBTC/FATC/PEFTTC/POS) para esta linha ***
    var folder = getCertificateFolderForRow_(target, row, map);
    if (!folder) { skipped++; return; }

    var folderId = folder.getId();
    if (!indexCache[folderId]) {
      indexCache[folderId] = buildCertificateFileIndex_(folder);
    }
    var fileIndex = indexCache[folderId];

    var changed = false;
    var file = findCertificateFileForRow_(fileIndex, row, map);

    if (file) {
      matched++;
      if (originalLink !== file.url) {
        row[linkIdx] = file.url;
        changed = true;
      }
      if (statusIdx > -1 && currentStatus !== 'ENVIADO' && originalStatus !== 'PENDENTE') {
        row[statusIdx] = 'PENDENTE';
        changed = true;
      }
      if (sentAtIdx > -1 && currentStatus !== 'ENVIADO' && originalSentAt) {
        row[sentAtIdx] = '';
        changed = true;
      }
      if (errorIdx > -1 && originalError) {
        row[errorIdx] = '';
        changed = true;
      }
    } else {
      notFound++;
      if (errorIdx > -1 && (!currentLink || force)) {
        var msg = 'Certificado não encontrado na pasta configurada.';
        if (originalError !== msg) {
          row[errorIdx] = msg;
          changed = true;
        }
      }
      if (statusIdx > -1 && currentStatus !== 'ENVIADO' && originalStatus !== 'ERRO') {
        row[statusIdx] = 'ERRO';
        changed = true;
      }
      if (sentAtIdx > -1 && originalSentAt) {
        row[sentAtIdx] = '';
        changed = true;
      }
    }

    if (changed) {
      rowsToSync.push(idx + 2); // linha na planilha (começa em 2)
    }
  });

  range.setValues(values);
  SpreadsheetApp.flush();
  syncCertificateChangesForRows_(sheet, header, values, rowsToSync);

  return { matched: matched, notFound: notFound, skipped: skipped };
}


  function sendCertificateEmailsForSheet_(sheet, target, options) {
  ensureCertificateColumns_(sheet);

  var rootFolder = getCertificateRootFolder_(target);
  if (!rootFolder) throw new Error('Não foi possível acessar a pasta raiz de certificados.');

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { sent: 0, notFound: 0, errors: 0, skipped: 0 };
  }

  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  var map = headerMap_(header);
  var linkIdx   = map.certLink;
  var statusIdx = map.certStatus;
  var sentAtIdx = map.certSentAt;
  var errorIdx  = map.certError;

  if (linkIdx === -1) {
    throw new Error('Coluna "CERTIFICADO (PDF)" não localizada.');
  }
  if (statusIdx === -1) {
    throw new Error('Coluna "CERTIFICADO - STATUS" não localizada.');
  }

  var selectionRows = getSelectedRowIndexes_(sheet);
  var useSelection  = !!(options && options.onlySelection);
  var force         = !!(options && options.force);

  var range  = sheet.getRange(2, 1, lastRow - 1, lastCol);
  var values = range.getValues();

  var sent = 0, notFound = 0, errors = 0, skipped = 0;
  var rowsToSync = [];

  // cache de índices por subpasta: folderId -> fileIndex
  var indexCache = {};

  values.forEach(function(row, idx) {
    var rowNumber = idx + 2;

    if (useSelection && selectionRows.indexOf(rowNumber) === -1) {
      skipped++;
      return;
    }

    var emailIdx = map.email;
    var email = emailIdx > -1 ? sanitize_(row[emailIdx]) : '';
    if (!isValidEmail_(email)) {
      skipped++;
      return;
    }

    var hasCpf  = map.cpf  > -1 ? !!onlyDigits_(row[map.cpf])  : false;
    var hasNome = map.nome > -1 ? !!sanitize_(row[map.nome])   : false;
    if (!hasCpf && !hasNome) {
      skipped++;
      return;
    }

    var currentStatus = statusIdx > -1 ? sanitize_(row[statusIdx]) : '';
    var shouldSend = force ||
      !currentStatus ||
      currentStatus === CERTIFICATE_STATUS_PENDING ||
      currentStatus === CERTIFICATE_STATUS_ERROR;

    if (!shouldSend) {
      skipped++;
      return;
    }

    var originalLink   = linkIdx   > -1 ? row[linkIdx]   : '';
    var originalStatus = statusIdx > -1 ? row[statusIdx] : '';
    var originalSentAt = sentAtIdx > -1 ? row[sentAtIdx] : '';
    var originalError  = errorIdx  > -1 ? row[errorIdx]  : '';
    var changed = false;

    // ===== 1) Certificado principal (curso da própria linha) =====
    var folderMain = getCertificateFolderForRow_(target, row, map);
    if (!folderMain) {
      notFound++;
      if (statusIdx > -1 && originalStatus !== CERTIFICATE_STATUS_ERROR) {
        row[statusIdx] = CERTIFICATE_STATUS_ERROR;
        changed = true;
      }
      if (sentAtIdx > -1 && originalSentAt) {
        row[sentAtIdx] = '';
        changed = true;
      }
      if (errorIdx > -1) {
        var msgMain = 'Pasta de certificados não encontrada para este curso.';
        if (originalError !== msgMain) {
          row[errorIdx] = msgMain;
          changed = true;
        }
      }
      if (changed) rowsToSync.push(rowNumber);
      return;
    }

    var folderIdMain = folderMain.getId();
    if (!indexCache[folderIdMain]) {
      indexCache[folderIdMain] = buildCertificateFileIndex_(folderMain);
    }
    var fileIndexMain = indexCache[folderIdMain];

    var mainFile = null;
    var link = sanitize_(row[linkIdx]);

    // tenta primeiro pelo link (Drive direto)
    if (link) {
      var fileId = extractDriveFileIdFromUrl_(link);
      if (fileId) {
        try {
          var fetched = DriveApp.getFileById(fileId);
          mainFile = buildCertificateEntryFromFile_(fetched);
        } catch (e) {
          // se der erro, cai para busca por CPF/NOME
        }
      }
    }

    // se não achou pelo link, tenta pelo índice da subpasta
    if (!mainFile) {
      mainFile = findCertificateFileForRow_(fileIndexMain, row, map);
    }

    // array de anexos
    var attachments = [];
    var payloadFileForTemplate = null;

    if (mainFile) {
      attachments.push(mainFile.file.getAs(MimeType.PDF));
      payloadFileForTemplate = mainFile;
    }

    // ===== 2) Regra especial: se curso = FATC, tenta anexar também PEFTTC =====
    var courseCode = getCourseCodeForRow_(row, map);
    if (courseCode === 'FATC') {
      var folderCombo = getCertificateFolderByCode_(target, 'PEFTTC');
      if (folderCombo) {
        var comboFolderId = folderCombo.getId();
        if (!indexCache[comboFolderId]) {
          indexCache[comboFolderId] = buildCertificateFileIndex_(folderCombo);
        }
        var fileIndexCombo = indexCache[comboFolderId];
        var comboFile = findCertificateFileForRow_(fileIndexCombo, row, map);
        if (comboFile && (!mainFile || comboFile.id !== mainFile.id)) {
          attachments.push(comboFile.file.getAs(MimeType.PDF));
          // se não tiver arquivo principal, usa o combo como base do template
          if (!payloadFileForTemplate) {
            payloadFileForTemplate = comboFile;
          }
        }
      }
    }

    // Se ainda não tem nenhum PDF, trata como "não encontrado"
    if (!attachments.length || !payloadFileForTemplate) {
      notFound++;
      if (statusIdx > -1 && originalStatus !== CERTIFICATE_STATUS_ERROR) {
        row[statusIdx] = CERTIFICATE_STATUS_ERROR;
        changed = true;
      }
      if (sentAtIdx > -1 && originalSentAt) {
        row[sentAtIdx] = '';
        changed = true;
      }
      if (errorIdx > -1) {
        var msgNotFound = 'Certificado não encontrado na pasta configurada.';
        if (originalError !== msgNotFound) {
          row[errorIdx] = msgNotFound;
          changed = true;
        }
      }
      if (changed) rowsToSync.push(rowNumber);
      return;
    }

    // ===== 3) Envia o e-mail com 1 ou 2 anexos =====
    try {
      var payload = buildCertificateEmailPayload_(row, map, payloadFileForTemplate);

      GmailApp.sendEmail(email, payload.subject, ' ', {
        name: payload.senderName,
        htmlBody: payload.htmlBody,
        attachments: attachments
      });

      sent++;
      if (statusIdx > -1 && originalStatus !== CERTIFICATE_STATUS_SENT) {
        row[statusIdx] = CERTIFICATE_STATUS_SENT;
        changed = true;
      }
      if (sentAtIdx > -1) {
        row[sentAtIdx] = new Date();
        changed = true;
      }
      if (errorIdx > -1 && originalError) {
        row[errorIdx] = '';
        changed = true;
      }
      // garante que tenha link, se ainda estiver vazio e tivermos ao menos o principal
      if (linkIdx > -1 && !sanitize_(originalLink) && mainFile) {
        row[linkIdx] = mainFile.url;
        changed = true;
      }

    } catch (err) {
      errors++;
      if (statusIdx > -1 && originalStatus !== CERTIFICATE_STATUS_ERROR) {
        row[statusIdx] = CERTIFICATE_STATUS_ERROR;
        changed = true;
      }
      if (sentAtIdx > -1) {
        row[sentAtIdx] = '';
        changed = true;
      }
      if (errorIdx > -1) {
        var msgSend = 'Falha ao enviar e-mail: ' + err;
        if (originalError !== msgSend) {
          row[errorIdx] = msgSend;
          changed = true;
        }
      }
    }

    if (changed) {
      rowsToSync.push(rowNumber);
    }
  });

  range.setValues(values);
  SpreadsheetApp.flush();
  syncCertificateChangesForRows_(sheet, header, values, rowsToSync);

  return { sent: sent, notFound: notFound, errors: errors, skipped: skipped };
}


  // ====== Sync (usa hooks já existentes no seu Code.gs) ======
  function syncCertificateChangesForRows_(sheet, header, values, rowNumbers) {
    if (!Array.isArray(rowNumbers) || !rowNumbers.length) return;
    let sourceTarget = null;
    try {
      // usa a função do seu Code.gs, se existir
      sourceTarget = (typeof global.getSyncTargetForSheet_ === 'function') ? global.getSyncTargetForSheet_(sheet) : null;
    } catch (err) { Logger.log('Falha ao identificar destino de sincronização: ' + err); }
    if (!sourceTarget || typeof global.syncTargetsForRow_ !== 'function') return;

    rowNumbers.forEach(rowNumber => {
      const idx = rowNumber - 2;
      if (idx < 0 || idx >= values.length) return;
      try { global.syncTargetsForRow_(sourceTarget, header, values[idx]); }
      catch (err) { Logger.log('Falha ao sincronizar certificados (linha '+rowNumber+'): ' + err); }
    });
  }

  // ====== Drive/Index ======
  function getCertificateFolder_(target) {
    if (!target) return null;
    const folderId = String(target.folderId || '').trim();
    if (!folderId) return null;
    try { return DriveApp.getFolderById(folderId); }
    catch (err) { Logger.log('Falha ao abrir pasta de certificados: ' + err); return null; }
  }
  function buildCertificateFileIndex_(folder) {
    const entries = [], byId = {};
    const it = folder.getFiles();
    while (it.hasNext()) {
      const file = it.next();
      const entry = buildCertificateEntryFromFile_(file);
      if (!entry) continue;
      entries.push(entry); byId[entry.id] = entry;
    }
    return { entries, byId };
  }
  function buildCertificateEntryFromFile_(file) {
    try {
      const mime = file.getMimeType(), name = file.getName();
      const isPdf = mime === MimeType.PDF || /\.pdf$/i.test(String(name||''));
      if (!isPdf) return null;
      return {
        id: file.getId(), url: file.getUrl(), name,
        normalized: norm_(name), digits: onlyDigits_(name), file
      };
    } catch (err) { Logger.log('Indexação falhou: '+err); return null; }
  }
  function findCertificateFileForRow_(fileIndex, row, map) {
    if (!fileIndex || !fileIndex.entries || !fileIndex.entries.length) return null;
    const cpf = map.cpf > -1 ? onlyDigits_(row[map.cpf]) : '';
    if (cpf) {
      const hit = fileIndex.entries.find(e => e.digits && e.digits.includes(cpf));
      if (hit) return hit;
    }
    const name = map.nome > -1 ? norm_(row[map.nome]) : '';
    if (name) {
      const exact = fileIndex.entries.find(e => e.normalized === name);
      if (exact) return exact;
      const hit = fileIndex.entries.find(e => e.normalized.includes(name));
      if (hit) return hit;
      const tokens = name.split(' ').filter(Boolean);
      if (tokens.length >= 2) {
        const first = tokens[0], last = tokens[tokens.length - 1];
        const partial = fileIndex.entries.find(e => e.normalized.includes(first) && e.normalized.includes(last));
        if (partial) return partial;
      }
    }
    return null;
  }
  function extractDriveFileIdFromUrl_(url) {
    const raw = sanitize_(url);
    if (!raw) return '';
    const patterns = [/\/d\/([a-zA-Z0-9_-]+)/, /id=([a-zA-Z0-9_-]+)/, /\/(?:file|folders)\/([a-zA-Z0-9_-]+)/];
    for (let i=0;i<patterns.length;i++){ const m=raw.match(patterns[i]); if (m && m[1]) return m[1]; }
    if (/^[a-zA-Z0-9_-]{10,}$/.test(raw)) return raw;
    return '';
  }

  // ====== E-mail ======
  function buildCertificateEmailPayload_(row, map, fileEntry) {
    const nome    = map.nome    > -1 ? sanitize_(row[map.nome])    : '';
    const curso   = map.curso   > -1 ? sanitize_(row[map.curso])   : '';
    const localEv = map.local   > -1 ? sanitize_(row[map.local])   : '';
    const ciclo   = map.ciclo   > -1 ? sanitize_(row[map.ciclo])   : '';
    const empresa = map.empresa > -1 ? sanitize_(row[map.empresa]) : NOME_EMPRESA_G;
    const status  = map.status  > -1 ? sanitize_(row[map.status])  : '';

    let template = '';
    try { template = HtmlService.createHtmlOutputFromFile(CERTIFICATE_EMAIL_TEMPLATE_FILE).getContent(); }
    catch (_e) {
      template = '<p>Olá, {{NOME}}!</p>' +
                 '<p>Seu certificado do curso <strong>{{CURSO}}</strong> {{CICLO}} {{LOCAL_EVENTO}} está disponível.</p>' +
                 '<p>Acesse o PDF: <a href="{{CERTIFICATE_LINK}}">Abrir Certificado</a></p>' +
                 '<p>Status: {{STATUS}}</p><p>Atenciosamente,<br>{{NOME_EMPRESA}}</p>';
    }

    const htmlBody = replacePlaceholders_(template, {
      NOME: escapeHtml_(nome),
      CURSO: escapeHtml_(curso || 'Programa de Certificações'),
      LOCAL_EVENTO: escapeHtml_(localEv || 'Online'),
      CICLO: escapeHtml_(ciclo || '—'),
      STATUS: escapeHtml_(status || ''),
      NOME_EMPRESA: escapeHtml_(empresa || NOME_EMPRESA_G),
      CERTIFICATE_LINK: escapeHtml_(fileEntry.url),
    });

    const subject = replacePlaceholders_(CERTIFICATE_EMAIL_SUBJECT_TEMPLATE, {
      CURSO: curso || 'Programa de Certificações',
      NOME: nome || '',
      NOME_EMPRESA: empresa || NOME_EMPRESA_G,
    });

    return { subject, htmlBody, senderName: empresa || NOME_EMPRESA_G };
  }

  // ====== Header map (leve e isolado para não colidir) ======
  function headerMap_(header){
    const clean = s => (s || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
    const H = header.map(clean);
    const find = (...labels) => {
      const L = labels.map(clean);
      return H.findIndex(h => L.includes(h));
    };
    return {
      cpf: find('cpf'),
      nome: find('nome completo sem abreviacoes','nome completo','nome'),
      email: find('e mail','email','e-mail'),
      curso: find('curso'),
      local: find('local do evento','local evento','local'),
      ciclo: find('ciclo'),
      status: find('status'),
      empresa: find('empresa'),
      certLink: find('certificado pdf','certificado (pdf)'),
      certStatus: find('certificado - status','certificado status'),
      certSentAt: find('certificado - enviado em','certificado enviado em'),
      certError: find('certificado - ultimo erro','certificado ultimo erro'),
    };
  }

  // ====== Helpers PRIVADOS (não poluem global) ======
  function sanitize_(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    return String(v).trim();
  }
  function norm_(v) {
    return String(sanitize_(v))
      .toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
      .replace(/\s+/g,' ')
      .replace(/[^\w\s@.-]/g,'')
      .trim();
  }
  function onlyDigits_(v) { return sanitize_(v).replace(/\D+/g, ''); }
  function isValidEmail_(email) { const e = sanitize_(email); return !!e && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); }
  function escapeHtml_(s) {
    const map = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
    return String(s).replace(/[&<>"']/g, m => map[m]);
  }
  function replacePlaceholders_(template, repl) {
    return Object.keys(repl).reduce((acc, k) => acc.replace(new RegExp(`{{${k}}}`, 'g'), repl[k]), template);
  }

  function ensureColumnSlot_(sheet, columnIndex) {
    if (columnIndex > sheet.getMaxColumns()) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), columnIndex - sheet.getMaxColumns());
    }
    if (columnIndex > sheet.getMaxColumns()) return;
    const currentHeader = sheet.getRange(1, columnIndex).getValue();
    if (norm_(currentHeader)) sheet.insertColumnBefore(columnIndex);
  }
  function moveColumnToIndex_(sheet, fromIndex, toIndex) {
    if (fromIndex === toIndex) return;
    const maxRows = Math.max(sheet.getMaxRows(), 1);
    if (fromIndex < toIndex) {
      sheet.insertColumnAfter(toIndex);
      const destIndex = toIndex + 1;
      sheet.getRange(1, fromIndex, maxRows, 1)
           .copyTo(sheet.getRange(1, destIndex, maxRows, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      sheet.deleteColumn(fromIndex);
    } else {
      sheet.insertColumnBefore(toIndex);
      const destIndex = toIndex;
      sheet.getRange(1, fromIndex + 1, maxRows, 1)
           .copyTo(sheet.getRange(1, destIndex, maxRows, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      sheet.deleteColumn(fromIndex + 1);
    }
  }
  function getSelectedRowIndexes_(sheet) {
    try {
      const range = sheet.getActiveRange();
      if (!range) return [];
      const start = range.getRow(), count = range.getNumRows();
      const rows = [];
      for (let i = 0; i < count; i++) rows.push(start + i);
      return rows;
    } catch (_e) { return []; }
  }

    // ====== Exporta só o necessário (compatível com seu Code.gs) ======
  global.menuUpdateCertificateLinks = menuUpdateCertificateLinks;
  global.menuSendCertificateEmails = menuSendCertificateEmails;
  global.menuResendSelectedCertificates = menuResendSelectedCertificates;

  // IMPORTANTÍSSIMO: o seu Code.gs chama ensureCertificateColumns_ no espelho
  global.ensureCertificateColumns_ = ensureCertificateColumns_;

  // Exporta também o construtor de menu para ser usado por um onOpen global
  global.maybeAddCertificateMenuForActiveSpreadsheet_ = maybeAddCertificateMenuForActiveSpreadsheet_;

})(this);

// ====== Simple trigger GLOBAL (fora do IIFE) ======
/**
 * Este é o gatilho simples que o Google Sheets enxerga como "onOpen".
 * Ele só delega para a função que está dentro do módulo de certificados.
 */
function onOpen(e) {
  if (typeof maybeAddCertificateMenuForActiveSpreadsheet_ === 'function') {
    maybeAddCertificateMenuForActiveSpreadsheet_();
  }
}

