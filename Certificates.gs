// ---------- Certificates.gs — automações de certificados ----------

/**
 * Configurações específicas de certificados.
 *
 * >>> ATENÇÃO: ajuste o folderId de cada planilha conforme a coordenação. <<<
 * Para replicar para outras coordenações, basta duplicar os objetos abaixo
 * informando o ID da planilha (sid), o nome da aba (sheet) e o ID da pasta
 * do Google Drive onde os certificados em PDF ficarão armazenados.
 */
const CERTIFICATE_TARGETS = [
  {
    sid: SID_DEFAULT,
    sheet: SHEET_NAME,
    folderId: '1Ja_jnk8_0PFmse8aIQ66X1vqmaOBp1mH', // Pasta "CERTIFICADOS"
  },
  {
    sid: DEST_SPREADSHEET_ID,
    sheet: DEST_SHEET_NAME,
    folderId: '1Ja_jnk8_0PFmse8aIQ66X1vqmaOBp1mH',
  },
  // { sid: 'ID_DA_PLANILHA', sheet: 'NOME_DA_ABA', folderId: 'ID_DA_PASTA' },
].filter(t => String(t.sid || '').trim());

const CERTIFICATE_FIRST_COLUMN_INDEX = 27; // Coluna AA

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

function onOpen(e) {
  try {
    maybeAddCertificateMenuForActiveSpreadsheet_();
  } catch (err) {
    Logger.log('Falha ao construir menu de certificados: ' + err);
  }
}

function maybeAddCertificateMenuForActiveSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  const sid = ss.getId();
  const enabled = CERTIFICATE_TARGETS.some(t => t.sid === sid);
  if (!enabled) return;

  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Certificados')
    .addItem('Atualizar links (planilha atual)', 'menuUpdateCertificateLinks')
    .addItem('Enviar certificados pendentes', 'menuSendCertificateEmails')
    .addSeparator()
    .addItem('Reenviar certificados selecionados', 'menuResendSelectedCertificates')
    .addToUi();
}

function getCertificateTargetForSheet_(sheet) {
  if (!sheet) return null;
  try {
    const sid = sheet.getParent().getId();
    const sheetName = sheet.getName();
    return CERTIFICATE_TARGETS.find(t => t.sid === sid && (t.sheet || SHEET_NAME) === sheetName) || null;
  } catch (err) {
    return null;
  }
}

function ensureCertificateColumns_(sheet) {
  if (!sheet) return;
  const target = getCertificateTargetForSheet_(sheet);
  if (!target) return; // não cria colunas nas planilhas que não estão configuradas

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
      const existingIndex = existingIdx0 === -1 ? -1 : existingIdx0 + 1; // 1-based

      if (existingIndex === desiredIndex) {
        const cell = sheet.getRange(1, desiredIndex);
        if (cell.getValue() !== col.name) {
          cell.setValue(col.name);
        }
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
      break;
    }

    if (!updated) {
      sheet.setFrozenRows(1);
      return;
    }
  }

  sheet.setFrozenRows(1);
}

function ensureColumnSlot_(sheet, columnIndex) {
  if (columnIndex > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), columnIndex - sheet.getMaxColumns());
  }

  if (columnIndex > sheet.getMaxColumns()) {
    return;
  }

  const currentHeader = sheet.getRange(1, columnIndex).getValue();
  if (norm_(currentHeader)) {
    sheet.insertColumnBefore(columnIndex);
  }
}

function moveColumnToIndex_(sheet, fromIndex, toIndex) {
  if (fromIndex === toIndex) return;

  const maxRows = Math.max(sheet.getMaxRows(), 1);

  if (fromIndex < toIndex) {
    sheet.insertColumnAfter(toIndex);
    const destIndex = toIndex + 1;
    const srcRange = sheet.getRange(1, fromIndex, maxRows, 1);
    const destRange = sheet.getRange(1, destIndex, maxRows, 1);
    srcRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    sheet.deleteColumn(fromIndex);
  } else {
    sheet.insertColumnBefore(toIndex);
    const destIndex = toIndex;
    const srcRange = sheet.getRange(1, fromIndex + 1, maxRows, 1);
    const destRange = sheet.getRange(1, destIndex, maxRows, 1);
    srcRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    sheet.deleteColumn(fromIndex + 1);
  }
}

function menuUpdateCertificateLinks() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const target = getCertificateTargetForSheet_(sheet);
  if (!target) {
    SpreadsheetApp.getUi().alert('Esta planilha não está configurada para automação de certificados.');
    return;
  }

  const result = updateCertificateLinksForSheet_(sheet, target, { force: false });
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
  if (!target) {
    SpreadsheetApp.getUi().alert('Esta planilha não está configurada para automação de certificados.');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Envio de certificados',
    'Deseja enviar os certificados pendentes para os alunos desta planilha?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const result = sendCertificateEmailsForSheet_(sheet, target, { force: false, onlySelection: false });
  ui.alert(
    'Envio concluído!\n' +
    `Certificados enviados: ${result.sent}\n` +
    `Sem certificado localizado: ${result.notFound}\n` +
    `Com erro de envio: ${result.errors}\n` +
    `Ignorados: ${result.skipped}`
  );
}

function menuResendSelectedCertificates() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const target = getCertificateTargetForSheet_(sheet);
  if (!target) {
    SpreadsheetApp.getUi().alert('Esta planilha não está configurada para automação de certificados.');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Reenvio de certificados',
    'Deseja reenviar os certificados apenas para as linhas selecionadas?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const result = sendCertificateEmailsForSheet_(sheet, target, { force: true, onlySelection: true });
  ui.alert(
    'Reenvio concluído!\n' +
    `Certificados enviados: ${result.sent}\n` +
    `Sem certificado localizado: ${result.notFound}\n` +
    `Com erro de envio: ${result.errors}\n` +
    `Ignorados: ${result.skipped}`
  );
}

function updateCertificateLinksForSheet_(sheet, target, options) {
  ensureCertificateColumns_(sheet);

  const folder = getCertificateFolder_(target);
  if (!folder) {
    throw new Error('Não foi possível acessar a pasta de certificados. Verifique o ID configurado e autorize o acesso ao Drive quando solicitado.');
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { matched: 0, notFound: 0, skipped: 0 };
  }

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const map = headerMap_(header);
  const linkIdx = map.certLink;
  if (linkIdx === -1) throw new Error('Coluna "CERTIFICADO (PDF)" não localizada.');

  const statusIdx = map.certStatus;
  const sentAtIdx = map.certSentAt;
  const errorIdx = map.certError;

  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const values = range.getValues();
  const fileIndex = buildCertificateFileIndex_(folder);

  let matched = 0;
  let notFound = 0;
  let skipped = 0;
  const force = options && options.force;

  values.forEach((row, idx) => {
    const currentLink = sanitize_(row[linkIdx]);
    const currentStatus = statusIdx > -1 ? sanitize_(row[statusIdx]) : '';
    const hasCpf = map.cpf > -1 ? !!onlyDigits_(row[map.cpf]) : false;
    const hasNome = map.nome > -1 ? !!sanitize_(row[map.nome]) : false;
    if (!hasCpf && !hasNome) {
      skipped++;
      return;
    }

    if (!force && currentLink) {
      skipped++;
      return;
    }

    const file = findCertificateFileForRow_(fileIndex, row, map);
    if (file) {
      matched++;
      row[linkIdx] = file.url;

      if (statusIdx > -1 && currentStatus !== CERTIFICATE_STATUS_SENT) {
        row[statusIdx] = CERTIFICATE_STATUS_PENDING;
      }
      if (sentAtIdx > -1 && currentStatus !== CERTIFICATE_STATUS_SENT) {
        row[sentAtIdx] = '';
      }
      if (errorIdx > -1) {
        row[errorIdx] = '';
      }
    } else {
      notFound++;
      if (errorIdx > -1 && (!currentLink || force)) {
        row[errorIdx] = 'Certificado não encontrado na pasta configurada.';
      }
      if (statusIdx > -1 && currentStatus !== CERTIFICATE_STATUS_SENT) {
        row[statusIdx] = CERTIFICATE_STATUS_ERROR;
      }
    }
  });

  range.setValues(values);
  return { matched, notFound, skipped };
}

function sendCertificateEmailsForSheet_(sheet, target, options) {
  ensureCertificateColumns_(sheet);

  const folder = getCertificateFolder_(target);
  if (!folder) {
    throw new Error('Não foi possível acessar a pasta de certificados. Verifique o ID configurado e autorize o acesso ao Drive quando solicitado.');
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { sent: 0, notFound: 0, errors: 0, skipped: 0 };
  }

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const map = headerMap_(header);
  const linkIdx = map.certLink;
  const statusIdx = map.certStatus;
  const sentAtIdx = map.certSentAt;
  const errorIdx = map.certError;

  if (linkIdx === -1) throw new Error('Coluna "CERTIFICADO (PDF)" não localizada.');
  if (statusIdx === -1) throw new Error('Coluna "CERTIFICADO - STATUS" não localizada.');

  const selectionRows = getSelectedRowIndexes_(sheet);
  const useSelection = !!(options && options.onlySelection);
  const force = !!(options && options.force);

  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const values = range.getValues();
  const fileIndex = buildCertificateFileIndex_(folder);

  let sent = 0;
  let notFound = 0;
  let errors = 0;
  let skipped = 0;

  values.forEach((row, idx) => {
    const rowNumber = idx + 2;
    if (useSelection && !selectionRows.includes(rowNumber)) {
      skipped++;
      return;
    }

    const emailIdx = map.email;
    const email = emailIdx > -1 ? sanitize_(row[emailIdx]) : '';
    if (!isValidEmail_(email)) {
      skipped++;
      return;
    }

    const hasCpf = map.cpf > -1 ? !!onlyDigits_(row[map.cpf]) : false;
    const hasNome = map.nome > -1 ? !!sanitize_(row[map.nome]) : false;
    if (!hasCpf && !hasNome) {
      skipped++;
      return;
    }

    const currentStatus = statusIdx > -1 ? sanitize_(row[statusIdx]) : '';
    const shouldSend = force || !currentStatus || currentStatus === CERTIFICATE_STATUS_PENDING || currentStatus === CERTIFICATE_STATUS_ERROR;
    if (!shouldSend) {
      skipped++;
      return;
    }

    let file = null;
    const link = sanitize_(row[linkIdx]);
    if (link) {
      const fileId = extractDriveFileIdFromUrl_(link);
      if (fileId) {
        const hit = fileIndex.byId[fileId];
        if (hit) {
          file = hit;
        } else {
          try {
            const fetched = DriveApp.getFileById(fileId);
            const entry = buildCertificateEntryFromFile_(fetched);
            if (entry) {
              file = entry;
              fileIndex.entries.push(entry);
              fileIndex.byId[entry.id] = entry;
            }
          } catch (err) {
            // segue para tentar localizar por CPF/nome
          }
        }
      }
    }

    if (!file) {
      file = findCertificateFileForRow_(fileIndex, row, map);
    }

    if (!file) {
      notFound++;
      row[statusIdx] = CERTIFICATE_STATUS_ERROR;
      if (sentAtIdx > -1) row[sentAtIdx] = '';
      if (errorIdx > -1) row[errorIdx] = 'Certificado não encontrado na pasta configurada.';
      return;
    }

    try {
      const emailPayload = buildCertificateEmailPayload_(row, map, file);
      GmailApp.sendEmail(email, emailPayload.subject, ' ', {
        name: emailPayload.senderName,
        htmlBody: emailPayload.htmlBody,
        attachments: [file.file.getAs(MimeType.PDF)],
      });

      sent++;
      row[statusIdx] = CERTIFICATE_STATUS_SENT;
      if (sentAtIdx > -1) row[sentAtIdx] = new Date();
      if (errorIdx > -1) row[errorIdx] = '';
      if (linkIdx > -1 && !sanitize_(row[linkIdx])) {
        row[linkIdx] = file.url;
      }
    } catch (err) {
      errors++;
      row[statusIdx] = CERTIFICATE_STATUS_ERROR;
      if (sentAtIdx > -1) row[sentAtIdx] = '';
      if (errorIdx > -1) row[errorIdx] = 'Falha ao enviar e-mail: ' + err;
    }
  });

  range.setValues(values);
  return { sent, notFound, errors, skipped };
}

function getCertificateFolder_(target) {
  if (!target) return null;
  const folderId = String(target.folderId || '').trim();
  if (!folderId) return null;
  try {
    return DriveApp.getFolderById(folderId);
  } catch (err) {
    Logger.log('Falha ao abrir pasta de certificados: ' + err);
    return null;
  }
}

function buildCertificateFileIndex_(folder) {
  const entries = [];
  const byId = {};
  const iterator = folder.getFiles();
  while (iterator.hasNext()) {
    const file = iterator.next();
    const entry = buildCertificateEntryFromFile_(file);
    if (!entry) continue;
    entries.push(entry);
    byId[entry.id] = entry;
  }
  return { entries, byId };
}

function buildCertificateEntryFromFile_(file) {
  if (!file) return null;
  try {
    const mime = file.getMimeType();
    const name = file.getName();
    const isPdf = mime === MimeType.PDF || /\.pdf$/i.test(String(name || ''));
    if (!isPdf) return null;

    return {
      id: file.getId(),
      url: file.getUrl(),
      name,
      normalized: norm_(name),
      digits: onlyDigits_(name),
      file,
    };
  } catch (err) {
    Logger.log('Falha ao indexar arquivo de certificado: ' + err);
    return null;
  }
}

function findCertificateFileForRow_(fileIndex, row, map) {
  if (!fileIndex || !fileIndex.entries || !fileIndex.entries.length) return null;

  const cpf = map.cpf > -1 ? onlyDigits_(row[map.cpf]) : '';
  if (cpf) {
    const hit = fileIndex.entries.find(entry => entry.digits && entry.digits.includes(cpf));
    if (hit) return hit;
  }

  const name = map.nome > -1 ? norm_(row[map.nome]) : '';
  if (name) {
    const exact = fileIndex.entries.find(entry => entry.normalized === name);
    if (exact) return exact;

    const hit = fileIndex.entries.find(entry => entry.normalized.includes(name));
    if (hit) return hit;

    const tokens = name.split(' ').filter(Boolean);
    if (tokens.length >= 2) {
      const first = tokens[0];
      const last = tokens[tokens.length - 1];
      const partial = fileIndex.entries.find(entry => entry.normalized.includes(first) && entry.normalized.includes(last));
      if (partial) return partial;
    }
  }

  return null;
}

function extractDriveFileIdFromUrl_(url) {
  const raw = sanitize_(url);
  if (!raw) return '';

  const patterns = [
    /\/d\/([a-zA-Z0-9_-]+)/,
    /id=([a-zA-Z0-9_-]+)/,
    /\/(?:file|folders)\/([a-zA-Z0-9_-]+)/,
  ];
  for (let i = 0; i < patterns.length; i++) {
    const match = raw.match(patterns[i]);
    if (match && match[1]) return match[1];
  }

  if (/^[a-zA-Z0-9_-]{10,}$/.test(raw)) return raw;
  return '';
}

function buildCertificateEmailPayload_(row, map, fileEntry) {
  const nome = map.nome > -1 ? sanitize_(row[map.nome]) : '';
  const curso = map.curso > -1 ? sanitize_(row[map.curso]) : '';
  const localEvento = map.local > -1 ? sanitize_(row[map.local]) : '';
  const ciclo = map.ciclo > -1 ? sanitize_(row[map.ciclo]) : '';
  const empresa = map.empresa > -1 ? sanitize_(row[map.empresa]) : NOME_EMPRESA_DEFAULT;
  const status = map.status > -1 ? sanitize_(row[map.status]) : '';

  const template = HtmlService.createHtmlOutputFromFile(CERTIFICATE_EMAIL_TEMPLATE_FILE).getContent();
  const replacements = {
    NOME: escapeHtml_(nome),
    CURSO: escapeHtml_(curso),
    LOCAL_EVENTO: escapeHtml_(localEvento || 'Online'),
    CICLO: escapeHtml_(ciclo || '—'),
    STATUS: escapeHtml_(status || ''),
    NOME_EMPRESA: escapeHtml_(empresa || NOME_EMPRESA_DEFAULT),
    CERTIFICATE_LINK: escapeHtml_(fileEntry.url),
  };

  const htmlBody = replaceCertificatePlaceholders_(template, replacements);
  const subject = replaceCertificatePlaceholders_(CERTIFICATE_EMAIL_SUBJECT_TEMPLATE, {
    CURSO: curso || 'Programa de Certificações',
    NOME: nome || '',
    NOME_EMPRESA: empresa || NOME_EMPRESA_DEFAULT,
  });

  return {
    subject,
    htmlBody,
    senderName: empresa || NOME_EMPRESA_DEFAULT,
  };
}

function replaceCertificatePlaceholders_(template, replacements) {
  return Object.keys(replacements).reduce((acc, key) => {
    const value = replacements[key];
    const pattern = new RegExp(`{{${key}}}`, 'g');
    return acc.replace(pattern, value);
  }, template);
}

function getSelectedRowIndexes_(sheet) {
  try {
    const range = sheet.getActiveRange();
    if (!range) return [];
    const start = range.getRow();
    const count = range.getNumRows();
    const rows = [];
    for (let i = 0; i < count; i++) {
      rows.push(start + i);
    }
    return rows;
  } catch (err) {
    return [];
  }
}
