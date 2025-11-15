/** Certificates.gs — versão completa para a PLANILHA ESPELHO (isolada) **/
(function (global) {
  'use strict';

  /******************************
   * CONFIG — AJUSTE AQUI
   ******************************/
  const ORIGIN_SID  = '1gCBRIGT1sFXlHQPCdWai0Mn88ATJtS7fjDyHqg32mrw'; // Planilha ORIGEM
  const MIRROR_SID  = '1kVX0TH9_lM7e2nceMi6OktrGUZ0LootvVTLEx3AFOIo'; // Planilha ESPELHO
  const CERTS_FOLDER_ID = '1Ja_jnk8_0PFmse8aIQ66X1vqmaOBp1mH';          // Pasta com PDFs

  // Na ORIGEM exigimos essa aba; na ESPELHO aceitamos qualquer aba
  const ORIGIN_REQUIRED_SHEET_NAME = 'BASE DE DADOS CADASTRAIS';

  // Título do menu
  const MENU_TITLE = 'Certificados';

  // Colunas e status
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

  // E-mail
  const CERTIFICATE_EMAIL_TEMPLATE_FILE = 'certificate-email'; // opcional (usa fallback se não existir)
  const CERTIFICATE_EMAIL_SUBJECT_TEMPLATE = 'Certificado disponível — {{CURSO}}';
  const DEFAULT_SENDER_NAME = 'Coordenação';

  // Alvos habilitados (origem exige aba; espelho qualquer aba)
  const CERTIFICATE_TARGETS = [
    { sid: ORIGIN_SID, sheet: ORIGIN_REQUIRED_SHEET_NAME, folderId: CERTS_FOLDER_ID },
    { sid: MIRROR_SID, sheet: '', /* curinga */           folderId: CERTS_FOLDER_ID },
  ];

  /******************************
   * MENU
   ******************************/
  function onOpen() {
    try { maybeAddCertificateMenu_(); } catch (e) { Logger.log(e); }
  }

  function maybeAddCertificateMenu_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    const sid = ss.getId();
    const sheetName = ss.getActiveSheet().getName();

    const enabled = CERTIFICATE_TARGETS.some(t =>
      t.sid === sid && (!t.sheet || t.sheet === sheetName)
    );
    if (!enabled) return;

    SpreadsheetApp.getUi()
      .createMenu(MENU_TITLE)
      .addItem('Atualizar links (planilha atual)', 'menuUpdateCertificateLinks')
      .addItem('Enviar certificados pendentes', 'menuSendCertificateEmails')
      .addSeparator()
      .addItem('Reenviar certificados selecionados', 'menuResendSelectedCertificates')
      .addToUi();
  }

  function menuUpdateCertificateLinks() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const target = getTargetForSheet_(sheet);
    if (!target) return SpreadsheetApp.getUi().alert('Esta aba não está habilitada para certificados.');
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
    const target = getTargetForSheet_(sheet);
    if (!target) return SpreadsheetApp.getUi().alert('Esta aba não está habilitada para certificados.');
    const ui = SpreadsheetApp.getUi();
    if (ui.alert('Envio de certificados', 'Deseja enviar os certificados pendentes?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    const result = sendCertificateEmailsForSheet_(sheet, target, { force:false, onlySelection:false });
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
    const target = getTargetForSheet_(sheet);
    if (!target) return SpreadsheetApp.getUi().alert('Esta aba não está habilitada para certificados.');
    const ui = SpreadsheetApp.getUi();
    if (ui.alert('Reenvio de certificados', 'Reenviar apenas para as linhas selecionadas?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    const result = sendCertificateEmailsForSheet_(sheet, target, { force:true, onlySelection:true });
    ui.alert(
      'Reenvio concluído!\n' +
      `Certificados enviados: ${result.sent}\n` +
      `Sem certificado localizado: ${result.notFound}\n` +
      `Com erro de envio: ${result.errors}\n` +
      `Ignorados: ${result.skipped}`
    );
  }
  function isSecondCopy_(statusRaw) {
  const s = String(statusRaw || '').toLowerCase();
  return /\b(segunda|2a|2ª)\s*via\b/.test(s); // pega "SEGUNDA VIA", "2a via", "2ª via"
}

function formatDateBR_(d) {
  const tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';
  const dia = Utilities.formatDate(d, tz, 'dd/MM/yyyy');
  const hora = Utilities.formatDate(d, tz, 'HH:mm');
  return `${dia} às ${hora}`;
}


  /******************************
   * CORE
   ******************************/
  function getTargetForSheet_(sheet) {
    if (!sheet) return null;
    try {
      const sid = sheet.getParent().getId();
      const name = sheet.getName();
      return CERTIFICATE_TARGETS.find(t => t.sid === sid && (!t.sheet || t.sheet === name)) || null;
    } catch (_) { return null; }
  }

  function ensureCertificateColumns_(sheet) {
    const target = getTargetForSheet_(sheet);
    if (!target) return; // só cria se habilitado
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
        break;
      }
      if (!updated) { sheet.setFrozenRows(1); return; }
    }
    sheet.setFrozenRows(1);
  }

  function updateCertificateLinksForSheet_(sheet, target, options) {
    ensureCertificateColumns_(sheet);
    const folder = getFolder_(target.folderId);
    if (!folder) throw new Error('Não foi possível acessar a pasta de certificados.');

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { matched:0, notFound:0, skipped:0 };

    const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
    const map = headerMap_(header);

    const linkIdx   = map.certLink;
    const statusIdx = map.certStatus;
    const sentAtIdx = map.certSentAt;
    const errorIdx  = map.certError;
    if (linkIdx === -1) throw new Error('Coluna "CERTIFICADO (PDF)" não localizada.');

    const range  = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = range.getValues();
    const index  = buildFileIndex_(folder);

    const force = !!(options && options.force);
    let matched=0, notFound=0, skipped=0;
    const rowsToSync = [];

    values.forEach((row, i) => {
      const hasCpf  = map.cpf  > -1 ? !!onlyDigits_(row[map.cpf])  : false;
      const hasNome = map.nome > -1 ? !!sanitize_(row[map.nome])   : false;
      if (!hasCpf && !hasNome) { skipped++; return; }

      const currentLink   = sanitize_(row[linkIdx]);
      const currentStatus = statusIdx > -1 ? sanitize_(row[statusIdx]) : '';
      if (!force && currentLink) { skipped++; return; }

      const origLink   = linkIdx   > -1 ? row[linkIdx]   : '';
      const origStatus = statusIdx  > -1 ? row[statusIdx] : '';
      const origSentAt = sentAtIdx  > -1 ? row[sentAtIdx] : '';
      const origError  = errorIdx   > -1 ? row[errorIdx]  : '';

      let changed = false;
      const file = findFileForRow_(index, row, map);

      if (file) {
        matched++;
        if (origLink !== file.url) { row[linkIdx] = file.url; changed = true; }
        if (statusIdx > -1 && currentStatus !== CERTIFICATE_STATUS_SENT && origStatus !== CERTIFICATE_STATUS_PENDING) {
          row[statusIdx] = CERTIFICATE_STATUS_PENDING; changed = true;
        }
        if (sentAtIdx > -1 && currentStatus !== CERTIFICATE_STATUS_SENT && origSentAt) { row[sentAtIdx] = ''; changed = true; }
        if (errorIdx > -1 && origError) { row[errorIdx] = ''; changed = true; }
      } else {
        notFound++;
        if (errorIdx > -1 && (!currentLink || force)) {
          const msg = 'Certificado não encontrado na pasta configurada.';
          if (origError !== msg) { row[errorIdx] = msg; changed = true; }
        }
        if (statusIdx > -1 && currentStatus !== CERTIFICATE_STATUS_SENT && origStatus !== CERTIFICATE_STATUS_ERROR) {
          row[statusIdx] = CERTIFICATE_STATUS_ERROR; changed = true;
        }
        if (sentAtIdx > -1 && origSentAt) { row[sentAtIdx] = ''; changed = true; }
      }

      if (changed) rowsToSync.push(i + 2);
    });

    range.setValues(values);
    SpreadsheetApp.flush();
    // nesta versão (espelho) não precisamos sincronizar para outros destinos
    return { matched, notFound, skipped };
  }

  function sendCertificateEmailsForSheet_(sheet, target, options) {
    ensureCertificateColumns_(sheet);
    const folder = getFolder_(target.folderId);
    if (!folder) throw new Error('Não foi possível acessar a pasta de certificados.');

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { sent:0, notFound:0, errors:0, skipped:0 };

    const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
    const map = headerMap_(header);

    const linkIdx   = map.certLink;
    const statusIdx = map.certStatus;
    const sentAtIdx = map.certSentAt;
    const errorIdx  = map.certError;
    if (linkIdx === -1)   throw new Error('Coluna "CERTIFICADO (PDF)" não localizada.');
    if (statusIdx === -1) throw new Error('Coluna "CERTIFICADO - STATUS" não localizada.');

    const selectionRows = getSelectedRowIndexes_(sheet);
    const useSelection = !!(options && options.onlySelection);
    const force        = !!(options && options.force);

    const range  = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = range.getValues();
    const index  = buildFileIndex_(folder);

    let sent=0, notFound=0, errors=0, skipped=0;

    values.forEach((row, i) => {
      const rowNumber = i + 2;
      if (useSelection && !selectionRows.includes(rowNumber)) { skipped++; return; }

      const emailIdx = map.email;
      const email = emailIdx > -1 ? sanitize_(row[emailIdx]) : '';
      if (!isValidEmail_(email)) { skipped++; return; }

      const hasCpf  = map.cpf  > -1 ? !!onlyDigits_(row[map.cpf])  : false;
      const hasNome = map.nome > -1 ? !!sanitize_(row[map.nome])   : false;
      if (!hasCpf && !hasNome) { skipped++; return; }

      const currentStatus = statusIdx > -1 ? sanitize_(row[statusIdx]) : '';
      const shouldSend = force || !currentStatus || currentStatus === CERTIFICATE_STATUS_PENDING || currentStatus === CERTIFICATE_STATUS_ERROR;
      if (!shouldSend) { skipped++; return; }

      let file = null;
      const link = sanitize_(row[linkIdx]);
      if (link) {
        const fileId = extractDriveId_(link);
        if (fileId) {
          const hit = index.byId[fileId];
          if (hit) file = hit;
          else {
            try {
              const fetched = DriveApp.getFileById(fileId);
              const entry = entryFromFile_(fetched);
              if (entry) { file = entry; index.entries.push(entry); index.byId[entry.id] = entry; }
            } catch (_) { /* tenta por CPF/Nome */ }
          }
        }
      }
      if (!file) file = findFileForRow_(index, row, map);

      const origLink   = linkIdx   > -1 ? row[linkIdx]   : '';
      const origStatus = statusIdx  > -1 ? row[statusIdx] : '';
      const origSentAt = sentAtIdx  > -1 ? row[sentAtIdx] : '';
      const origError  = errorIdx   > -1 ? row[errorIdx]  : '';

      let changed = false;

      if (!file) {
        notFound++;
        if (statusIdx > -1 && origStatus !== CERTIFICATE_STATUS_ERROR) { row[statusIdx] = CERTIFICATE_STATUS_ERROR; changed = true; }
        if (sentAtIdx > -1 && origSentAt) { row[sentAtIdx] = ''; changed = true; }
        if (errorIdx > -1) {
          const msg = 'Certificado não encontrado na pasta configurada.';
          if (origError !== msg) { row[errorIdx] = msg; changed = true; }
        }
        if (changed) range.setValues(values);
        return;
      }

      try {
        const payload = buildEmailPayload_(row, map, file);
        GmailApp.sendEmail(email, payload.subject, ' ', {
          name: payload.senderName,
          htmlBody: payload.htmlBody,
          attachments: [file.file.getAs(MimeType.PDF)],
        });

        sent++;
        if (statusIdx > -1 && origStatus !== CERTIFICATE_STATUS_SENT) { row[statusIdx] = CERTIFICATE_STATUS_SENT; changed = true; }
        if (sentAtIdx > -1) { row[sentAtIdx] = new Date(); changed = true; }
        if (errorIdx > -1 && origError) { row[errorIdx] = ''; changed = true; }
        if (linkIdx > -1 && !sanitize_(origLink)) { row[linkIdx] = file.url; changed = true; }
      } catch (err) {
        errors++;
        if (statusIdx > -1 && origStatus !== CERTIFICATE_STATUS_ERROR) { row[statusIdx] = CERTIFICATE_STATUS_ERROR; changed = true; }
        if (sentAtIdx > -1) { row[sentAtIdx] = ''; changed = true; }
        if (errorIdx > -1) {
          const msg = 'Falha ao enviar e-mail: ' + err;
          if (origError !== msg) { row[errorIdx] = msg; changed = true; }
        }
      }
    });

    range.setValues(values);
    SpreadsheetApp.flush();
    return { sent, notFound, errors, skipped };
  }

  /******************************
   * DRIVE/INDEX
   ******************************/
  function getFolder_(id) {
    const folderId = String(id || '').trim();
    if (!folderId) return null;
    try { return DriveApp.getFolderById(folderId); }
    catch (err) { Logger.log('Falha ao abrir pasta: ' + err); return null; }
  }

  function buildFileIndex_(folder) {
    const entries = [], byId = {};
    const it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      const e = entryFromFile_(f);
      if (!e) continue;
      entries.push(e);
      byId[e.id] = e;
    }
    return { entries, byId };
  }

  function entryFromFile_(file) {
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
    } catch (e) { Logger.log('Indexação falhou: ' + e); return null; }
  }

  function findFileForRow_(index, row, map) {
    if (!index || !index.entries || !index.entries.length) return null;

    const cpf = map.cpf > -1 ? onlyDigits_(row[map.cpf]) : '';
    if (cpf) {
      const hit = index.entries.find(en => en.digits && en.digits.includes(cpf));
      if (hit) return hit;
    }

    const name = map.nome > -1 ? norm_(row[map.nome]) : '';
    if (name) {
      const exact = index.entries.find(en => en.normalized === name);
      if (exact) return exact;
      const hit = index.entries.find(en => en.normalized.includes(name));
      if (hit) return hit;
      const tokens = name.split(' ').filter(Boolean);
      if (tokens.length >= 2) {
        const first = tokens[0], last = tokens[tokens.length - 1];
        const partial = index.entries.find(en => en.normalized.includes(first) && en.normalized.includes(last));
        if (partial) return partial;
      }
    }
    return null;
  }

  function extractDriveId_(url) {
    const raw = sanitize_(url);
    if (!raw) return '';
    const patterns = [/\/d\/([a-zA-Z0-9_-]+)/, /id=([a-zA-Z0-9_-]+)/, /\/(?:file|folders)\/([a-zA-Z0-9_-]+)/];
    for (let i=0;i<patterns.length;i++){ const m=raw.match(patterns[i]); if (m && m[1]) return m[1]; }
    if (/^[a-zA-Z0-9_-]{10,}$/.test(raw)) return raw;
    return '';
  }

  /******************************
   * E-MAIL
   ******************************/
  function buildEmailPayload_(row, map, fileEntry) {
  const nome    = map.nome    > -1 ? sanitize_(row[map.nome])    : '';
  const curso   = map.curso   > -1 ? sanitize_(row[map.curso])   : '';
  const localEv = map.local   > -1 ? sanitize_(row[map.local])   : '';
  const ciclo   = map.ciclo   > -1 ? sanitize_(row[map.ciclo])   : '';
  const empresa = map.empresa > -1 ? sanitize_(row[map.empresa]) : DEFAULT_SENDER_NAME;
  const status  = map.status  > -1 ? sanitize_(row[map.status])  : '';

  // Detecta “2ª via” pelo STATUS
  const segundaVia = isSecondCopy_(status);
  const carimbo = formatDateBR_(new Date());

  // Carrega template HTML, com fallback
  let template = '';
  try {
    template = HtmlService.createHtmlOutputFromFile(CERTIFICATE_EMAIL_TEMPLATE_FILE).getContent();
  } catch (_e) {
    // Fallback já com badge de 2ª via quando aplicável
    template =
      (segundaVia
        ? '<p style="margin:0 0 12px 0;"><span style="background:#ffd166;color:#7a4e00;font-weight:700;padding:6px 10px;border-radius:6px;">2ª via emitida em {{SEGUNDA_VIA_DATA}}</span></p>'
        : ''
      ) +
      '<p>Olá, {{NOME}}!</p>' +
      '<p>Seu certificado do curso <strong>{{CURSO}}</strong> {{CICLO}} {{LOCAL_EVENTO}} está disponível.</p>' +
      '<p>Acesse o PDF: <a href="{{CERTIFICATE_LINK}}">Abrir Certificado</a></p>' +
      '<p>Status: {{STATUS}}</p>' +
      (segundaVia ? '<p style="color:#555">Esta é uma reemissão (segunda via).</p>' : '') +
      '<p>Atenciosamente,<br>{{NOME_EMPRESA}}</p>';
  }

  // Se você tiver um template HTML próprio, pode (opcionalmente) inserir os marcadores abaixo nele:
  // {{SEGUNDA_VIA_BADGE}} e {{SEGUNDA_VIA_DATA}}
  // Ex.: colocar {{SEGUNDA_VIA_BADGE}} logo após a saudação.
  const segundaViaBadgeHtml = segundaVia
    ? `<p style="margin:0 0 12px 0;"><span style="background:#ffd166;color:#7a4e00;font-weight:700;padding:6px 10px;border-radius:6px;">2ª via emitida em ${carimbo}</span></p>`
    : '';

  const htmlBody = replaceTags_(template, {
    NOME: escapeHtml_(nome),
    CURSO: escapeHtml_(curso || 'Programa de Certificações'),
    LOCAL_EVENTO: escapeHtml_(localEv || 'Online'),
    CICLO: escapeHtml_(ciclo || '—'),
    STATUS: escapeHtml_(status || (segundaVia ? 'SEGUNDA VIA' : '')),
    NOME_EMPRESA: escapeHtml_(empresa || DEFAULT_SENDER_NAME),
    CERTIFICATE_LINK: escapeHtml_(fileEntry.url),
    SEGUNDA_VIA_BADGE: segundaViaBadgeHtml,
    SEGUNDA_VIA_DATA: carimbo
  });

  // Assunto com tag [2ª via]
  const baseSubject = replaceTags_(CERTIFICATE_EMAIL_SUBJECT_TEMPLATE, {
    CURSO: curso || 'Programa de Certificações',
    NOME: nome || '',
    NOME_EMPRESA: empresa || DEFAULT_SENDER_NAME,
  });
  const subject = segundaVia ? `[2ª via] ${baseSubject}` : baseSubject;

  return { subject, htmlBody, senderName: empresa || DEFAULT_SENDER_NAME };
}

  /******************************
   * HEADER MAP (flexível)
   ******************************/
  function headerMap_(header){
    const clean = s => (s || '').toString()
      .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
      .toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
    const H = header.map(clean);
    const find = (...labels) => {
      const L = labels.map(clean);
      return H.findIndex(h => L.includes(h));
    };
    return {
      // dados base
      cpf: find('cpf'),
      nome: find('nome completo sem abreviacoes','nome completo','nome'),
      email: find('e mail','email','e-mail'),
      curso: find('curso'),
      local: find('local do evento','local evento','local','cidade'),
      ciclo: find('ciclo','turma'),
      status: find('status'),
      empresa: find('empresa'),
      // certificados
      certLink: find('certificado (pdf)','certificado pdf','link certificado'),
      certStatus: find('certificado - status','certificado status'),
      certSentAt: find('certificado - enviado em','certificado enviado em','certificado envio'),
      certError: find('certificado - ultimo erro','certificado ultimo erro','certificado erro'),
    };
  }

  /******************************
   * HELPERS (privados)
   ******************************/
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
      const out = [];
      for (let i = 0; i < count; i++) out.push(start + i);
      return out;
    } catch (_) { return []; }
  }

  function sanitize_(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    return String(v).trim();
  }
  function onlyDigits_(v) { return sanitize_(v).replace(/\D+/g, ''); }
  function norm_(v) {
    return String(sanitize_(v))
      .toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
      .replace(/\s+/g,' ')
      .replace(/[^\w\s@.-]/g,'')
      .trim();
  }
  function isValidEmail_(email) {
    const e = sanitize_(email);
    return !!e && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);
  }
  function escapeHtml_(s) {
    const map = { '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' };
    return String(s).replace(/[&<>"']/g, m => map[m]);
  }
  function replaceTags_(template, obj) {
    return Object.keys(obj).reduce((acc, k) => acc.replace(new RegExp(`{{${k}}}`, 'g'), obj[k]), template);
  }

  /******************************
   * EXPORTA (menu)
   ******************************/
  global.certOnOpen = onOpen; // <— renomeado
global.menuUpdateCertificateLinks = menuUpdateCertificateLinks;
global.menuSendCertificateEmails = menuSendCertificateEmails;
global.menuResendSelectedCertificates = menuResendSelectedCertificates;

})(this);
function installCertMenuTrigger() {
  // remove antigos
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'certOnOpen') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // cria novo (executa certOnOpen a cada abertura dessa planilha)
  ScriptApp.newTrigger('certOnOpen')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
}
function certDebug() {
  const sh = SpreadsheetApp.getActiveSheet();
  const sid = sh.getParent().getId();
  const name = sh.getName();
  const enabled = (typeof this.CERTIFICATE_TARGETS !== 'undefined') && this.CERTIFICATE_TARGETS.some(t =>
    t.sid === sid && (!t.sheet || t.sheet === name)
  );
  SpreadsheetApp.getUi().alert(
    'DEBUG CERTIFICADOS\n' +
    `SID atual: ${sid}\n` +
    `ABA atual: ${name}\n` +
    `Habilitado pelos TARGETS? ${enabled ? 'SIM' : 'NÃO'}`
  );
}

