/**
 * Qualification & Arbitrage des Projets WebApp — backend Google Apps Script.
 * Sert index.html et stocke les projets dans une feuille Google Sheets partagée.
 */
const SHEET_NAME = 'Projets';
const HEADERS = ['id', 'date', 'statut', 'catégorie', 'titre', 'demandeur', 'responsable métier', 'json'];
const JSON_COL = HEADERS.length;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Qualification des Projets WebApp')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Retourne la feuille de données (script lié à un Sheet, sinon création d'un classeur dédié). */
function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) {
    const id = props.getProperty('SS_ID');
    if (id) ss = SpreadsheetApp.openById(id);
    else {
      ss = SpreadsheetApp.create('Qualification projets WebApp - données');
      props.setProperty('SS_ID', ss.getId());
    }
  }
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#E8E22F');
  }
  return sh;
}

function getProjects() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, JSON_COL, last - 1, 1).getValues()
    .map(r => { try { return JSON.parse(r[0]); } catch (e) { return null; } })
    .filter(Boolean)
    .sort((a, b) => String(b.date).localeCompare(String(a.date)));
}

const STATUS_LABELS_ = {
  GO: 'GO', GO_RESERVE: 'GO sous réserve - Scope restreint', NO_GO_DSI: 'NO-GO - Orienté DSI',
  NO_GO_REFERENT: 'NO-GO - Manque de Référent Métier', NO_GO_PREREQ: 'NO-GO - Prérequis bloquants manquants',
};

function saveProject(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    const row = [p.id, p.date, STATUS_LABELS_[p.verdict.status] || p.verdict.status, p.verdict.category,
      p.answers.title || '', p.answers.requester || '', p.answers.referentName || '', JSON.stringify(p)];
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
    const i = ids.indexOf(p.id);
    if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return true;
  } finally { lock.releaseLock(); }
}

function deleteProject(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    if (sh.getLastRow() < 2) return false;
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]);
    const i = ids.indexOf(id);
    if (i >= 0) sh.deleteRow(i + 2);
    return i >= 0;
  } finally { lock.releaseLock(); }
}
