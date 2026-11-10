// Recebe as simulações enviadas pelo site e grava uma linha na aba "Simulações".
// Projeto do Apps Script implantado como "App da Web" (executar como: eu; acesso: qualquer pessoa).

const PLANILHA_ID = '17N7Ux9hMcOCGiaoxvuRb-0nUmxsXOwoC0A-YiFdP2a8';
const ABA = 'Simulações';
const EVENTOS = ['simulação', 'pdf', 'copiar'];

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (EVENTOS.indexOf(d.evento) < 0) return ContentService.createTextOutput('ignorado');

  const num = v => { const n = Number(v); return isFinite(n) ? n : ''; };
  const txt = v => "'" + String(v == null ? '' : v).slice(0, 80);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(ABA).appendRow([
      new Date(), txt(d.sessao), d.evento, txt(d.cidade), txt(d.empreendimento), txt(d.entrega),
      num(d.valor), num(d.nNormal), num(d.parcelaNormal), num(d.totalNormal),
      num(d.nBF), num(d.parcelaBF), num(d.economia), num(d.economiaPct),
    ]);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}
