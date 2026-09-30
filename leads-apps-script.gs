// PollySim — recebe os leads do site, grava no Google Sheets e avisa por e-mail.
const ABA = 'Leads';
const DESTINO = 'contato@pollypet.com.br';
const CAB = ['Cadastro em','Nome','E-mail','Telefone','Consentimento LGPD','Última etapa','Último paciente','Acertos','XP total','Nível','Pacientes atendidos','Interesse no curso','Atualizado em'];

// evita que textos iniciados por = + - @ virem fórmulas na planilha
const limpa = v => { v = (v == null ? '' : String(v)).slice(0, 300); return /^[=+\-@]/.test(v) ? "'" + v : v; };

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = JSON.parse(e.postData.contents);
    if (!p.email || !p.nome) return ContentService.createTextOutput('ignorado');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(ABA) || ss.insertSheet(ABA);
    if (sh.getLastRow() === 0) { sh.appendRow(CAB); sh.setFrozenRows(1); }
    const agora = new Date();
    const n = sh.getLastRow() - 1;
    const emails = n > 0 ? sh.getRange(2, 3, n, 1).getValues().flat().map(x => String(x).toLowerCase()) : [];
    const i = emails.indexOf(String(p.email).toLowerCase());
    const interesse = String(p.etapa || '').indexOf('INTERESSE') === 0;
    let novo = false;
    if (i < 0) {
      novo = true;
      sh.appendRow([agora, limpa(p.nome), limpa(p.email), limpa(p.telefone), limpa(p.consentimento_LGPD), limpa(p.etapa),
        limpa(p.paciente), limpa(p.acertos), p.xp_total || '', limpa(p.nivel), p.pacientes_atendidos || '', interesse ? 'SIM' : '', agora]);
    } else {
      const r = i + 2, row = sh.getRange(r, 1, 1, 13).getValues()[0];
      row[5] = limpa(p.etapa);
      if (p.paciente) row[6] = limpa(p.paciente);
      if (p.acertos) row[7] = limpa(p.acertos);
      if (p.xp_total) row[8] = p.xp_total;
      if (p.nivel) row[9] = limpa(p.nivel);
      if (p.pacientes_atendidos) row[10] = p.pacientes_atendidos;
      if (interesse) row[11] = 'SIM';
      row[12] = agora;
      sh.getRange(r, 1, 1, 13).setValues([row]);
    }
    // e-mail só no cadastro novo e quando o aluno clica em "Quero conhecer o curso"
    if (novo || interesse) {
      MailApp.sendEmail(DESTINO, (interesse ? 'LEAD QUENTE (quer o curso) — ' : 'Novo lead PollySim — ') + p.nome,
        ['Nome: ' + p.nome, 'E-mail: ' + p.email, 'Telefone: ' + p.telefone, 'Etapa: ' + p.etapa,
         'XP: ' + (p.xp_total || '-'), 'Nível: ' + (p.nivel || '-'), 'Pacientes atendidos: ' + (p.pacientes_atendidos || '-'),
         'Consentimento LGPD: ' + p.consentimento_LGPD].join('\n'));
    }
    return ContentService.createTextOutput('ok');
  } finally { lock.releaseLock(); }
}
