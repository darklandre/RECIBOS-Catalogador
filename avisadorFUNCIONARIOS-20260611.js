/**
 * AVISADOR DE FUNCIONÁRIOS — recibos de vencimento em falta.
 *
 * Corre por TRIGGER (definido no acionador do Apps Script), tipicamente a partir
 * de ~dia 10, quando o recibo do mês já deveria ter sido enviado.
 *
 * Para cada colaborador ativo, no MÊS CORRENTE:
 *   - Se a célula do mês na planilha de registos TEM conteúdo  -> ok, nada a fazer.
 *   - Se está VAZIA -> vai verificar a pasta do Drive:
 *       · existe na pasta mas NÃO na folha -> ANOMALIA (falha do catalogador): avisa o ADMIN.
 *       · não existe em lado nenhum         -> avisa o COLABORADOR (lembrete).
 *
 * Config e helpers (_config, encontraColunaNoCabecalho, LINHA_CABECALHO_*) vêm do
 * catalogador — ambos os ficheiros vivem no mesmo projeto Apps Script.
 */
function avisarFuncionariosSemRecibo() {
  const cfg = _config();

  // Mês alvo = mês corrente (o trigger corre depois de o recibo já dever ter chegado).
  // Se algum dia precisares de verificar o mês ANTERIOR, ajusta aqui.
  const hoje = new Date();
  const mes = hoje.getMonth() + 1;
  const ano = hoje.getFullYear().toString();
  const mesFormatado = mes.toString().padStart(2, "0");
  const mesAno = `${mesFormatado}/${ano}`;

  const colaboradores = carregarColaboradoresParaAviso();

  const planilhaReg = SpreadsheetApp.openById(cfg.PLANILHA_REGISTROS_ID);
  const abaAno = planilhaReg.getSheetByName(ano);

  const lembretes = []; // avisados (colaborador) — em falta nos dois
  const anomalias = []; // na pasta mas não na folha
  const semEmail = [];  // em falta mas sem "Email pessoal"

  colaboradores.forEach(col => {
    if (temConteudoNaExcel(abaAno, col.sigla, mesAno)) return; // já registado/tratado

    const naPasta = existeReciboNaPasta(cfg, col.sigla, mesFormatado, ano);

    if (naPasta) {
      anomalias.push(col); // existe ficheiro mas não está na folha -> não é suposto
      return;
    }

    // Em falta nos dois -> lembrar o colaborador
    if (col.email) {
      enviarLembreteColaborador(col, mesAno, cfg);
      lembretes.push(col);
    } else {
      semEmail.push(col);
    }
  });

  enviarResumoAdmin(cfg, mesAno, lembretes, anomalias, semEmail);
  Logger.log(`Avisador ${mesAno}: ${lembretes.length} lembrete(s), ${anomalias.length} anomalia(s), ${semEmail.length} sem email.`);
}

/**
 * Carrega colaboradores ativos com Sigla, Nome e Email pessoal (aba "Ativos").
*/
function carregarColaboradoresParaAviso() {
  const cfg = _config();
  const planilha = SpreadsheetApp.openById(cfg.PLANILHA_COLABORADORES_ID);
  const aba = planilha.getSheetByName("Ativos");
  if (!aba) throw new Error("A aba 'Ativos' não foi encontrada.");

  const colSigla = encontraColunaNoCabecalho(aba, "Sigla", LINHA_CABECALHO_COLABORADORES);
  const colNome = encontraColunaNoCabecalho(aba, "Nome", LINHA_CABECALHO_COLABORADORES);
  const colEmail = encontraColunaNoCabecalho(aba, "Email pessoal", LINHA_CABECALHO_COLABORADORES);
  if (colSigla === -1 || colNome === -1 || colEmail === -1) {
    throw new Error("Colunas 'Sigla', 'Nome' ou 'Email pessoal' não encontradas na aba Ativos.");
  }

  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha <= LINHA_CABECALHO_COLABORADORES) return [];
  const numLinhas = ultimaLinha - LINHA_CABECALHO_COLABORADORES;
  const inicio = LINHA_CABECALHO_COLABORADORES + 1;

  const siglas = aba.getRange(inicio, colSigla, numLinhas, 1).getValues();
  const nomes = aba.getRange(inicio, colNome, numLinhas, 1).getValues();
  const emails = aba.getRange(inicio, colEmail, numLinhas, 1).getDisplayValues();

  const lista = [];
  for (let i = 0; i < numLinhas; i++) {
    const sigla = (siglas[i][0] || "").toString().trim().toUpperCase();
    if (!sigla) continue;
    lista.push({
      sigla: sigla,
      nome: (nomes[i][0] || "").toString().trim(),
      email: (emails[i][0] || "").toString().trim()
    });
  }
  return lista;
}

/**
 * True se a célula (sigla x mês) na aba do ano tiver QUALQUER conteúdo
 * (LINK, "Saiu a...", "FALTA", "(recusou)", ...). Vazia = em falta.
*/
function temConteudoNaExcel(abaAno, sigla, mesAno) {
  if (!abaAno) return false;
  const colMes = encontraColunaNoCabecalho(abaAno, mesAno, LINHA_CABECALHO_REGISTROS);
  if (colMes === -1) return false;
  const linha = encontraLinhaSiglaRegistos(abaAno, sigla);
  if (linha === -1) return false;
  return abaAno.getRange(linha, colMes).getDisplayValue().toString().trim() !== "";
}

/**
 * Encontra a linha da sigla na aba de registos (sigla na coluna B). -1 se não existir.
*/
function encontraLinhaSiglaRegistos(abaAno, sigla) {
  const valores = abaAno.getRange("B:B").getValues();
  for (let i = 0; i < valores.length; i++) {
    if ((valores[i][0] || "").toString().trim().toUpperCase() === sigla) return i + 1;
  }
  return -1;
}

/**
 * True se existir um ficheiro REC_<sigla>_<MM><YYYY>... na pasta do mês.
 * SOMENTE leitura: nunca cria pastas (ao contrário do obterPasta do catalogador).
*/
function existeReciboNaPasta(cfg, sigla, mesFormatado, ano) {
  const pastaAno = obterSubpastaSeExistir(cfg.PASTA_DESTINO_ID, ano);
  if (!pastaAno) return false;
  const pastaMes = obterSubpastaSeExistir(pastaAno.getId(), `${mesFormatado}/${ano}`);
  if (!pastaMes) return false;

  const prefixo = `REC_${sigla}_${mesFormatado}${ano}`.toUpperCase();
  const files = pastaMes.getFiles();
  while (files.hasNext()) {
    if (files.next().getName().toUpperCase().indexOf(prefixo) === 0) return true;
  }
  return false;
}

/**
 * Devolve a subpasta com este nome, ou null se não existir (não cria nada).
*/
function obterSubpastaSeExistir(pastaPaiId, nome) {
  const pastas = DriveApp.getFolderById(pastaPaiId).getFoldersByName(nome);
  return pastas.hasNext() ? pastas.next() : null;
}

/**
 * Envia o lembrete ao colaborador.
*/
function enviarLembreteColaborador(col, mesAno, cfg) {
  const assunto = `[${cfg.NOME_EMPRESA}] Lembrete: falta o teu recibo de vencimento de ${mesAno}`;
  const corpo =
    `Olá ${col.nome || col.sigla},\n\n` +
    `Ainda não consta o teu recibo de vencimento de ${mesAno} nos nossos registos.\n\n` +
    `Se já o tens, por favor reenvia o PDF assinado para documentos@arrowplus.pt o quanto antes.\n` +
    `Se ainda não o recebeste, podes ignorar este email — voltaremos a lembrar mais tarde.\n\n` +
    `Obrigado,\n${cfg.NOME_EMPRESA} — Financeiro`;
  MailApp.sendEmail(col.email, assunto, corpo);
  Logger.log(`Lembrete enviado a ${col.sigla} <${col.email}> (${mesAno}).`);
}

/**
 * Envia um resumo ao admin (só se houver algo a reportar).
*/
function enviarResumoAdmin(cfg, mesAno, lembretes, anomalias, semEmail) {
  if (lembretes.length === 0 && anomalias.length === 0 && semEmail.length === 0) {
    Logger.log(`Avisador ${mesAno}: tudo em ordem, nada a reportar ao admin.`);
    return;
  }

  let corpo = `Avisador de recibos de vencimento — ${mesAno} (${cfg.NOME_EMPRESA})\n\n`;

  if (lembretes.length) {
    corpo += `Lembretes enviados aos colaboradores (recibo em falta):\n` +
      lembretes.map(c => `  • ${c.sigla} — ${c.nome} <${c.email}>`).join("\n") + "\n\n";
  }
  if (anomalias.length) {
    corpo += `ATENÇÃO — recibo ESTÁ na pasta do Drive mas NÃO na folha (verificar catalogador):\n` +
      anomalias.map(c => `  • ${c.sigla} — ${c.nome}`).join("\n") + "\n\n";
  }
  if (semEmail.length) {
    corpo += `Em falta mas SEM "Email pessoal" preenchido (não foi possível avisar):\n` +
      semEmail.map(c => `  • ${c.sigla} — ${c.nome}`).join("\n") + "\n\n";
  }

  MailApp.sendEmail(cfg.EMAIL_NOTIFICACAO, `[${cfg.NOME_EMPRESA}] Avisador de recibos — ${mesAno}`, corpo);
}
