/**
 * AVISADOR DE FUNCIONÁRIOS — recibos de vencimento em falta.
 *
 * Corre por TRIGGER (definido no acionador do Apps Script), tipicamente a partir
 * de ~dia 10.
 *
 * Verifica TODOS os meses ANTERIORES ao mês corrente (no ano corrente). Para cada
 * colaborador ativo e cada mês alvo:
 *   - Célula tratada (TEM conteúdo OU fundo PRETO) -> ok, nada a fazer.
 *   - Vazia (e não-preta) -> vai verificar a pasta do Drive:
 *       · existe na pasta mas NÃO na folha -> ANOMALIA (falha do catalogador): avisa o ADMIN.
 *       · não existe em lado nenhum         -> conta como falta -> avisa o COLABORADOR.
 *
 * Os meses em falta de cada colaborador são agregados num ÚNICO email.
 * Config e helpers (_config, encontraColunaNoCabecalho, LINHA_CABECALHO_*) vêm do
 * catalogador — ambos os ficheiros vivem no mesmo projeto Apps Script.
 */
function avisarFuncionariosSemRecibo() {
  const cfg = _config();

  const hoje = new Date();
  const mesAtual = hoje.getMonth() + 1;
  const ano = hoje.getFullYear().toString();

  // Todos os meses ANTERIORES ao atual, no ano corrente.
  const mesesAlvo = [];
  for (let m = 1; m < mesAtual; m++) mesesAlvo.push(m);
  if (mesesAlvo.length === 0) {
    Logger.log(`Avisador: ${mesAtual}/${ano} — não há meses anteriores no ano para verificar.`);
    return;
  }

  const colaboradores = carregarColaboradoresParaAviso();
  const planilhaReg = SpreadsheetApp.openById(cfg.PLANILHA_REGISTROS_ID);
  const idx = construirIndiceRegistos(planilhaReg.getSheetByName(ano));

  const lembretes = []; // {col, meses:[]}
  const anomalias = []; // {col, meses:[]}
  const semEmail = [];  // {col, meses:[]}
  const cachePastas = {};

  colaboradores.forEach(col => {
    const emFalta = [];   // em falta nos dois (folha + pasta)
    const naPastaSo = []; // na pasta mas não na folha

    mesesAlvo.forEach(m => {
      const mesFmt = m.toString().padStart(2, "0");
      const mesAno = `${mesFmt}/${ano}`;
      if (celulaTratada(idx, col.sigla, mesAno)) return;
      if (existeReciboNaPasta(cfg, col.sigla, mesFmt, ano, cachePastas)) naPastaSo.push(mesAno);
      else emFalta.push(mesAno);
    });

    if (emFalta.length) {
      if (col.email) {
        enviarLembreteColaborador(col, emFalta, cfg);
        lembretes.push({ col: col, meses: emFalta });
      } else {
        semEmail.push({ col: col, meses: emFalta });
      }
    }
    if (naPastaSo.length) anomalias.push({ col: col, meses: naPastaSo });
  });

  enviarResumoAdmin(cfg, ano, lembretes, anomalias, semEmail);
  Logger.log(`Avisador ${ano}: ${lembretes.length} colaborador(es) lembrado(s), ${anomalias.length} com anomalia, ${semEmail.length} sem email.`);
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
 * Lê a aba do ano de uma vez e devolve um índice para consultas rápidas:
 *   { colPorMes: {"MM/AAAA": colIdx0}, linhaPorSigla: {SIGLA: rowIdx0}, valores, fundos }
 * (rowIdx0/colIdx0 são relativos ao bloco de dados, já sem a linha de cabeçalho).
 * Devolve null se a aba não existir / estiver vazia.
*/
function construirIndiceRegistos(abaAno) {
  if (!abaAno) return null;
  const lastRow = abaAno.getLastRow();
  const lastCol = abaAno.getLastColumn();
  if (lastRow <= LINHA_CABECALHO_REGISTROS) return null;

  const cabecalho = abaAno.getRange(LINHA_CABECALHO_REGISTROS, 1, 1, lastCol).getDisplayValues()[0];
  const colPorMes = {};
  cabecalho.forEach((h, i) => {
    const v = (h || "").toString().trim();
    if (/^\d{2}\/\d{4}$/.test(v)) colPorMes[v] = i; // índice 0-based
  });

  const numRows = lastRow - LINHA_CABECALHO_REGISTROS;
  const range = abaAno.getRange(LINHA_CABECALHO_REGISTROS + 1, 1, numRows, lastCol);
  const valores = range.getDisplayValues();
  const fundos = range.getBackgrounds();

  const linhaPorSigla = {};
  for (let r = 0; r < numRows; r++) {
    const sig = (valores[r][1] || "").toString().trim().toUpperCase(); // coluna B = índice 1
    if (sig && !(sig in linhaPorSigla)) linhaPorSigla[sig] = r;
  }

  return { colPorMes: colPorMes, linhaPorSigla: linhaPorSigla, valores: valores, fundos: fundos };
}

/**
 * True se a célula (sigla x mês) estiver TRATADA: tem conteúdo OU fundo preto.
*/
function celulaTratada(idx, sigla, mesAno) {
  if (!idx) return false;
  const col = idx.colPorMes[mesAno];
  if (col === undefined) return false;       // coluna do mês não existe -> não tratada
  const linha = idx.linhaPorSigla[sigla];
  if (linha === undefined) return false;     // colaborador sem linha -> não tratada
  const valor = (idx.valores[linha][col] || "").toString().trim();
  if (valor !== "") return true;             // tem conteúdo (LINK, "Saiu", "FALTA", ...)
  const bg = (idx.fundos[linha][col] || "").toLowerCase();
  return bg === "#000000" || bg === "black"; // fundo preto = tratada
}

/**
 * True se existir um ficheiro REC_<sigla>_<MM><AAAA>... na pasta do mês.
 * SOMENTE leitura (nunca cria pastas). Faz cache da listagem por mês.
*/
function existeReciboNaPasta(cfg, sigla, mesFmt, ano, cache) {
  const chave = `${ano}/${mesFmt}`;
  let nomes = cache[chave];
  if (nomes === undefined) {
    nomes = [];
    const pastaAno = obterSubpastaSeExistir(cfg.PASTA_DESTINO_ID, ano);
    if (pastaAno) {
      const pastaMes = obterSubpastaSeExistir(pastaAno.getId(), `${mesFmt}/${ano}`);
      if (pastaMes) {
        const files = pastaMes.getFiles();
        while (files.hasNext()) nomes.push(files.next().getName().toUpperCase());
      }
    }
    cache[chave] = nomes;
  }
  const prefixo = `REC_${sigla}_${mesFmt}${ano}`.toUpperCase();
  return nomes.some(n => n.indexOf(prefixo) === 0);
}

/**
 * Devolve a subpasta com este nome, ou null se não existir (não cria nada).
*/
function obterSubpastaSeExistir(pastaPaiId, nome) {
  const pastas = DriveApp.getFolderById(pastaPaiId).getFoldersByName(nome);
  return pastas.hasNext() ? pastas.next() : null;
}

/**
 * Envia o lembrete ao colaborador (agrega todos os meses em falta).
*/
function enviarLembreteColaborador(col, meses, cfg) {
  const lista = meses.join(", ");
  const linhaFalta = meses.length === 1
    ? `O teu recibo de vencimento de ${lista} ainda não está na BD.`
    : `Os teus recibos de vencimento dos meses ${lista} ainda não estão na BD.`;

  const assunto = `[${cfg.NOME_EMPRESA}] Recibo${meses.length > 1 ? "s" : ""} de vencimento em falta`;
  const corpo =
    `Olá colega ${col.nome || col.sigla},\n\n` +
    `${linhaFalta}\n\n` +
    `Se já o tens, por favor reenvia o PDF assinado com cartão de cidadão para documentos@arrowplus.pt o quanto antes, com o nome correto (ver manual de processos).\n` +
    `Se ainda não o recebeste, podes ignorar este email — voltaremos a lembrar mais tarde.\n\n` +
    `Obrigado,\nFERH`;

  MailApp.sendEmail(col.email, assunto, corpo);
  Logger.log(`Lembrete enviado a ${col.sigla} <${col.email}>: ${lista}.`);
}

/**
 * Envia um resumo ao admin (só se houver algo a reportar).
*/
function enviarResumoAdmin(cfg, ano, lembretes, anomalias, semEmail) {
  if (!lembretes.length && !anomalias.length && !semEmail.length) {
    Logger.log(`Avisador ${ano}: tudo em ordem, nada a reportar ao admin.`);
    return;
  }

  let corpo = `Avisador de recibos de vencimento — ano ${ano} (${cfg.NOME_EMPRESA})\n\n`;

  if (lembretes.length) {
    corpo += `Lembretes enviados aos colaboradores (recibo em falta):\n` +
      lembretes.map(x => `  • ${x.col.sigla} — ${x.col.nome}: ${x.meses.join(", ")}`).join("\n") + "\n\n";
  }
  if (anomalias.length) {
    corpo += `ATENÇÃO — recibo ESTÁ na pasta do Drive mas NÃO na folha (verificar catalogador):\n` +
      anomalias.map(x => `  • ${x.col.sigla} — ${x.col.nome}: ${x.meses.join(", ")}`).join("\n") + "\n\n";
  }
  if (semEmail.length) {
    corpo += `Em falta mas SEM "Email pessoal" preenchido (não foi possível avisar):\n` +
      semEmail.map(x => `  • ${x.col.sigla} — ${x.col.nome}: ${x.meses.join(", ")}`).join("\n") + "\n\n";
  }

  MailApp.sendEmail(cfg.EMAIL_NOTIFICACAO, `[${cfg.NOME_EMPRESA}] Avisador de recibos — ${ano}`, corpo);
}
