/**
 * CONFIGURAÇÃO — lida das Script Properties (Project Settings → Script Properties).
 * Mantém o código IDÊNTICO entre Darkland e Darkpurple; só as propriedades diferem.
 *
 * Chaves obrigatórias (definir em cada projeto Apps Script):
 *   PASTA_ORIGEM_ID            -> "0 - Por processar" (anexos do N8N de documentos@arrowplus.pt)
 *   PASTA_DESTINO_ID           -> "Recibos de vencimento" (onde ficam arquivados)
 *   PASTA_NAO_IDENTIFICADOS_ID -> subpasta para nomes errados / não identificados
 *   PLANILHA_COLABORADORES_ID  -> ficheiro de Colaboradores (aba "Ativos")
 *   PLANILHA_REGISTROS_ID      -> ficheiro de registos (links finais)
 *   EMAIL_NOTIFICACAO          -> email para o relatório
 *   NOME_EMPRESA               -> "DARKLAND" ou "DARKPURPLE" (vai no assunto do email)
 * Opcional:
 *   COR_LINK_PASTA             -> cor do link da pasta no cabeçalho (default "#1155cc")
 */
var __configCache = null;
function _config() {
  if (__configCache) return __configCache;
  var props = PropertiesService.getScriptProperties();
  __configCache = {
    PASTA_ORIGEM_ID:            props.getProperty("PASTA_ORIGEM_ID") || "",
    PASTA_DESTINO_ID:           props.getProperty("PASTA_DESTINO_ID") || "",
    PASTA_NAO_IDENTIFICADOS_ID: props.getProperty("PASTA_NAO_IDENTIFICADOS_ID") || "",
    PLANILHA_COLABORADORES_ID:  props.getProperty("PLANILHA_COLABORADORES_ID") || "",
    PLANILHA_REGISTROS_ID:      props.getProperty("PLANILHA_REGISTROS_ID") || "",
    EMAIL_NOTIFICACAO:          props.getProperty("EMAIL_NOTIFICACAO") || "",
    NOME_EMPRESA:               props.getProperty("NOME_EMPRESA") || "",
    COR_LINK_PASTA:             props.getProperty("COR_LINK_PASTA") || "#1155cc"
  };
  var obrigatorias = ["PASTA_ORIGEM_ID", "PASTA_DESTINO_ID", "PASTA_NAO_IDENTIFICADOS_ID",
    "PLANILHA_COLABORADORES_ID", "PLANILHA_REGISTROS_ID", "EMAIL_NOTIFICACAO", "NOME_EMPRESA"];
  var faltam = obrigatorias.filter(function (k) { return !__configCache[k]; });
  if (faltam.length) {
    throw new Error("Script Properties em falta: " + faltam.join(", ") +
      ". Configura em Project Settings → Script Properties.");
  }
  return __configCache;
}

//REC_<sigla><MM/YYYY><_signed>.pdf
const REGEX_RECIBO = /^REC_([A-Za-z0-9.]+)_(\d{2})(\d{4})(_signed)?\.pdf$/i;
const LINHA_CABECALHO_COLABORADORES = 2;
const LINHA_CABECALHO_REGISTROS = 2;
const NOME_ABA_TEMPLATE = "2025_BK";

/**
 * FUNÇÃO PRINCIPAL
*/
function catalogarRecibosDeVencimento() {

  try {
    const colaboradores = carregarColaboradores();
    const pastaOrigem = DriveApp.getFolderById(_config().PASTA_ORIGEM_ID);
    const arquivos = pastaOrigem.getFiles();
    const processados = [];

    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      processados.push(processarArquivo(arquivo, colaboradores));
    }

    enviarRelatorioEmail(processados);
    Logger.log("" + processados.length + " Documentos Processados e Enviado Emails");

  } catch (erro) {
    Logger.log("Erro no processamento principal: " + erro);
    throw erro;
  }

}

/**
 * Gera os meses no formato MM/YYYY dinamicamente
*/
function gerarMesesAno(ano) {
  return Array.from({ length: 12 }, (_, i) => {
    const mes = (i + 1).toString().padStart(2, "0");
    return `${mes}/${ano}`;
  });
}

/**
 * Encontra o número da coluna com base no nome do cabeçalho.
 * Lê os valores exibidos nas células (como texto visível) para evitar problemas com formatação de data.
*/
function encontraColunaNoCabecalho(sheet, columnName, linhaDoCabecalho) {
  const lastColumn = sheet.getLastColumn();
  const headerRowRange = sheet.getRange(linhaDoCabecalho, 1, 1, lastColumn);
  const headerRowValues = headerRowRange.getDisplayValues()[0]; // Use getDisplayValues() to get visible text

  // Normalize the target column name (trim spaces)
  const normalizedColumnName = columnName.trim();

  for (let i = 0; i < headerRowValues.length; i++) {
    const headerValue = headerRowValues[i];

    if (headerValue) {
      const normalizedHeaderValue = headerValue.trim(); // Trim spaces from the display value

      // Compare the normalized header value with the normalized column name
      if (normalizedHeaderValue === normalizedColumnName) {
        return i + 1; // Return 1-based column index
      }
    }
  }

  return -1; // Column not found
}

/**
 * Carrega a lista de colaboradores da planilha.
*/
function carregarColaboradores() {
  try {
    const planilha = SpreadsheetApp.openById(_config().PLANILHA_COLABORADORES_ID);
    const aba = planilha.getSheetByName("Ativos");

    if (!aba) throw new Error("A aba 'Ativos' não foi encontrada.");

    const colSigla = encontraColunaNoCabecalho(aba, "Sigla", LINHA_CABECALHO_COLABORADORES);
    const colNome = encontraColunaNoCabecalho(aba, "Nome", LINHA_CABECALHO_COLABORADORES);

    if (colSigla === -1 || colNome === -1)
      throw new Error("Colunas 'Sigla' ou 'Nome' não encontradas.");

    const colaboradores = new Map();
    const valores = aba.getRange(LINHA_CABECALHO_COLABORADORES + 1, colSigla, aba.getLastRow() - LINHA_CABECALHO_COLABORADORES, 2).getValues();

    valores.forEach(linha => {
      const sigla = linha[0]?.toString().trim().toUpperCase();
      const nome = linha[1]?.toString().trim();
      if (sigla) colaboradores.set(sigla, nome);
    });

    return colaboradores;
  } catch (erro) {
    Logger.log("Erro ao carregar colaboradores: " + erro);
    return new Map();
  }
}

/**
 * Processa um único arquivo REC.
*/
function processarArquivo(arquivo, colaboradores) {

  const nome = arquivo.getName();
  Logger.log("Processando arquivo: " + nome);
  let pastaDestino = null; // Track destination folder
  let motivoErro = "";
  let cor = "#93c47d"; // Default green color

  try {
    if (!nome.includes("_signed")) {
      cor = "#f4cccc"; // Red color for files without "_signed"
      Logger.log(`Arquivo sem "_signed" no nome: ${nome}`);
    }

    if (!REGEX_RECIBO.test(nome)) {
      moverParaNaoIdentificados(arquivo, "Formato inválido");
      pastaDestino = "Não Identificados (Formato inválido)";
      motivoErro = "Formato inválido";
      return criarRegistro(nome, "Erro", "Formato inválido", pastaDestino);
    }

    const dados = extrairDadosArquivo(nome);

    if (!dados) {
      moverParaNaoIdentificados(arquivo, "Padrão inválido");
      pastaDestino = "Não Identificados (Padrão inválido)";
      motivoErro = "Padrão inválido";
      return criarRegistro(nome, "Erro", "Padrão inválido", pastaDestino);
    }

    const infoColaborador = buscarColaborador(dados.sigla, colaboradores);

    if (!infoColaborador) {
      moverParaNaoIdentificados(arquivo, "Sigla não encontrada");
      pastaDestino = "Não Identificados (Sigla não encontrada)";
      motivoErro = "Sigla não encontrada";
      return criarRegistro(nome, "Erro", "Sigla não encontrada", pastaDestino);
    }

    if (verificarDuplicata(dados)) {
      tratarDuplicata(arquivo);
      pastaDestino = "Não Identificados (Arquivo duplicado)";
      motivoErro = "Arquivo duplicado";
      return criarRegistro(nome, "Erro", "Arquivo duplicado", pastaDestino);
    }

    // Pasta destino do mês (criada se não existir), SEM mover o ficheiro ainda.
    const pastaMes = obterPastaDestinoMes(dados);

    // Regista PRIMEIRO na planilha. Só movemos o ficheiro depois de o registo ficar
    // garantido — assim nunca se arquiva um recibo sem ele aparecer na folha. Se o
    // registo falhar, o ficheiro fica em "Por processar" e é re-tentado na próxima execução.
    const resultadoRegisto = atualizarPlanilha(dados, arquivo, infoColaborador, cor, pastaMes.getId());

    if (resultadoRegisto === "preto") {
      // atualizarPlanilha já moveu o ficheiro para Não Identificados.
      return criarRegistro(nome, "Erro", "Célula marcada a preto", "Não Identificados (célula a preto)");
    }

    arquivo.moveTo(pastaMes);
    const mesFormatado = dados.mes.toString().padStart(2, "0");
    pastaDestino = `${dados.ano}/${mesFormatado}`;
    Logger.log(`Arquivo movido para: ${pastaDestino}`);
    return criarRegistro(nome, "Sucesso", "", pastaDestino);

  } catch (erro) {
    Logger.log(`Erro no arquivo ${nome}: ` + erro);
    pastaDestino = "Erro no processamento";
    motivoErro = erro.message;
    return criarRegistro(nome, "Erro", motivoErro, pastaDestino);
  }

}

/**
 * Extrai os dados do nome do arquivo.
*/
function extrairDadosArquivo(nome) {
  const match = nome.match(REGEX_RECIBO);
  if (!match) return null;

  return {
    sigla: match[1].toUpperCase(),
    mes: parseInt(match[2]),
    ano: match[3],
    nomeArquivo: nome
  };
}

function buscarColaborador(sigla, colaboradores) {
  return colaboradores.get(sigla);
}

/**
 * Verifica se o arquivo já existe na pasta de destino.
*/
function verificarDuplicata(dados) {
  try {
    const pastaMes = obterPastaDestinoMes(dados);
    return pastaMes.getFilesByName(dados.nomeArquivo).hasNext();
  } catch (erro) {
    Logger.log("Erro ao verificar duplicata: " + erro);
    return false;
  }
}

/**
 * Obtém (ou cria) a pasta de destino do mês: PASTA_DESTINO/ANO/MM/ANO.
*/
function obterPastaDestinoMes(dados) {
  const pastaAno = obterPasta(_config().PASTA_DESTINO_ID, dados.ano);
  const mesFormatado = dados.mes.toString().padStart(2, "0");
  return obterPasta(pastaAno.getId(), `${mesFormatado}/${dados.ano}`);
}

/**
 * Obtém ou cria uma pasta.
*/
function obterPasta(pastaPaiId, nome) {
  const pastaPai = DriveApp.getFolderById(pastaPaiId);
  const pastas = pastaPai.getFoldersByName(nome);
  if (pastas.hasNext()) {
    return pastas.next();
  } else {
    Logger.log(`Criando pasta: ${nome} na pasta com ID: ${pastaPaiId}`); // Log folder creation
    return pastaPai.createFolder(nome);
  }
}

/**
 * Move o arquivo para a pasta de não identificados.
*/
function moverParaNaoIdentificados(arquivo, motivo) {
  Logger.log(`Movendo arquivo ${arquivo.getName()} para Não Identificados. Motivo: ${motivo}`);
  arquivo.setName(`ERRO_${motivo}_${arquivo.getName()}`)
    .moveTo(DriveApp.getFolderById(_config().PASTA_NAO_IDENTIFICADOS_ID));
}

/**
 * Trata um arquivo duplicado.
*/
function tratarDuplicata(arquivo) {
  Logger.log(`Tratando duplicata: ${arquivo.getName()}`);
  arquivo.setName(`DUPLICADO_${arquivo.getName()}`)
    .moveTo(DriveApp.getFolderById(_config().PASTA_NAO_IDENTIFICADOS_ID));
}

/**
 * Atualiza a planilha com o link do arquivo e nome do colaborador.
 * Devolve "ok" se registou, "preto" se a célula estava marcada a preto (ficheiro mandado p/ Não Identificados).
*/
function atualizarPlanilha(dados, arquivo, nomeColaborador, cor, pastaDestinoId) {

  const cfg = _config();
  const planilha = SpreadsheetApp.openById(cfg.PLANILHA_REGISTROS_ID);
  const ano = dados.ano;
  const aba = obterAbaAnual(planilha, ano);

  const colunaMes = obterColunaMes(aba, dados.mes, ano);
  const linhaSigla = obterLinhaSigla(aba, dados.sigla, nomeColaborador);

  const range = aba.getRange(linhaSigla, colunaMes);
  const backgroundColor = range.getBackground();

  // Célula marcada a preto = "não registar" (regra de negócio): manda para Não Identificados.
  if (backgroundColor.toLowerCase() === '#000000' || backgroundColor.toLowerCase() === 'black') {
    moverParaNaoIdentificados(arquivo, "excel_marcado a preto");
    Logger.log(`Arquivo ${arquivo.getName()} movido para Não Identificados devido a célula marcada a preto.`);
    return "preto";
  }

  // Link para o ficheiro do recibo na célula colaborador x mês.
  range.setFormula(`=HYPERLINK("${arquivo.getUrl()}", "LINK")`)
    .setBackground(cor).setHorizontalAlignment("center");

  // Link para a pasta do mês no cabeçalho (linha 2).
  const mesAnoFormatado = `${dados.mes.toString().padStart(2, "0")}/${ano}`;
  const range2 = aba.getRange(2, colunaMes);
  range2.setFormula(`=HYPERLINK("https://drive.google.com/drive/folders/${pastaDestinoId}", "${mesAnoFormatado}")`)
    .setFontColor(cfg.COR_LINK_PASTA).setHorizontalAlignment("center");

  return "ok";
}

/**
 * Obtém ou cria a aba anual.
*/
function obterAbaAnual(planilha, ano) {
  let aba = planilha.getSheetByName(ano);

  if (!aba) {
    const templateSheet = planilha.getSheetByName(NOME_ABA_TEMPLATE);
    if (!templateSheet) {
      throw new Error(`A aba template "${NOME_ABA_TEMPLATE}" não foi encontrada.`);
    }
    aba = templateSheet.copyTo(planilha).setName(ano);
    Logger.log(`Aba "${ano}" criada a partir do template "${NOME_ABA_TEMPLATE}".`);
  }

  // Garante que os 12 cabeçalhos mensais (MM/ANO) existem e estão no ano correto
  // (o template tem o ano antigo no cabeçalho; aqui repomos para o ano da aba).
  garantirCabecalhosMeses(aba, ano);

  return aba;
}

/**
 * Garante que a linha de cabeçalho tem as 12 colunas de meses no formato MM/ANO.
 * Layout fixo da folha: coluna B = Sigla, C = Nome, D..O = meses 01..12.
*/
function garantirCabecalhosMeses(aba, ano) {
  const meses = gerarMesesAno(ano); // ["01/ANO", ..., "12/ANO"]
  const COLUNA_PRIMEIRO_MES = 4;    // coluna D

  // Lê os cabeçalhos atuais de uma vez (texto visível).
  const atuais = aba.getRange(LINHA_CABECALHO_REGISTROS, COLUNA_PRIMEIRO_MES, 1, meses.length).getDisplayValues()[0];

  // Só escreve nas células vazias ou com o ano errado. As que já mostram o mês certo
  // NÃO são tocadas — assim preserva-se o HYPERLINK para a pasta do mês que o
  // atualizarPlanilha já lá tenha posto.
  meses.forEach((mesAno, i) => {
    if ((atuais[i] || "").trim() !== mesAno) {
      aba.getRange(LINHA_CABECALHO_REGISTROS, COLUNA_PRIMEIRO_MES + i).setValue(mesAno);
      Logger.log(`Cabeçalho do mês reposto na aba "${ano}": ${mesAno}.`);
    }
  });
}

/**
 * Obtém a coluna do mês.
*/
function obterColunaMes(aba, mes, ano) {
  const mesFormatado = `${mes.toString().padStart(2, "0")}/${ano}`;
  let coluna = encontraColunaNoCabecalho(aba, mesFormatado, LINHA_CABECALHO_REGISTROS);

  if (coluna === -1) {
    // Cria a coluna do mês on-the-fly. Layout fixo: B=Sigla, C=Nome, D..O = meses 01..12.
    coluna = 3 + mes;
    aba.getRange(LINHA_CABECALHO_REGISTROS, coluna).setValue(mesFormatado);
    Logger.log(`Coluna "${mesFormatado}" criada on-the-fly na coluna ${coluna} da aba ${ano}.`);
  }

  return coluna;
}

/**
 * Obtém a linha da sigla e atualiza/insere com formatação.
*/
function obterLinhaSigla(aba, sigla, nomeColaborador) {
  const valores = aba.getRange("B:B").getValues();

  for (let i = 0; i < valores.length; i++) {
    if (valores[i][0] === sigla) {
      return i + 1; // Return the existing row
    }
  }

  // Se Sigla nao econtrada
  const novaLinha = aba.getLastRow() + 1;

  aba.getRange(novaLinha, 2).setValue(sigla);
  aba.getRange(novaLinha, 3).setValue(nomeColaborador);

  const celulaSigla = aba.getRange(novaLinha, 2);
  celulaSigla.setBackground("#ffffff");
  celulaSigla.setFontColor("#000000");
  celulaSigla.setBorder(true, true, true, true, true, true);

  return novaLinha;
}

/**
 * Envia o relatório de processamento por email.
*/
function enviarRelatorioEmail(processados) {

  if (processados.length === 0) return;

  const cfg = _config();
  let sucessos = 0;
  let erros = 0;
  let detalhesSucesso = "";
  let detalhesErro = "";

  processados.forEach(p => {
    if (p.status === "Sucesso") {
      sucessos++;
      detalhesSucesso += `  • ${p.arquivo} → ${p.destino}\n`;
    } else {
      erros++;
      detalhesErro += `  • ${p.arquivo} — ${p.motivo} (${p.destino})\n`;
    }
  });

  let corpo = `Foi EXECUTADA a catalogação de recibos na pasta RECIBOS DE VENCIMENTO.\n\n` +
    `Resumo:\n` +
    `  Processados: ${processados.length}\n` +
    `  Catalogados com sucesso: ${sucessos}\n` +
    `  Com erro: ${erros}\n\n`;

  if (sucessos > 0) {
    corpo += `Recibos catalogados:\n${detalhesSucesso}\n`;
  }
  if (erros > 0) {
    corpo += `ATENÇÃO — recibos NÃO catalogados (verificar):\n${detalhesErro}\n`;
  }

  MailApp.sendEmail(cfg.EMAIL_NOTIFICACAO, `Catalogação de Recibos de Vencimento EXECUTADA (${cfg.NOME_EMPRESA})`, corpo);

}

/**
 * Cria um registro de processamento.
*/
function criarRegistro(nome, status, motivo, destino) {
  return { arquivo: nome, status: status, motivo: motivo, destino: destino };
}
