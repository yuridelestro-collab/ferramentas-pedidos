/* Ferramentas de Pedidos - Office Add-in para Excel Online */

const MESES = new Set(["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]);

function textoErro(error) {
  if (!error) return "Erro não identificado.";
  return error.message || String(error);
}

function validarAbaMensal(nome) {
  if (!MESES.has(String(nome || "").trim().toUpperCase())) {
    throw new Error("Esta ferramenta deve ser executada em uma aba mensal (JAN, FEV, MAR, ABR, MAI, JUN, JUL, AGO, SET, OUT, NOV ou DEZ). Aba ativa: " + nome);
  }
}

async function imprimirPedido() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    const selecionado = context.workbook.getSelectedRange();
    selecionado.load("rowIndex");
    await context.sync();

    validarAbaMensal(sheet.name);

    const linhaSelecionada = selecionado.rowIndex + 1; // Office.js usa índice iniciado em zero.
    if (linhaSelecionada < 5) {
      throw new Error("Selecione primeiro uma célula dentro do pedido que deseja imprimir (a partir da linha 5).");
    }

    const deslocamento = (linhaSelecionada - 5) % 38;
    if (deslocamento > 36) {
      throw new Error("A célula selecionada está na linha em branco entre pedidos. Selecione uma célula dentro do pedido.");
    }

    const inicio = linhaSelecionada - deslocamento;
    const fim = inicio + 36;
    const endereco = `A${inicio}:M${fim}`;
    sheet.getRange(endereco).select();
    await context.sync();

    return `O intervalo ${endereco} da aba ${sheet.name} foi selecionado. Pressione Ctrl+P para abrir a impressão do Excel. Na tela de impressão, confira se a opção de imprimir a seleção está disponível e selecionada.`;
  });
}

async function novoPedido() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await context.sync();
    validarAbaMensal(sheet.name);

    // Insere 38 linhas inteiras, preservando o espaçamento entre pedidos.
    sheet.getRange("5:42").insert(Excel.InsertShiftDirection.down);
    await context.sync();

    // Após a inserção, o primeiro pedido original passa a ocupar A43:M79.
    // Copiamos tudo: valores, fórmulas, formatos, bordas e validações disponíveis.
    sheet.getRange("A5:M41").copyFrom(
      sheet.getRange("A43:M79"),
      Excel.RangeCopyType.all,
      false,
      false
    );

    // Mantém exatamente os campos que a macro VBA original limpa.
    sheet.getRange("C37").clear(Excel.ClearApplyTo.contents);
    sheet.getRange("B38:B40").clear(Excel.ClearApplyTo.contents);
    sheet.getRange("C7:C10").clear(Excel.ClearApplyTo.contents);
    sheet.getRange("B12:E36").clear(Excel.ClearApplyTo.contents);

    await context.sync();
    return `Novo pedido criado no início da aba ${sheet.name}, com 38 linhas inseridas. Os campos definidos na macro original foram limpos. Não foi criado botão físico dentro da planilha: a impressão será feita pelo botão do suplemento.`;
  });
}

function estaPreenchido(valor) {
  return valor !== null && valor !== undefined && valor !== "";
}

async function importarDados() {
  return Excel.run(async (context) => {
    const origem = context.workbook.worksheets.getActiveWorksheet();
    origem.load("name");
    const bd = context.workbook.worksheets.getItem("BD");

    const origemA = origem.getRange("B50:B74");
    const origemBC = origem.getRange("D50:E74");
    const usadosA = bd.getRange("A:A").getUsedRangeOrNullObject(true);
    const usadosB = bd.getRange("B:B").getUsedRangeOrNullObject(true);

    origemA.load("values");
    origemBC.load("values");
    usadosA.load("isNullObject,rowIndex,rowCount");
    usadosB.load("isNullObject,rowIndex,rowCount");
    await context.sync();

    validarAbaMensal(origem.name);

    // A macro original procura a próxima linha da coluna A sem impor linha mínima.
    // Índices aqui são zero-based; se a coluna estiver vazia, reproduzimos End(xlUp)+1 = linha 2.
    const inicioA = usadosA.isNullObject ? 1 : usadosA.rowIndex + usadosA.rowCount;

    // Para B/C, a macro original impõe linha mínima 5.
    const proximaB = usadosB.isNullObject ? 1 : usadosB.rowIndex + usadosB.rowCount;
    const inicioBC = Math.max(4, proximaB);

    const dadosA = origemA.values
      .map((linha) => linha[0])
      .filter(estaPreenchido)
      .map((valor) => [valor]);

    const dadosBC = origemBC.values
      .filter((linha) => estaPreenchido(linha[0]))
      .map((linha) => [linha[0], linha[1]]);

    const limiteLinhas = 1048576;
    if (inicioA + dadosA.length > limiteLinhas || inicioBC + dadosBC.length > limiteLinhas) {
      throw new Error("A importação ultrapassaria o limite de linhas do Excel.");
    }

    if (dadosA.length > 0) {
      bd.getRangeByIndexes(inicioA, 0, dadosA.length, 1).values = dadosA;
    }

    if (dadosBC.length > 0) {
      bd.getRangeByIndexes(inicioBC, 1, dadosBC.length, 2).values = dadosBC;
      const formulasD = dadosBC.map((_, i) => {
        const linhaExcel = inicioBC + i + 1;
        return [`=A${linhaExcel}*C${linhaExcel}`];
      });
      bd.getRangeByIndexes(inicioBC, 3, dadosBC.length, 1).formulas = formulasD;
    }

    await context.sync();
    return `Importação concluída a partir da aba ${origem.name}. Registros enviados para a coluna A: ${dadosA.length}. Registros enviados para as colunas B e C, com fórmula em D: ${dadosBC.length}. A coluna E não foi modificada.`;
  });
}

function concluirComDialogo(event, titulo, mensagem) {
  const base = (window.ADDIN_BASE_URL || "").replace(/\/$/, "");
  if (!base || base.includes("__ADDIN_BASE_URL__")) {
    console.error("Configure ADDIN_BASE_URL em config.js antes de hospedar o suplemento.");
    event.completed();
    return;
  }

  const url = `${base}/dialog.html?title=${encodeURIComponent(titulo)}&message=${encodeURIComponent(mensagem)}`;
  try {
    Office.context.ui.displayDialogAsync(
      url,
      { height: 34, width: 42, displayInIframe: true },
      (resultado) => {
        if (resultado.status === Office.AsyncResultStatus.Failed) {
          console.error("Não foi possível abrir a mensagem do suplemento:", resultado.error.message);
        }
        event.completed();
      }
    );
  } catch (error) {
    console.error(error);
    event.completed();
  }
}

async function executarComando(event, funcao) {
  try {
    const mensagem = await funcao();
    concluirComDialogo(event, "Operação concluída", mensagem);
  } catch (error) {
    concluirComDialogo(event, "Não foi possível concluir", textoErro(error));
  }
}

function comandoImprimir(event) { executarComando(event, imprimirPedido); }
function comandoNovoPedido(event) { executarComando(event, novoPedido); }
function comandoImportar(event) { executarComando(event, importarDados); }

Office.onReady(() => {
  Office.actions.associate("cmdImprimirPedido", comandoImprimir);
  Office.actions.associate("cmdNovoPedido", comandoNovoPedido);
  Office.actions.associate("cmdImportarDados", comandoImportar);
});
