// @ts-nocheck
/**
 * @file MarcacaoProdutosController.js
 * @description Funções de servidor para a página de marcação de recebimento de produtos.
 */

/**
 * Obtém os subprodutos das cotações que estão com status 'Aguardando Faturamento' ou 'Faturado'
 * e que ainda não foram marcados como 'Recebido', 'Cortado' ou 'Recebido Parcialmente'.
 * Os dados são agrupados por fornecedor.
 *
 * @returns {object} Um objeto com { success: boolean, dados: object|null, message: string|null }.
 * 'dados' é um objeto onde cada chave é um fornecedor e o valor é um array de subprodutos.
 */
function MarcacaoProdutosController_obterProdutosParaMarcacao() {
  Logger.log("MarcacaoProdutosController_obterProdutosParaMarcacao: Iniciando busca de produtos pendentes.");
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaCotacoes = planilha.getSheetByName(ABA_COTACOES);

    if (!abaCotacoes) {
      const msg = `Aba "${ABA_COTACOES}" não encontrada.`;
      Logger.log(`MarcacaoProdutosController: ${msg}`);
      return { success: false, dados: null, message: msg };
    }

    const ultimaLinha = abaCotacoes.getLastRow();
    if (ultimaLinha <= 1) {
      Logger.log("MarcacaoProdutosController: Aba de cotações vazia.");
      return { success: true, dados: {}, message: "Nenhum produto a ser listado." };
    }

    const cabecalhos = Utilities_obterCabecalhos(ABA_COTACOES);
    const indiceStatusCotacao = cabecalhos.indexOf("Status da Cotação");
    const indiceStatusSubproduto = cabecalhos.indexOf("Status do SubProduto");
    const indiceFornecedor = cabecalhos.indexOf("Fornecedor");
    const indiceSubproduto = cabecalhos.indexOf("SubProduto");
    const indiceIdCotacao = cabecalhos.indexOf("ID da Cotação");

    if ([indiceStatusCotacao, indiceStatusSubproduto, indiceFornecedor, indiceSubproduto, indiceIdCotacao].includes(-1)) {
      const msg = "Não foi possível encontrar todas as colunas necessárias na aba Cotações: 'Status da Cotação', 'Status do SubProduto', 'Fornecedor', 'SubProduto', 'ID da Cotação'.";
      Logger.log(`MarcacaoProdutosController: ${msg}`);
      return { success: false, dados: null, message: msg };
    }

    const range = abaCotacoes.getRange(2, 1, ultimaLinha - 1, cabecalhos.length);
    const valores = range.getValues();

    const produtosAgrupados = {};

    valores.forEach((linha, index) => {
      const statusCotacao = linha[indiceStatusCotacao];
      const statusSubproduto = linha[indiceStatusSubproduto];

      const statusValidos = ["Aguardando Faturamento", "Faturado"];
      const statusInvalidosSubproduto = ["Recebido", "Cortado", "Recebido Parcialmente"];

      if (statusValidos.includes(statusCotacao) && !statusInvalidosSubproduto.includes(statusSubproduto)) {
        const fornecedor = linha[indiceFornecedor] || "Fornecedor Não Especificado";
        
        if (!produtosAgrupados[fornecedor]) {
          produtosAgrupados[fornecedor] = [];
        }

        produtosAgrupados[fornecedor].push({
          nomeSubproduto: linha[indiceSubproduto],
          idCotacao: linha[indiceIdCotacao],
          linhaPlanilha: index + 2 // A linha real na planilha (index é baseado em 0 e os dados começam na linha 2)
        });
      }
    });

    Logger.log(`MarcacaoProdutosController: Produtos pendentes encontrados e agrupados para ${Object.keys(produtosAgrupados).length} fornecedores.`);
    return { success: true, dados: produtosAgrupados, message: "Dados obtidos com sucesso." };

  } catch (e) {
    Logger.log(`ERRO em MarcacaoProdutosController_obterProdutosParaMarcacao: ${e.toString()}\n${e.stack}`);
    return { success: false, dados: null, message: `Erro no servidor: ${e.message}` };
  }
}

/**
 * Atualiza o status de um subproduto específico na aba de Cotações para status que não requerem dados adicionais.
 *
 * @param {number} numeroLinha O número da linha na planilha a ser atualizada.
 * @param {string} novoStatus O novo status a ser definido ('Recebido' ou 'Cortado').
 * @returns {object} Um objeto com { success: boolean, message: string }.
 */
function MarcacaoProdutosController_atualizarStatusSubproduto(numeroLinha, novoStatus) {
  Logger.log(`MarcacaoProdutosController_atualizarStatusSubproduto: Tentando atualizar linha ${numeroLinha} para status '${novoStatus}'.`);

  if (!numeroLinha || !novoStatus) {
    return { success: false, message: "Parâmetros inválidos (número da linha ou novo status não fornecido)." };
  }

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaCotacoes = planilha.getSheetByName(ABA_COTACOES);

    if (!abaCotacoes) {
      return { success: false, message: `Aba "${ABA_COTACOES}" não encontrada.` };
    }
    
    const cabecalhos = Utilities_obterCabecalhos(ABA_COTACOES);
    const indiceStatusSubproduto = cabecalhos.indexOf("Status do SubProduto");

    if (indiceStatusSubproduto === -1) {
      return { success: false, message: "Coluna 'Status do SubProduto' não encontrada." };
    }

    const colunaParaAtualizar = indiceStatusSubproduto + 1;
    abaCotacoes.getRange(numeroLinha, colunaParaAtualizar).setValue(novoStatus);
    
    SpreadsheetApp.flush(); 

    Logger.log(`MarcacaoProdutosController: Linha ${numeroLinha} atualizada com sucesso para '${novoStatus}'.`);
    return { success: true, message: "Status atualizado com sucesso." };

  } catch (e) {
    Logger.log(`ERRO em MarcacaoProdutosController_atualizarStatusSubproduto: ${e.toString()}\n${e.stack}`);
    return { success: false, message: `Erro no servidor ao atualizar status: ${e.message}` };
  }
}

/**
 * Atualiza o status de um subproduto para 'Recebido Parcialmente' e registra a quantidade.
 *
 * @param {number} numeroLinha O número da linha na planilha a ser atualizada.
 * @param {number} quantidade A quantidade recebida.
 * @returns {object} Um objeto com { success: boolean, message: string }.
 */
function MarcacaoProdutosController_atualizarStatusParcial(numeroLinha, quantidade) {
  const novoStatus = "Recebido Parcialmente";
  Logger.log(`MarcacaoProdutosController_atualizarStatusParcial: Tentando atualizar linha ${numeroLinha} para status '${novoStatus}' com quantidade ${quantidade}.`);

  if (!numeroLinha || quantidade === undefined || quantidade === null) {
    return { success: false, message: "Parâmetros inválidos (número da linha ou quantidade não fornecida)." };
  }

  try {
    const abaCotacoes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_COTACOES);
    if (!abaCotacoes) {
      return { success: false, message: `Aba "${ABA_COTACOES}" não encontrada.` };
    }
    
    const cabecalhos = Utilities_obterCabecalhos(ABA_COTACOES);
    const indiceStatus = cabecalhos.indexOf("Status do SubProduto");
    const indiceQuantidade = cabecalhos.indexOf("Quantidade Recebida");

    if (indiceStatus === -1 || indiceQuantidade === -1) {
      return { success: false, message: "Não foi possível encontrar as colunas 'Status do SubProduto' ou 'Quantidade Recebida'." };
    }

    // Atualiza as duas colunas
    abaCotacoes.getRange(numeroLinha, indiceStatus + 1).setValue(novoStatus);
    abaCotacoes.getRange(numeroLinha, indiceQuantidade + 1).setValue(quantidade);
    
    SpreadsheetApp.flush();

    Logger.log(`MarcacaoProdutosController: Linha ${numeroLinha} atualizada com sucesso.`);
    return { success: true, message: "Status e quantidade atualizados com sucesso." };

  } catch (e) {
    Logger.log(`ERRO em MarcacaoProdutosController_atualizarStatusParcial: ${e.toString()}\n${e.stack}`);
    return { success: false, message: `Erro no servidor ao atualizar status: ${e.message}` };
  }
}
