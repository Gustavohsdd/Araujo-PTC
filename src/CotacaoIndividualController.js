// @ts-nocheck

//####################################################################################################
// MÓDULO: COTACAO INDIVIDUAL (SERVER-SIDE CONTROLLER PRINCIPAL) - CotacaoIndividualController.gs
// Funções controller para operações da página de cotação individual,
// excluindo lógicas de módulos específicos como Etapas e Funções do Portal.
//####################################################################################################

/**
 * @file CotacaoIndividualController.gs
 * @description Controlador principal para as operações da página de cotação individual.
 * As lógicas específicas das "Etapas" foram movidas para EtapasController.gs.
 * As lógicas específicas das "Funções do Portal" foram movidas para FuncoesController.gs.
 */

/**
 * Obtém os detalhes de uma cotação específica para exibição na página.
 * @param {string} idCotacao O ID da cotação.
 * @return {object} Um objeto com { success: boolean, dados: Array<object>, cabecalhos?: Array<string>, message?: string }.
 */
function CotacaoIndividualController_obterDetalhesDaCotacao(idCotacao) {
  console.log(`CotacaoIndividualController_obterDetalhesDaCotacao: Solicitado para ID '${idCotacao}'.`);
  try {
    if (!idCotacao) {
      console.warn("CotacaoIndividualController: ID da cotação não fornecido em obterDetalhesDaCotacao.");
      return { success: false, dados: null, message: "ID da Cotação não fornecido." };
    }
    const produtosDaCotacao = CotacaoIndividualCRUD_buscarProdutosPorIdCotacao(idCotacao); 

    if (produtosDaCotacao === null) {
      return { success: false, dados: null, message: `Falha ao buscar produtos para cotação ID ${idCotacao} no CRUD.` };
    }
    return {
      success: true,
      dados: produtosDaCotacao,
      cabecalhos: CABECALHOS_COTACOES, // Constante global
      message: `Dados da cotação ${idCotacao} carregados.`
    };
  } catch (error) {
    console.error(`ERRO em CotacaoIndividualController_obterDetalhesDaCotacao para ID '${idCotacao}': ${error.toString()} Stack: ${error.stack}`);
    return { success: false, dados: null, message: "Erro no controller ao processar detalhes da cotação: " + error.message };
  }
}


/**
 * Controller para salvar a edição de uma célula individual feita na página de Cotação Individual.
 * @param {string} idCotacao O ID da cotação.
 * @param {object} identificadoresLinha Contém Produto, SubProdutoChave, Fornecedor.
 * @param {string} colunaAlterada O nome da coluna alterada.
 * @param {string|number|null} novoValor O novo valor.
 * @return {object} Resultado da operação CRUD.
 */
function CotacaoIndividualController_salvarEdicaoCelulaIndividual(idCotacao, identificadoresLinha, colunaAlterada, novoValor) {
  console.log(`CotacaoIndividualController_salvarEdicaoCelulaIndividual: ID Cotação '${idCotacao}', Identificadores: ${JSON.stringify(identificadoresLinha)}, Coluna: ${colunaAlterada}, Novo Valor: ${novoValor}`);
  try {
    if (!idCotacao) {
      return { success: false, message: "ID da Cotação não fornecido." };
    }
    if (!identificadoresLinha || !identificadoresLinha.Produto || !identificadoresLinha.SubProdutoChave || !identificadoresLinha.Fornecedor) {
      return { success: false, message: "Identificadores da linha incompletos." };
    }
    if (colunaAlterada === undefined || colunaAlterada === null) {
      return { success: false, message: "Coluna alterada não especificada." };
    }

    const resultado = CotacaoIndividualCRUD_salvarEdicaoCelulaCotacao(idCotacao, identificadoresLinha, colunaAlterada, novoValor);
    return resultado;

  } catch (error) {
    console.error(`ERRO em CotacaoIndividualController_salvarEdicaoCelulaIndividual: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no controller ao salvar edição da célula: " + error.message };
  }
}

/**
 * NOVO CONTROLLER: Rota para salvar um conjunto de edições do modal de detalhes.
 * @param {string} idCotacao O ID da cotação.
 * @param {object} identificadoresLinha Contém Produto, SubProdutoChave, Fornecedor.
 * @param {object} alteracoes Objeto com as colunas e novos valores.
 * @return {object} Resultado da operação CRUD.
 */
function CotacaoIndividualController_salvarEdicoesModalDetalhes(idCotacao, identificadoresLinha, alteracoes) {
  try {
    if (!idCotacao || !identificadoresLinha || !alteracoes || Object.keys(alteracoes).length === 0) {
      return { success: false, message: "Dados insuficientes para salvar." };
    }
    // A função CRUD agora faz o trabalho pesado
    return CotacaoIndividualCRUD_salvarEdicoesModalDetalhes(idCotacao, identificadoresLinha, alteracoes);
  } catch (error) {
    console.error(`ERRO em CotacaoIndividualController_salvarEdicoesModalDetalhes: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no controller ao salvar detalhes: " + error.message };
  }
}

/**
 * Controller para acrescentar novos itens a uma cotação existente.
 * @param {string} idCotacao O ID da cotação existente onde os itens serão adicionados.
 * @param {object} opcoesCriacao Objeto contendo o 'tipo' de criação (categoria, fornecedor, etc.) 
 * e as 'selecoes' (array de IDs ou valores selecionados).
 * @return {object} Um objeto com { success: boolean, idCotacao: string|null, numItens: int|null, message: string|null }.
 */
function CotacaoIndividualController_acrescentarItensCotacao(idCotacao, opcoesCriacao) {
  Logger.log(`CotacaoIndividualController_acrescentarItensCotacao: Iniciando para ID '${idCotacao}' com opções: ${JSON.stringify(opcoesCriacao)}`);
  
  if (!idCotacao) {
    return { success: false, message: "ID da cotação existente não foi fornecido." };
  }
  if (!opcoesCriacao || !opcoesCriacao.tipo || !opcoesCriacao.selecoes) {
    return { success: false, message: "Opções para acrescentar itens são inválidas ou incompletas." };
  }

  try {
    // A lógica pesada é delegada para a camada CRUD
    const resultadoCRUD = CotacaoIndividualCRUD_acrescentarItensCotacao(idCotacao, opcoesCriacao);
    
    if (resultadoCRUD && resultadoCRUD.success) {
      Logger.log(`CotacaoIndividualController: Itens acrescentados com sucesso ao ID ${resultadoCRUD.idCotacao}. Itens adicionados: ${resultadoCRUD.numItens}`);
      return {
        success: true,
        idCotacao: resultadoCRUD.idCotacao,
        numItens: resultadoCRUD.numItens,
        message: "Itens acrescentados com sucesso."
      };
    } else {
      Logger.log(`CotacaoIndividualController: Falha ao acrescentar itens no CRUD. Mensagem: ${resultadoCRUD ? resultadoCRUD.message : 'Resultado nulo do CRUD'}`);
      return {
        success: false,
        message: resultadoCRUD ? resultadoCRUD.message : "Erro desconhecido ao acrescentar itens na camada de dados."
      };
    }
  } catch (error) {
    Logger.log(`ERRO CRÍTICO em CotacaoIndividualController_acrescentarItensCotacao: ${error.toString()} Stack: ${error.stack}`);
    return {
      success: false,
      message: "Erro geral no controlador ao acrescentar itens à cotação: " + error.message
    };
  }
}
