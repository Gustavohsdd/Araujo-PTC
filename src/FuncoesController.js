// @ts-nocheck

//####################################################################################################
// MÓDULO: FUNCOES (SERVER-SIDE CONTROLLER)
// Funções controller para as opções do menu "Funções".
//####################################################################################################

/**
 * @file FuncoesController.gs
 * @description Controlador do lado do servidor para as funcionalidades do menu "Funções".
 */

/**
 * Obtém os dados para a funcionalidade "Gerenciar Cotações (Portal)".
 * (Originada de PortalController_obterDadosGerenciarCotacoes)
 * @return {object} Um objeto com { success: boolean, dados: Array<object>, message?: string }.
 * 'dados' é um array de cotações, cada uma com idCotacao, fornecedores (com nome, link, textoPersonalizadoCotacao, statusResposta), e percentualRespondido.
 */
function FuncoesController_obterDadosGerenciarCotacoes() {
  Logger.log("FuncoesController_obterDadosGerenciarCotacoes: Iniciando busca de dados.");
  try {
    const dadosBrutos = FuncoesCRUD_getDadosGerenciarCotacoes(); // Chama a função no novo FuncoesCRUD
    if (dadosBrutos && dadosBrutos.success) {
      Logger.log(`FuncoesController_obterDadosGerenciarCotacoes: Dados brutos recebidos do CRUD com ${dadosBrutos.dados.length} cotações.`);
      return { success: true, dados: dadosBrutos.dados };
    } else {
      Logger.log(`FuncoesController_obterDadosGerenciarCotacoes: Falha ao obter dados do CRUD. Mensagem: ${dadosBrutos.message}`);
      return { success: false, dados: [], message: dadosBrutos.message || "Falha ao buscar dados das cotações no portal (FuncoesController)." };
    }
  } catch (error) {
    Logger.log(`ERRO em FuncoesController_obterDadosGerenciarCotacoes: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, dados: [], message: "Erro no FuncoesController ao obter dados do portal: " + error.message };
  }
}

/**
 * Controller para excluir uma lista de fornecedores de uma cotação específica (funcionalidade do Portal).
 * (Originada de PortalController_excluirFornecedoresDeCotacaoPortal)
 * @param {string} idCotacao O ID da cotação.
 * @param {Array<string>} nomesFornecedoresArray Array com os nomes dos fornecedores a serem excluídos.
 * @return {object} Resultado da operação.
 */
function FuncoesController_excluirFornecedoresDeCotacaoPortal(idCotacao, nomesFornecedoresArray) {
  Logger.log(`FuncoesController_excluirFornecedoresDeCotacaoPortal: ID Cotação '${idCotacao}', Fornecedores: ${JSON.stringify(nomesFornecedoresArray)}.`);
  try {
    if (!idCotacao || !Array.isArray(nomesFornecedoresArray) || nomesFornecedoresArray.length === 0) {
      return { success: false, message: "ID da Cotação e lista de Nomes dos Fornecedores são obrigatórios." };
    }
    
    let sucessoGeral = true;
    let mensagens = [];
    let fornecedoresExcluidosComSucesso = 0;

    for (const nomeFornecedor of nomesFornecedoresArray) {
        const resultadoExclusao = FuncoesCRUD_excluirFornecedorDaCotacaoPortal(idCotacao, nomeFornecedor); // Chama a função no novo FuncoesCRUD
        if (resultadoExclusao.success) {
            fornecedoresExcluidosComSucesso++;
        } else {
            sucessoGeral = false;
            mensagens.push(`Falha ao excluir ${nomeFornecedor}: ${resultadoExclusao.message}`);
        }
    }

    if (sucessoGeral && fornecedoresExcluidosComSucesso > 0) {
      return { success: true, message: `${fornecedoresExcluidosComSucesso} fornecedor(es) excluído(s) com sucesso da cotação ${idCotacao}.` };
    } else if (fornecedoresExcluidosComSucesso > 0 && !sucessoGeral) {
        return { success: false, message: `Alguns fornecedores foram excluídos, mas ocorreram erros: ${mensagens.join('; ')}` };
    } else if (!sucessoGeral) {
        return { success: false, message: `Falha ao excluir fornecedores: ${mensagens.join('; ')}` };
    } else {
        return { success: false, message: "Nenhum fornecedor foi processado para exclusão."}
    }

  } catch (error) {
    Logger.log(`ERRO em FuncoesController_excluirFornecedoresDeCotacaoPortal: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no FuncoesController ao excluir fornecedores da cotação: " + error.message };
  }
}


/**
 * Controller para salvar o texto personalizado GLOBAL para uma cotação (funcionalidade do Portal).
 * (Originada de PortalController_salvarTextoGlobalCotacaoPortal)
 * @param {string} idCotacao O ID da cotação.
 * @param {string} textoPersonalizado O texto a ser salvo.
 * @return {object} Resultado da operação.
 */
function FuncoesController_salvarTextoGlobalCotacaoPortal(idCotacao, textoPersonalizado) {
  Logger.log(`FuncoesController_salvarTextoGlobalCotacaoPortal: ID Cotação '${idCotacao}'.`);
  try {
    if (!idCotacao) {
      return { success: false, message: "ID da Cotação é obrigatório." };
    }
    if (textoPersonalizado === null || textoPersonalizado === undefined) {
        textoPersonalizado = ""; 
    }

    const resultado = FuncoesCRUD_salvarTextoGlobalCotacaoPortal(idCotacao, textoPersonalizado); // Chama a função no novo FuncoesCRUD
    return resultado;
  } catch (error) {
    Logger.log(`ERRO em FuncoesController_salvarTextoGlobalCotacaoPortal: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no FuncoesController ao salvar texto global da cotação: " + error.message };
  }
}

/**
 * Controller para iniciar o processo de preenchimento dos últimos preços em uma cotação.
 * @param {string} idCotacao O ID da cotação.
 * @return {object} Resultado da operação CRUD.
 */
function FuncoesController_preencherUltimosPrecos(idCotacao) {
  Logger.log(`FuncoesController_preencherUltimosPrecos: Solicitado para ID '${idCotacao}'.`);
  try {
    if (!idCotacao) {
      return { success: false, message: "ID da Cotação não fornecido." };
    }

    // A lógica pesada é delegada para a camada CRUD
    const resultado = FuncoesCRUD_preencherUltimosPrecos(idCotacao);
    return resultado;

  } catch (error) {
    console.error(`ERRO em FuncoesController_preencherUltimosPrecos: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no controller ao processar o preenchimento de preços: " + error.message };
  }
}

// @ts-nocheck

//####################################################################################################
// MÓDULO: FUNCOES (SERVER-SIDE CONTROLLER) - FuncoesController.js
// Funções controller para operações diversas do menu "Funções".
//####################################################################################################

/**
 * @file FuncoesController.gs
 * @description Controlador do lado do servidor para as funcionalidades do menu "Funções".
 */

/**
 * Controller para gerar os PDFs dos pedidos para envio manual.
 * @param {string} idCotacao O ID da cotação.
 * @returns {object} Um objeto com o resultado da operação e os dados dos links.
 */
function FuncoesController_gerarPdfsParaEnvioManual(idCotacao) {
  Logger.log(`FuncoesController_gerarPdfsParaEnvioManual: Solicitado para cotação ID '${idCotacao}'.`);
  try {
    if (!idCotacao) {
      return { success: false, dados: null, message: "ID da Cotação não fornecido." };
    }
    
    // Chama a função CRUD para fazer todo o trabalho pesado
    const resultado = FuncoesCRUD_gerarPdfsParaEnvioManual(idCotacao);
    
    return resultado;

  } catch (error) {
    console.error(`ERRO em FuncoesController_gerarPdfsParaEnvioManual para ID '${idCotacao}': ${error.toString()} Stack: ${error.stack}`);
    return { success: false, dados: null, message: "Erro no controller ao processar a geração de PDFs: " + error.message };
  }
}