// @ts-nocheck

//####################################################################################################
// MÓDULO: ETAPAS (SERVER-SIDE CONTROLLER)
// Funções controller para as etapas da cotação.
//####################################################################################################

/**
 * @file EtapasController.gs
 * @description Controlador do lado do servidor para as funcionalidades do menu "Etapas".
 */

/**
 * Controller para salvar os dados da contagem de estoque.
 * (Originada de CotacaoIndividualController_salvarContagemEstoque)
 * @param {string} idCotacao O ID da cotação.
 * @param {Array<object>} dadosContagem Array com os dados da contagem.
 * @return {object} Resultado da operação.
 */
function EtapasController_salvarContagemEstoque(idCotacao, dadosContagem) {
  console.log(`EtapasController_salvarContagemEstoque: ID '${idCotacao}'. Dados:`, JSON.stringify(dadosContagem).substring(0, 500)); // Log reduzido
  try {
    if (!idCotacao) {
      return { success: false, message: "ID da Cotação não fornecido para salvar contagem." };
    }
    if (!dadosContagem || !Array.isArray(dadosContagem) || dadosContagem.length === 0) {
      return { success: false, message: "Nenhum dado de contagem para salvar." };
    }

    const resultadoSalvar = EtapasCRUD_salvarDadosContagemEstoque(idCotacao, dadosContagem);

    if (resultadoSalvar && resultadoSalvar.success) {
      return { success: true, message: resultadoSalvar.message || "Contagem de estoque salva com sucesso!" };
    } else {
      return { success: false, message: resultadoSalvar ? resultadoSalvar.message : "Falha ao salvar contagem de estoque no CRUD (Etapas)." };
    }
  } catch (error) {
    console.error(`ERRO em EtapasController_salvarContagemEstoque para ID '${idCotacao}': ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no EtapasController ao salvar contagem: " + error.message };
  }
}

/**
 * Controller para atualizar o status de uma cotação.
 * (Originada de CotacaoIndividualController_atualizarStatusCotacao)
 * @param {string} idCotacao O ID da cotação.
 * @param {string} novoStatus O novo status para a cotação.
 * @return {object} Resultado da operação.
 */
function EtapasController_atualizarStatusCotacao(idCotacao, novoStatus) {
  console.log(`EtapasController_atualizarStatusCotacao: ID '${idCotacao}', Novo Status: '${novoStatus}'.`);
  try {
    if (!idCotacao) {
      return { success: false, message: "ID da Cotação não fornecido para atualizar status." };
    }
    if (!novoStatus) {
      return { success: false, message: "Novo status não fornecido." };
    }

    const resultado = EtapasCRUD_atualizarStatusCotacao(idCotacao, novoStatus);
    return resultado;
  } catch (error) {
    console.error(`ERRO em EtapasController_atualizarStatusCotacao para ID '${idCotacao}': ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no EtapasController ao atualizar status: " + error.message };
  }
}

/**
 * Controller para retirar produtos (linhas inteiras de produto principal) de uma cotação.
 * (Originada de CotacaoIndividualController_retirarProdutosCotacao)
 * @param {string} idCotacao O ID da cotação.
 * @param {Array<string>} nomesProdutosPrincipaisParaExcluir Array com os nomes dos produtos principais a serem excluídos.
 * @return {object} Resultado da operação.
 */
function EtapasController_retirarProdutosCotacao(idCotacao, nomesProdutosPrincipaisParaExcluir) {
  console.log(`EtapasController_retirarProdutosCotacao: ID '${idCotacao}'. Produtos a excluir:`, JSON.stringify(nomesProdutosPrincipaisParaExcluir));
  try {
    if (!idCotacao) {
      return { success: false, message: "ID da Cotação não fornecido para retirar produtos." };
    }
    if (!nomesProdutosPrincipaisParaExcluir || !Array.isArray(nomesProdutosPrincipaisParaExcluir)) {
      return { success: false, message: "Lista de produtos para exclusão inválida." };
    }

    const resultado = EtapasCRUD_excluirLinhasDaCotacaoPorProdutoPrincipal(idCotacao, nomesProdutosPrincipaisParaExcluir);
    return resultado;
  } catch (error) {
    console.error(`ERRO em EtapasController_retirarProdutosCotacao para ID '${idCotacao}': ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no EtapasController ao retirar produtos: " + error.message, linhasExcluidas: 0 };
  }
}

/**
 * Controller para a etapa "Enviar para Fornecedores", que gera/atualiza links no portal.
 * (Originada de PortalFornecedorController_gerarLinksParaFornecedores)
 * @param {string} idCotacao O ID da cotação.
 * @return {object} Resultado da operação.
 */
function EtapasController_gerarLinksParaFornecedoresParaEtapaEnvio(idCotacao) {
  Logger.log(`EtapasController_gerarLinksParaFornecedoresParaEtapaEnvio: Iniciando para Cotação ID '${idCotacao}'.`);
  try {
    if (!idCotacao) {
      return { success: false, message: "ID da Cotação não fornecido para gerar links." };
    }
    // A lógica de atualizar status para "Aguardando Preços" já foi feita no client-side antes de chamar esta função.
    // Esta função foca em gerar os links.
    const resultadoGeracao = EtapasCRUD_gerarOuAtualizarLinksPortalParaEtapaEnvio(idCotacao);

    if (resultadoGeracao.success) {
      return {
        success: true,
        message: resultadoGeracao.message || `Links para fornecedores da cotação ${idCotacao} foram processados.`,
        detalhesLinks: resultadoGeracao.detalhesLinks || []
      };
    } else {
      return {
        success: false,
        message: resultadoGeracao.message || `Falha ao processar links para cotação ${idCotacao} (EtapasController).`
      };
    }
  } catch (error) {
    Logger.log(`ERRO em EtapasController_gerarLinksParaFornecedoresParaEtapaEnvio ID '${idCotacao}': ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no EtapasController ao gerar links para fornecedores: " + error.message };
  }
}

/**
 * Controller para retirar subprodutos individuais de uma cotação.
 * @param {string} idCotacao O ID da cotação.
 * @param {Array<object>} subProdutosParaExcluir Array de objetos identificando cada subproduto.
 * @return {object} Resultado da operação.
 */
function EtapasController_retirarSubProdutosCotacao(idCotacao, subProdutosParaExcluir) {
  console.log(`EtapasController_retirarSubProdutosCotacao: ID '${idCotacao}'. Subprodutos a excluir:`, JSON.stringify(subProdutosParaExcluir));
  try {
    if (!idCotacao) {
      return { success: false, message: "ID da Cotação não fornecido." };
    }
    if (!subProdutosParaExcluir || !Array.isArray(subProdutosParaExcluir)) {
      return { success: false, message: "Lista de subprodutos para exclusão inválida." };
    }

    // Chama a função CRUD para fazer o trabalho pesado
    const resultado = EtapasCRUD_excluirLinhasDaCotacaoPorSubProduto(idCotacao, subProdutosParaExcluir);
    return resultado;

  } catch (error) {
    console.error(`ERRO em EtapasController_retirarSubProdutosCotacao para ID '${idCotacao}': ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no Controller ao retirar subprodutos: " + error.message, linhasExcluidas: 0 };
  }
}

/**
 * Controller que busca todos os dados necessários para a Etapa 5: Definir Empresa para Faturamento.
 * @returns {object} Um objeto com a lista de empresas e os pedidos mínimos.
 */
function EtapasController_obterDadosParaEtapaFaturamento() {
  Logger.log("EtapasController_obterDadosParaEtapaFaturamento: Iniciando busca de dados para Etapa 5.");
  try {
    const resultadoEmpresas = EtapasCRUD_obterEmpresasParaFaturamento();
    const resultadoPedidosMinimos = EtapasCRUD_obterPedidosMinimosFornecedores();

    if (!resultadoEmpresas.success || !resultadoPedidosMinimos.success) {
      return {
        success: false,
        message: (resultadoEmpresas.message || "") + " " + (resultadoPedidosMinimos.message || "")
      };
    }

    return {
      success: true,
      empresas: resultadoEmpresas.empresas,
      pedidosMinimos: resultadoPedidosMinimos.pedidosMinimos
    };

  } catch (error) {
    Logger.log(`ERRO em EtapasController_obterDadosParaEtapaFaturamento: ${error.toString()}`);
    return { success: false, message: "Erro geral no controller ao buscar dados da Etapa 5." };
  }
}

/**
 * Controller para buscar os dados necessários para a Etapa 6: Condições de Pagamento.
 * @returns {object} Um objeto com as condições de pagamento de cada fornecedor.
 */
function EtapasController_obterDadosParaEtapaCondicoes() {
  Logger.log("EtapasController_obterDadosParaEtapaCondicoes: Buscando dados para Etapa 6.");
  try {
    const resultadoCondicoes = EtapasCRUD_obterCondicoesPagamentoFornecedores();

    if (!resultadoCondicoes.success) {
      return {
        success: false,
        message: resultadoCondicoes.message
      };
    }

    return {
      success: true,
      condicoes: resultadoCondicoes.condicoes
    };

  } catch (error) {
    Logger.log(`ERRO em EtapasController_obterDadosParaEtapaCondicoes: ${error.toString()}`);
    return { success: false, message: "Erro geral no controller ao buscar dados da Etapa 6." };
  }
}

/**
 * Controller para salvar os dados de condições de pagamento na própria cotação.
 * @param {string} idCotacao O ID da cotação.
 * @param {Array<object>} dadosPagamento Array com os dados de pagamento de cada fornecedor/empresa.
 * @return {object} Resultado da operação.
 */
function EtapasController_salvarCondicoesPagamento(idCotacao, dadosPagamento) {
  Logger.log(`EtapasController_salvarCondicoesPagamento: Salvando ${dadosPagamento.length} condições para cotação ${idCotacao}`);
  try {
    if (!idCotacao) {
      return { success: false, message: "ID da Cotação não fornecido." };
    }
    if (!dadosPagamento || dadosPagamento.length === 0) {
      return { success: false, message: "Nenhum dado de pagamento para salvar." };
    }

    // Chama a nova função CRUD
    const resultado = EtapasCRUD_salvarCondicoesPagamentoNaCotacao(idCotacao, dadosPagamento);
    return resultado;

  } catch (error) {
    Logger.log(`ERRO em EtapasController_salvarCondicoesPagamento: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no Controller ao salvar condições de pagamento: " + error.message };
  }
}

/**
 * Controller que busca e agrupa os dados de pedidos para a página de impressão.
 * @param {string} idCotacao O ID da cotação.
 * @returns {object} Um objeto com o resultado da operação e os dados agrupados.
 */
function EtapasController_obterDadosParaImpressao(idCotacao) {
  Logger.log(`EtapasController_obterDadosParaImpressao: Buscando dados para impressão da cotação ID '${idCotacao}'.`);
  try {
    if (!idCotacao) {
      return { success: false, message: "ID da Cotação não foi fornecido." };
    }
    
    // Chama a função CRUD que faz o trabalho pesado
    const resultado = EtapasCRUD_buscarDadosAgrupadosParaImpressao(idCotacao);
    
    return resultado;

  } catch (error) {
    Logger.log(`ERRO em EtapasController_obterDadosParaImpressao: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro crítico no controller ao buscar dados para impressão: " + error.message };
  }
}

/**
 * Controller que busca os dados dos pedidos e gera uma estrutura com links (simulados) para o envio manual.
 * Substitui a necessidade de uma função em FuncoesController.
 * @param {string} idCotacao O ID da cotação.
 * @returns {object} Um objeto com o resultado da operação e os dados para o modal.
 */
function EtapasController_gerarDadosParaEnvioManual(idCotacao) {
  Logger.log(`EtapasController_gerarDadosParaEnvioManual: Iniciando para Cotação ID '${idCotacao}'.`);
  try {
    // 1. Reutiliza a função de buscar dados para impressão para obter os pedidos agrupados.
    const resultadoBusca = EtapasCRUD_buscarDadosAgrupadosParaImpressao(idCotacao);

    if (!resultadoBusca.success) {
      return { success: false, message: resultadoBusca.message };
    }
    
    if (!resultadoBusca.dados || Object.keys(resultadoBusca.dados).length === 0) {
      return { success: true, dados: [] }; // Sucesso, mas sem dados para processar.
    }

    // 2. Transforma a estrutura de dados para a lista que o modal espera.
    const dadosParaModal = [];
    const dadosAgrupados = resultadoBusca.dados;

    for (const nomeFornecedor in dadosAgrupados) {
      const pedidosDoFornecedor = dadosAgrupados[nomeFornecedor];
      pedidosDoFornecedor.forEach(pedido => {
        // Em um cenário real, aqui ocorreria a geração do PDF e obtenção do link do Google Drive.
        // Para esta implementação, vamos criar um link de placeholder.
        const idSimuladoPdf = Utilities.getUuid();
        const linkPdfSimulado = `https://drive.google.com/file/d/${idSimuladoPdf}/view?usp=sharing`;

        dadosParaModal.push({
          fornecedor: pedido.fornecedor,
          empresaFaturada: pedido.empresaFaturada,
          linkPdf: linkPdfSimulado // Link (simulado) para o PDF
        });
      });
    }

    Logger.log(`EtapasController_gerarDadosParaEnvioManual: ${dadosParaModal.length} links preparados para o modal.`);
    return { success: true, dados: dadosParaModal };

  } catch (error) {
    Logger.log(`ERRO em EtapasController_gerarDadosParaEnvioManual: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro crítico no controller ao gerar dados para envio manual: " + error.message };
  }
}