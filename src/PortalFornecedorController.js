// @ts-nocheck

/**
 * @file PortalFornecedorController.gs
 * @description Controlador para as operações do Portal do Fornecedor.
 */

/**
 * Busca os dados necessários para exibir no portal do fornecedor.
 * Esta função é chamada pelo App.gs (doGet) quando um token é detectado.
 * @param {string} token O token de acesso do fornecedor.
 * @return {object} Um objeto com os dados para o portal, incluindo a data de abertura formatada.
 */
function PortalController_buscarDadosParaPortalWebService(token) {
  const resultadoDefault = { valido: false, mensagemErro: "Token não encontrado ou inválido.", nomeFornecedor: null, idCotacao: null, produtos: [], pedidoFinalizado: null, dataAberturaFormatada: "" };

  if (!token) {
    return resultadoDefault;
  }

  try {
    const dadosToken = PortalCRUD_validarTokenEObterDadosLog(token);

    if (!dadosToken || !dadosToken.valido) {
      return { ...resultadoDefault, mensagemErro: dadosToken.mensagemErro || "Falha ao validar token." };
    }
    
    const statusPermitidos = [STATUS_PORTAL.LINK_GERADO, STATUS_PORTAL.EM_PREENCHIMENTO, STATUS_PORTAL.RESPONDIDO];
    if (!statusPermitidos.includes(dadosToken.status)) {
        const msgErro = `O acesso para esta cotação (ID: ${dadosToken.idCotacao}) não está mais ativo (Status Atual do Link: ${dadosToken.status}). Contate o comprador.`;
        return { ...resultadoDefault, mensagemErro: msgErro, idCotacao: dadosToken.idCotacao, nomeFornecedor: dadosToken.nomeFornecedor };
    }

    const produtosDoFornecedor = PortalCRUD_buscarProdutosFornecedorDaCotacao(dadosToken.idCotacao, dadosToken.nomeFornecedor);

    if (produtosDoFornecedor === null || !Array.isArray(produtosDoFornecedor)) {
      return { ...resultadoDefault, mensagemErro: "Erro interno ao carregar os produtos da cotação.", idCotacao: dadosToken.idCotacao, nomeFornecedor: dadosToken.nomeFornecedor };
    }

    let pedidoFinalizado = { pedidoExiste: false };
    if (dadosToken.status === STATUS_PORTAL.RESPONDIDO) {
      const dadosPedido = PortalCRUD_buscarDadosDoPedidoFinalizado(dadosToken.idCotacao, dadosToken.nomeFornecedor);
      if (dadosPedido) {
        pedidoFinalizado = dadosPedido;
      }
    }

    const dataAberturaFormatada = PortalController_formatarDataParaBrasileiro(dadosToken.dataEnvio);

    return {
      valido: true,
      nomeFornecedor: dadosToken.nomeFornecedor,
      idCotacao: dadosToken.idCotacao,
      status: dadosToken.status,
      produtos: produtosDoFornecedor,
      pedidoFinalizado: pedidoFinalizado,
      dataAberturaFormatada: dataAberturaFormatada, // NOVO CAMPO PARA A PRÉ-VISUALIZAÇÃO
      mensagemErro: null
    };

  } catch (error) {
    console.error(`PortalController_buscarDadosParaPortalWebService: Erro para token '${token}': ${error.toString()}`);
    return { ...resultadoDefault, mensagemErro: "Erro inesperado no servidor ao processar sua solicitação." };
  }
}


/**
 * Salva a alteração de uma célula individual feita pelo fornecedor no portal.
 * Chamado pelo PortalFornecedorScript.html.
 * @param {object} dadosCelula Objeto contendo {idLinha: number, coluna: string, valor: string|number|null}.
 * @param {string} token O token de acesso do fornecedor.
 * @param {string|number} idCotacao O ID da cotação sendo editada (pode ser string ou número do cliente).
 * @return {object} Um objeto { success: boolean, message: string, novoSubProdutoNomeSeAlterado?: string }.
 */
function PortalController_salvarCelulaFornecedor(dadosCelula, token, idCotacao) {
  const tokenLimpoCliente = String(token || "").trim().replace(/^"|"$/g, '');
  // Logger.log(`PortalController_salvarCelulaFornecedor: Recebido para salvar. Token (cliente): '${tokenLimpoCliente}', Cotação (cliente): '${idCotacao}' (Tipo: ${typeof idCotacao}), Célula: ${JSON.stringify(dadosCelula)}`); // Log removido
  
  const resultadoDefault = { success: false, message: "Não foi possível salvar a alteração." };

  if (!tokenLimpoCliente || idCotacao === undefined || idCotacao === null || !dadosCelula || dadosCelula.idLinha === undefined || !dadosCelula.coluna) {
    // Logger.log("PortalController_salvarCelulaFornecedor: Parâmetros inválidos ou ausentes."); // Log removido
    return { ...resultadoDefault, message: "Dados incompletos para salvar a alteração." };
  }

  try {
    const dadosTokenServidor = PortalCRUD_validarTokenEObterDadosLog(tokenLimpoCliente);
    // Logger.log(`PortalController_salvarCelulaFornecedor: Resultado da validação do token pelo CRUD: ${JSON.stringify(dadosTokenServidor)}`); // Log removido

    if (!dadosTokenServidor || !dadosTokenServidor.valido || String(dadosTokenServidor.idCotacao).trim() !== String(idCotacao).trim()) {
      // Logger.log(`PortalController_salvarCelulaFornecedor: Validação do token/ID da cotação falhou. Token Valido CRUD: ${dadosTokenServidor ? dadosTokenServidor.valido : 'N/A'}, ID Cotação Servidor: ${dadosTokenServidor ? dadosTokenServidor.idCotacao : 'N/A'}, ID Cotação Cliente: ${idCotacao}`); // Log removido
      return { ...resultadoDefault, message: dadosTokenServidor.mensagemErro || "Sessão inválida, token não corresponde à cotação ou token não encontrado." };
    }

    if (dadosTokenServidor.status !== STATUS_PORTAL.LINK_GERADO && dadosTokenServidor.status !== STATUS_PORTAL.RESPONDIDO && dadosTokenServidor.status !== STATUS_PORTAL.EM_PREENCHIMENTO) {
       // Logger.log(`PortalController_salvarCelulaFornecedor: Tentativa de salvar com status de link inválido: ${dadosTokenServidor.status}`); // Log removido
       return { ...resultadoDefault, message: `Não é possível salvar. O status atual desta cotação é '${dadosTokenServidor.status}'.` };
    }

    const resultadoSalvar = PortalCRUD_salvarAlteracaoCelulaIndividual(
      String(idCotacao).trim(), 
      dadosTokenServidor.nomeFornecedor, 
      dadosCelula.idLinha, 
      dadosCelula.coluna, 
      dadosCelula.valor
    );
    
    if (resultadoSalvar.success && dadosTokenServidor.status === STATUS_PORTAL.LINK_GERADO) {
      PortalCRUD_atualizarStatusLinkPortal(tokenLimpoCliente, STATUS_PORTAL.EM_PREENCHIMENTO, null); 
      // Logger.log(`PortalController_salvarCelulaFornecedor: Status do link para token ${tokenLimpoCliente} atualizado para '${STATUS_PORTAL.EM_PREENCHIMENTO}'.`); // Log removido
    }

    return resultadoSalvar;

  } catch (error) {
    // Logger.log(`ERRO CRÍTICO em PortalController_salvarCelulaFornecedor: ${error.toString()} Stack: ${error.stack}`); // Log removido
    console.error(`PortalController_salvarCelulaFornecedor: Erro: ${error.toString()}`);
    return { ...resultadoDefault, message: "Erro inesperado no servidor ao salvar a célula." };
  }
}


/**
 * Marca a cotação como finalizada pelo fornecedor.
 * Chamado pelo PortalFornecedorScript.html.
 * @param {string} token O token de acesso do fornecedor.
 * @param {string|number} idCotacao O ID da cotação.
 * @return {object} Um objeto { success: boolean, message: string }.
 */
function PortalController_finalizarCotacaoFornecedor(token, idCotacao) {
  const tokenLimpoCliente = String(token || "").trim().replace(/^"|"$/g, '');
  // Logger.log(`PortalController_finalizarCotacaoFornecedor: Token (cliente): '${tokenLimpoCliente}', Cotação: '${idCotacao}'`); // Log removido
  const resultadoDefault = { success: false, message: "Não foi possível finalizar a cotação." };

  if (!tokenLimpoCliente || idCotacao === undefined || idCotacao === null) {
    // Logger.log("PortalController_finalizarCotacaoFornecedor: Parâmetros inválidos."); // Log removido
    return { ...resultadoDefault, message: "Dados incompletos para finalizar." };
  }

  try {
    const dadosTokenServidor = PortalCRUD_validarTokenEObterDadosLog(tokenLimpoCliente);
    if (!dadosTokenServidor.valido || String(dadosTokenServidor.idCotacao).trim() !== String(idCotacao).trim()) {
      // Logger.log(`PortalController_finalizarCotacaoFornecedor: Token inválido ou não corresponde. Token: ${tokenLimpoCliente}, ID Cotação Esperado: ${idCotacao}, ID Cotação do Token: ${dadosTokenServidor.idCotacao}`); // Log removido
      return { ...resultadoDefault, message: "Sessão inválida ou token não corresponde." };
    }

    if (dadosTokenServidor.status !== STATUS_PORTAL.LINK_GERADO && dadosTokenServidor.status !== STATUS_PORTAL.RESPONDIDO && dadosTokenServidor.status !== STATUS_PORTAL.EM_PREENCHIMENTO) {
       // Logger.log(`PortalController_finalizarCotacaoFornecedor: Tentativa de finalizar com status de link inválido: ${dadosTokenServidor.status}`); // Log removido
       return { ...resultadoDefault, message: `Não é possível finalizar. O status atual desta cotação é '${dadosTokenServidor.status}'.` };
    }
    
    const resultadoFinalizar = PortalCRUD_atualizarStatusLinkPortal(tokenLimpoCliente, STATUS_PORTAL.RESPONDIDO, new Date());
    return resultadoFinalizar;

  } catch (error) {
    // Logger.log(`ERRO CRÍTICO em PortalController_finalizarCotacaoFornecedor: ${error.toString()} Stack: ${error.stack}`); // Log removido
    console.error(`PortalController_finalizarCotacaoFornecedor: Erro: ${error.toString()}`);
    return { ...resultadoDefault, message: "Erro inesperado no servidor ao finalizar a cotação." };
  }
}

/**
 * Gera ou atualiza links de acesso para todos os fornecedores de uma cotação específica.
 * Chamada pela interface de Cotação Individual (CotacaoIndividualScript.html).
 * @param {string} idCotacao O ID da cotação para a qual gerar os links.
 * @return {object} Um objeto { success: boolean, message: string, detalhesLinks?: Array<object> }.
 */
function PortalFornecedorController_gerarLinksParaFornecedores(idCotacao) {
  // Logger.log(`PortalFornecedorController_gerarLinksParaFornecedores: Iniciando para Cotação ID '${idCotacao}'.`); // Log removido
  const resultadoGeral = { success: false, message: "Falha ao iniciar geração de links.", detalhesLinks: [] };

  if (!idCotacao) {
    resultadoGeral.message = "ID da Cotação não fornecido.";
    // Logger.log(`PortalFornecedorController_gerarLinksParaFornecedores: ${resultadoGeral.message}`); // Log removido
    return resultadoGeral;
  }

  const webAppUrlBase = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL');
  // Logger.log(`PortalFornecedorController_gerarLinksParaFornecedores: Valor lido de WEB_APP_URL das propriedades: '${webAppUrlBase}'`); // Log removido

  if (!webAppUrlBase || !webAppUrlBase.includes("/exec")) {
    resultadoGeral.message = "URL do Web App não está configurada corretamente nas Propriedades do Script. Execute App_configurarUrlWebApp() no arquivo App.gs.";
    // Logger.log(`PortalFornecedorController_gerarLinksParaFornecedores: ${resultadoGeral.message}`); // Log removido
    return resultadoGeral;
  }
  // Logger.log(`PortalFornecedorController_gerarLinksParaFornecedores: Usando URL base do Web App: ${webAppUrlBase}`); // Log removido

  try {
    const fornecedores = CotacaoIndividualCRUD_obterFornecedoresUnicosDaCotacao(idCotacao);
    // Logger.log(`PortalFornecedorController_gerarLinksParaFornecedores: Fornecedores obtidos de CotacaoIndividualCRUD: ${JSON.stringify(fornecedores)}`); // Log removido

    if (fornecedores === null) { 
        resultadoGeral.message = `Erro ao buscar fornecedores para a cotação ID '${idCotacao}'. Verifique os logs do CRUD.`;
        // Logger.log(`PortalFornecedorController_gerarLinksParaFornecedores: ${resultadoGeral.message}`); // Log removido
        return resultadoGeral;
    }
    if (!Array.isArray(fornecedores)) {
        resultadoGeral.message = `Erro: CotacaoIndividualCRUD_obterFornecedoresUnicosDaCotacao não retornou um array para cotação ID '${idCotacao}'.`;
        // Logger.log(`PortalFornecedorController_gerarLinksParaFornecedores: ${resultadoGeral.message} - Tipo retornado: ${typeof fornecedores}`); // Log removido
        return resultadoGeral;
    }
    if (fornecedores.length === 0) {
      resultadoGeral.message = `Nenhum fornecedor encontrado na cotação ID '${idCotacao}' para gerar links.`;
      // Logger.log(`PortalFornecedorController_gerarLinksParaFornecedores: ${resultadoGeral.message}`); // Log removido
      resultadoGeral.success = true; 
      return resultadoGeral;
    }
    // Logger.log(`PortalFornecedorController_gerarLinksParaFornecedores: Fornecedores encontrados: ${fornecedores.join(', ')}`); // Log removido

    let linksGeradosComSucessoCount = 0;
    let linksComErroCount = 0;

    for (const nomeFornecedor of fornecedores) {
      if (!nomeFornecedor || String(nomeFornecedor).trim() === "") {
          // Logger.log(`PortalFornecedorController_gerarLinksParaFornecedores: Nome de fornecedor vazio encontrado para cotação ${idCotacao}, pulando.`); // Log removido
          resultadoGeral.detalhesLinks.push({
            fornecedor: "N/A (Nome Vazio)",
            link: "N/A",
            statusGeracao: "Ignorado",
            mensagem: "Nome do fornecedor estava vazio."
          });
          continue;
      }
      const resultadoLink = PortalCRUD_gerarOuAtualizarLinkFornecedor(idCotacao, nomeFornecedor, webAppUrlBase);
      // Logger.log(`PortalFornecedorController_gerarLinksParaFornecedores: Resultado da geração para ${nomeFornecedor}: ${JSON.stringify(resultadoLink)}`); // Log removido
      
      let statusDescricao = "Falha";
      if (resultadoLink.success) {
          linksGeradosComSucessoCount++;
          statusDescricao = resultadoLink.statusAnterior && resultadoLink.statusAnterior !== STATUS_PORTAL.LINK_GERADO && resultadoLink.statusAnterior !== STATUS_PORTAL.EM_PREENCHIMENTO ? `Atualizado (anterior: ${resultadoLink.statusAnterior})` : "Gerado/Reenviado";
      } else {
          linksComErroCount++;
      }

      resultadoGeral.detalhesLinks.push({
        fornecedor: nomeFornecedor,
        link: resultadoLink.link || "N/A",
        statusGeracao: statusDescricao,
        mensagem: resultadoLink.message
      });
    }

    if (linksGeradosComSucessoCount > 0 && linksComErroCount === 0) {
      resultadoGeral.success = true;
      resultadoGeral.message = `${linksGeradosComSucessoCount} link(s) para fornecedor(es) gerado(s)/atualizado(s) com sucesso para a cotação ${idCotacao}.`;
    } else if (linksGeradosComSucessoCount > 0 && linksComErroCount > 0) {
      resultadoGeral.success = true; 
      resultadoGeral.message = `${linksGeradosComSucessoCount} link(s) gerado(s)/atualizado(s) com sucesso. ${linksComErroCount} falha(s) ao gerar/atualizar links. Verifique os detalhes e a aba '${ABA_PORTAL}'.`;
    } else if (linksComErroCount > 0) {
      resultadoGeral.success = false;
      resultadoGeral.message = `Falha ao gerar/atualizar links para ${linksComErroCount} fornecedor(es). Verifique os detalhes e a aba '${ABA_PORTAL}'.`;
    } else { 
      resultadoGeral.success = true; 
      resultadoGeral.message = "Nenhum link novo precisou ser gerado ou atualizado (verifique status existentes na aba Portal).";
    }
    
    // Logger.log(`PortalFornecedorController_gerarLinksParaFornecedores: Concluído para cotação ${idCotacao}. Mensagem final: ${resultadoGeral.message}`); // Log removido
    return resultadoGeral;

  } catch (error) {
    // Logger.log(`ERRO CRÍTICO em PortalFornecedorController_gerarLinksParaFornecedores para cotação '${idCotacao}': ${error.toString()} Stack: ${error.stack}`); // Log removido
    console.error(`PortalFornecedorController_gerarLinksParaFornecedores: Erro para cotação '${idCotacao}': ${error.toString()}`);
    resultadoGeral.success = false;
    resultadoGeral.message = "Erro inesperado no servidor ao gerar links para fornecedores.";
    return resultadoGeral;
  }
}

/**
 * Formata um objeto Date para o formato "dd/MM/yyyy".
 * @param {Date} data O objeto Date a ser formatado.
 * @return {string} A data formatada ou uma string vazia se a data for inválida.
 */
function PortalController_formatarDataParaBrasileiro(data) {
  if (data instanceof Date && !isNaN(data)) {
    return Utilities.formatDate(data, "GMT-03:00", "dd/MM/yyyy");
  }
  return "";
}