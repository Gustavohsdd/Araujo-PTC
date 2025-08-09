// @ts-nocheck

//####################################################################################################
// MÓDULO: ETAPAS (SERVER-SIDE CRUD)
// Funções CRUD para as etapas da cotação.
//####################################################################################################

/**
 * @file EtapasCRUD.gs
 * @description Funções CRUD para as funcionalidades do menu "Etapas".
 */

/**
 * Salva os dados da contagem de estoque nas abas 'Produtos' (Estoque Mínimo) e 'Cotações' (Estoque Atual).
 * A informação de "Comprar" é salva junto com "Estoque Atual" na mesma célula.
 * (Originada de CotacaoIndividualCRUD_salvarDadosContagemEstoque)
 * @param {string} idCotacao O ID da cotação.
 * @param {Array<object>} dadosContagem Array com os dados da contagem.
 * @return {object} Resultado da operação.
 */
function EtapasCRUD_salvarDadosContagemEstoque(idCotacao, dadosContagem) {
  console.log(`EtapasCRUD_salvarDadosContagemEstoque: ID Cotação '${idCotacao}'. Dados:`, JSON.stringify(dadosContagem).substring(0, 500));
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaProdutos = planilha.getSheetByName(ABA_PRODUTOS); // Assumindo ABA_PRODUTOS como constante global
  const abaCotacoes = planilha.getSheetByName(ABA_COTACOES); // Assumindo ABA_COTACOES como constante global
  let atualizacoesProdutos = 0;
  let atualizacoesCotacoes = 0;

  if (!abaProdutos) return { success: false, message: `Aba "${ABA_PRODUTOS}" não encontrada.` };
  if (!abaCotacoes) return { success: false, message: `Aba "${ABA_COTACOES}" não encontrada.` };

  try {
    // Atualizar Estoque Mínimo na aba Produtos
    const cabecalhosProdutos = abaProdutos.getRange(1, 1, 1, abaProdutos.getLastColumn()).getValues()[0];
    const colIndexProdutoEmProdutos = cabecalhosProdutos.indexOf("Produto");
    const colIndexEstMinEmProdutos = cabecalhosProdutos.indexOf("Estoque Minimo");

    if (colIndexProdutoEmProdutos === -1) return { success: false, message: `Coluna "Produto" não encontrada na aba "${ABA_PRODUTOS}".` };
    if (colIndexEstMinEmProdutos === -1) return { success: false, message: `Coluna "Estoque Minimo" não encontrada na aba "${ABA_PRODUTOS}".` };

    const ultimaLinhaProdutos = abaProdutos.getLastRow();
    let dadosPlanilhaProdutos = [];
    if (ultimaLinhaProdutos > 1) {
      dadosPlanilhaProdutos = abaProdutos.getRange(2, 1, ultimaLinhaProdutos - 1, abaProdutos.getLastColumn()).getValues();
    }

    dadosContagem.forEach(contagemItem => {
      if (contagemItem.novoEstoqueMinimoProdutoPrincipal !== null &&
        contagemItem.novoEstoqueMinimoProdutoPrincipal !== undefined &&
        !isNaN(parseFloat(contagemItem.novoEstoqueMinimoProdutoPrincipal))) {
        for (let i = 0; i < dadosPlanilhaProdutos.length; i++) {
          if (String(dadosPlanilhaProdutos[i][colIndexProdutoEmProdutos]).trim() === String(contagemItem.nomeProdutoPrincipal).trim()) {
            abaProdutos.getRange(i + 2, colIndexEstMinEmProdutos + 1).setValue(parseFloat(contagemItem.novoEstoqueMinimoProdutoPrincipal));
            atualizacoesProdutos++;
            break;
          }
        }
      }
    });

    // Atualizar Estoque Atual na aba Cotações
    const cabecalhosCotacoes = abaCotacoes.getRange(1, 1, 1, abaCotacoes.getLastColumn()).getValues()[0];
    const colIndexIdCotacaoEmCotacoes = cabecalhosCotacoes.indexOf("ID da Cotação");
    const colIndexProdutoEmCotacoes = cabecalhosCotacoes.indexOf("Produto");
    const colIndexEstAtualEmCotacoes = cabecalhosCotacoes.indexOf("Estoque Atual");
    // A coluna "Comprar" não é mais necessária para escrita, então o índice é opcional
    // const colIndexComprarEmCotacoes = cabecalhosCotacoes.indexOf("Comprar"); 

    if (colIndexIdCotacaoEmCotacoes === -1) return { success: false, message: `Coluna "ID da Cotação" não encontrada na aba "${ABA_COTACOES}".` };
    if (colIndexProdutoEmCotacoes === -1) return { success: false, message: `Coluna "Produto" não encontrada na aba "${ABA_COTACOES}".` };
    if (colIndexEstAtualEmCotacoes === -1) return { success: false, message: `Coluna "Estoque Atual" não encontrada na aba "${ABA_COTACOES}".` };
    
    const ultimaLinhaCotacoes = abaCotacoes.getLastRow();
    let dadosPlanilhaCotacoes = [];
    if (ultimaLinhaCotacoes > 1) {
      dadosPlanilhaCotacoes = abaCotacoes.getRange(2, 1, ultimaLinhaCotacoes - 1, abaCotacoes.getLastColumn()).getValues();
    }

    for (let i = 0; i < dadosPlanilhaCotacoes.length; i++) {
      const linhaCotacao = dadosPlanilhaCotacoes[i];
      if (String(linhaCotacao[colIndexIdCotacaoEmCotacoes]) === String(idCotacao)) {
        const nomeProdutoNaLinhaCotacao = String(linhaCotacao[colIndexProdutoEmCotacoes]).trim();
        const dadosContagemParaEsteProduto = dadosContagem.find(
          item => String(item.nomeProdutoPrincipal).trim() === nomeProdutoNaLinhaCotacao
        );

        if (dadosContagemParaEsteProduto) {
          const estoqueContadoVal = dadosContagemParaEsteProduto.estoqueAtualContagem;
          const comprarSugestaoVal = dadosContagemParaEsteProduto.comprarSugestao;

          if (estoqueContadoVal !== null || comprarSugestaoVal !== null) {
            const textoCombinado = `Estoque Atual: ${estoqueContadoVal !== null && !isNaN(parseFloat(estoqueContadoVal)) ? parseFloat(estoqueContadoVal) : 'N/A'} / Comprar: ${comprarSugestaoVal !== null && !isNaN(parseFloat(comprarSugestaoVal)) ? parseFloat(comprarSugestaoVal) : 'N/A'}`;
            
            // Define o valor combinado na coluna "Estoque Atual"
            abaCotacoes.getRange(i + 2, colIndexEstAtualEmCotacoes + 1).setValue(textoCombinado);
            
            atualizacoesCotacoes++;
          }
        }
      }
    }

    let message = "Nenhuma alteração realizada na contagem.";
    if (atualizacoesProdutos > 0 || atualizacoesCotacoes > 0) {
      message = `Contagem salva! ${atualizacoesProdutos} produto(s) atualizado(s) em Estoque Mínimo. ${atualizacoesCotacoes} linha(s) de cotação atualizada(s).`;
    }
    return { success: true, message: message };

  } catch (error) {
    console.error(`ERRO em EtapasCRUD_salvarDadosContagemEstoque: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no EtapasCRUD ao salvar dados da contagem: " + error.message };
  }
}

/**
 * Atualiza o status de todas as linhas de uma cotação na aba 'Cotações'.
 * (Originada de CotacaoIndividualCRUD_atualizarStatusCotacao)
 * @param {string} idCotacao O ID da cotação.
 * @param {string} novoStatus O novo status.
 * @return {object} Resultado da operação.
 */
function EtapasCRUD_atualizarStatusCotacao(idCotacao, novoStatus) {
  console.log(`EtapasCRUD_atualizarStatusCotacao: ID '${idCotacao}', Novo Status: '${novoStatus}'.`);
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaCotacoes = planilha.getSheetByName(ABA_COTACOES); // Constante global
  let linhasAtualizadas = 0;

  if (!abaCotacoes) {
    return { success: false, message: `Aba "${ABA_COTACOES}" não encontrada.` };
  }

  try {
    const ultimaLinha = abaCotacoes.getLastRow();
    if (ultimaLinha <= 1) {
      return { success: false, message: `Aba "${ABA_COTACOES}" vazia ou só cabeçalho.` };
    }

    const rangeCompleto = abaCotacoes.getRange(1, 1, ultimaLinha, abaCotacoes.getLastColumn());
    const todosOsValores = rangeCompleto.getValues();
    const cabecalhos = todosOsValores[0];

    const colIndexIdCotacao = cabecalhos.indexOf("ID da Cotação");
    const colIndexStatusCotacao = cabecalhos.indexOf("Status da Cotação");

    if (colIndexIdCotacao === -1) return { success: false, message: `Coluna "ID da Cotação" não encontrada em "${ABA_COTACOES}".` };
    if (colIndexStatusCotacao === -1) return { success: false, message: `Coluna "Status da Cotação" não encontrada em "${ABA_COTACOES}".` };

    for (let i = 1; i < todosOsValores.length; i++) {
      if (String(todosOsValores[i][colIndexIdCotacao]) === String(idCotacao)) {
        abaCotacoes.getRange(i + 1, colIndexStatusCotacao + 1).setValue(novoStatus);
        linhasAtualizadas++;
      }
    }

    if (linhasAtualizadas > 0) {
      return { success: true, message: `Status da cotação ID '${idCotacao}' atualizado para "${novoStatus}".` };
    } else {
      return { success: false, message: `Nenhuma linha encontrada para cotação ID '${idCotacao}'.` };
    }
  } catch (error) {
    console.error(`ERRO em EtapasCRUD_atualizarStatusCotacao: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no EtapasCRUD ao atualizar status: " + error.message };
  }
}

/**
 * Exclui todas as linhas da aba 'Cotações' que correspondem a um ID de cotação e a uma lista de nomes de Produtos Principais.
 * (Originada de CotacaoIndividualCRUD_excluirLinhasDaCotacao)
 * @param {string} idCotacao O ID da cotação.
 * @param {Array<string>} nomesProdutosPrincipaisParaExcluir Array com os nomes dos produtos principais.
 * @return {object} Resultado da operação.
 */
function EtapasCRUD_excluirLinhasDaCotacaoPorProdutoPrincipal(idCotacao, nomesProdutosPrincipaisParaExcluir) {
  console.log(`EtapasCRUD_excluirLinhasDaCotacaoPorProdutoPrincipal: ID '${idCotacao}'. Produtos:`, JSON.stringify(nomesProdutosPrincipaisParaExcluir));
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaCotacoes = planilha.getSheetByName(ABA_COTACOES); // Constante global
  let linhasExcluidasCount = 0;

  if (!abaCotacoes) return { success: false, message: `Aba "${ABA_COTACOES}" não encontrada.`, linhasExcluidas: 0 };
  if (!nomesProdutosPrincipaisParaExcluir || nomesProdutosPrincipaisParaExcluir.length === 0) {
    return { success: false, message: "Nenhum produto principal para exclusão.", linhasExcluidas: 0 };
  }

  try {
    const ultimaLinha = abaCotacoes.getLastRow();
    if (ultimaLinha <= 1) return { success: true, message: "Aba cotações vazia.", linhasExcluidas: 0 };

    const rangeCompleto = abaCotacoes.getRange(1, 1, ultimaLinha, abaCotacoes.getLastColumn());
    const todosOsValores = rangeCompleto.getValues();
    const cabecalhos = todosOsValores[0];

    const colIndexIdCotacao = cabecalhos.indexOf("ID da Cotação");
    const colIndexProdutoPrincipal = cabecalhos.indexOf("Produto");

    if (colIndexIdCotacao === -1) return { success: false, message: `Coluna "ID da Cotação" não encontrada.`, linhasExcluidas: 0 };
    if (colIndexProdutoPrincipal === -1) return { success: false, message: `Coluna "Produto" não encontrada.`, linhasExcluidas: 0 };

    const indicesLinhasParaExcluir = [];
    for (let i = todosOsValores.length - 1; i >= 1; i--) { // Itera de baixo para cima
      const linhaAtual = todosOsValores[i];
      const idCotacaoLinha = String(linhaAtual[colIndexIdCotacao]);
      const produtoPrincipalLinha = String(linhaAtual[colIndexProdutoPrincipal]).trim();

      if (idCotacaoLinha === String(idCotacao) && nomesProdutosPrincipaisParaExcluir.includes(produtoPrincipalLinha)) {
        indicesLinhasParaExcluir.push(i + 1); // Guarda o número da linha na planilha (1-based)
      }
    }

    if (indicesLinhasParaExcluir.length > 0) {
      // A exclusão já é de baixo para cima pela ordem de iteração e coleta.
      // Se não fosse, seria necessário ordenar `indicesLinhasParaExcluir.sort((a, b) => b - a);`
      indicesLinhasParaExcluir.forEach(numLinha => {
        abaCotacoes.deleteRow(numLinha);
        linhasExcluidasCount++;
      });
    }
    return {
      success: true,
      message: `${linhasExcluidasCount} linha(s) de produto(s) excluída(s) com sucesso da cotação.`,
      linhasExcluidas: linhasExcluidasCount
    };
  } catch (error) {
    console.error(`ERRO em EtapasCRUD_excluirLinhasDaCotacaoPorProdutoPrincipal: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no EtapasCRUD ao excluir linhas: " + error.message, linhasExcluidas: 0 };
  }
}


/**
 * Obtém uma lista de nomes de fornecedores únicos para uma determinada cotação, usado na etapa de envio.
 * (Originada de CotacaoIndividualCRUD_obterFornecedoresUnicosDaCotacao)
 * @param {string} idCotacao O ID da cotação.
 * @return {Array<string>|null} Um array de nomes de fornecedores únicos, ou null em caso de erro.
 */
function EtapasCRUD_obterFornecedoresUnicosDaCotacaoParaEtapaEnvio(idCotacao) {
  Logger.log(`EtapasCRUD_obterFornecedoresUnicosDaCotacaoParaEtapaEnvio: Buscando fornecedores para Cotação ID '${idCotacao}'.`);
  const fornecedores = new Set();
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaCotacoes = planilha.getSheetByName(ABA_COTACOES); // Constante global

    if (!abaCotacoes) {
      Logger.log(`EtapasCRUD_obterFornecedoresUnicosDaCotacaoParaEtapaEnvio: Aba "${ABA_COTACOES}" não encontrada.`);
      return null;
    }
    const ultimaLinha = abaCotacoes.getLastRow();
    if (ultimaLinha < 2) return [];

    const dadosCompletos = abaCotacoes.getRange(1, 1, ultimaLinha, abaCotacoes.getLastColumn()).getValues();
    const cabecalhos = dadosCompletos[0];
    // Asume que CABECALHOS_COTACOES é uma constante global acessível
    const idxIdCotacao = cabecalhos.indexOf(CABECALHOS_COTACOES[0]); // "ID da Cotação"
    const idxFornecedor = cabecalhos.indexOf(CABECALHOS_COTACOES[5]); // "Fornecedor"

    if (idxIdCotacao === -1 || idxFornecedor === -1) {
      Logger.log(`EtapasCRUD_obterFornecedoresUnicosDaCotacaoParaEtapaEnvio: Colunas chave não encontradas.`);
      return null;
    }

    for (let i = 1; i < dadosCompletos.length; i++) {
      const linhaAtual = dadosCompletos[i];
      if (String(linhaAtual[idxIdCotacao]).trim() === String(idCotacao).trim()) {
        const nomeFornecedorLinha = String(linhaAtual[idxFornecedor]).trim();
        if (nomeFornecedorLinha) fornecedores.add(nomeFornecedorLinha);
      }
    }
    return Array.from(fornecedores);
  } catch (error) {
    Logger.log(`ERRO CRÍTICO em EtapasCRUD_obterFornecedoresUnicosDaCotacaoParaEtapaEnvio: ${error.toString()} Stack: ${error.stack}`);
    return null;
  }
}


/**
 * Gera ou atualiza links/tokens para fornecedores na ABA_PORTAL para uma dada cotação.
 * (Originada de PortalCRUD_gerarOuAtualizarLinksPortal, adaptada para o contexto de "Etapas")
 * @param {string} idCotacao O ID da cotação.
 * @return {object} { success: boolean, message: string, detalhesLinks?: Array<{fornecedor: string, link: string}> }
 */
function EtapasCRUD_gerarOuAtualizarLinksPortalParaEtapaEnvio(idCotacao) {
  Logger.log(`EtapasCRUD_gerarOuAtualizarLinksPortalParaEtapaEnvio: Iniciando para Cotação ID '${idCotacao}'.`);
  const NOME_ABA_PORTAL = ABA_PORTAL; // Constante global
  const resultado = { success: false, message: "", detalhesLinks: [] };
  const NOME_COLUNA_TEXTO_PERSONALIZADO = "Texto Personalizado Link"; // Deve existir na aba Portal

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaPortal = planilha.getSheetByName(NOME_ABA_PORTAL);

    if (!abaPortal) {
      resultado.message = `Aba "${NOME_ABA_PORTAL}" não encontrada.`;
      return resultado;
    }

    // Usa a função de CRUD de Etapas para obter fornecedores da aba Cotações
    const fornecedoresDaCotacao = EtapasCRUD_obterFornecedoresUnicosDaCotacaoParaEtapaEnvio(idCotacao);
    if (fornecedoresDaCotacao === null) {
      resultado.message = `Falha ao obter fornecedores da cotação ${idCotacao} (EtapasCRUD).`;
      return resultado;
    }
    if (fornecedoresDaCotacao.length === 0) {
      resultado.message = `Nenhum fornecedor encontrado para a cotação ${idCotacao}. Nenhum link a gerar/atualizar.`;
      resultado.success = true;
      return resultado;
    }

    const ultimaLinhaPortal = abaPortal.getLastRow();
    let dadosPortal = [];
    let cabecalhosPortal = [];

    if (ultimaLinhaPortal >= 1) {
      const ultimaColunaPortal = abaPortal.getLastColumn() > 0 ? abaPortal.getLastColumn() : 1;
      dadosPortal = abaPortal.getRange(1, 1, ultimaLinhaPortal, ultimaColunaPortal).getValues();
      cabecalhosPortal = dadosPortal[0].map(c => String(c).trim());
    } else { // Aba portal vazia, cria cabeçalhos
      cabecalhosPortal = CABECALHOS_PORTAL.slice(); // Constante global
      if (!cabecalhosPortal.includes(NOME_COLUNA_TEXTO_PERSONALIZADO)) {
         const idxDataRespostaConst = CABECALHOS_PORTAL.indexOf("Data Resposta");
         if (idxDataRespostaConst !== -1) {
            cabecalhosPortal.splice(idxDataRespostaConst + 1, 0, NOME_COLUNA_TEXTO_PERSONALIZADO);
         } else {
            cabecalhosPortal.push(NOME_COLUNA_TEXTO_PERSONALIZADO);
         }
      }
      abaPortal.appendRow(cabecalhosPortal);
      dadosPortal = [cabecalhosPortal];
    }

    // Assumindo CABECALHOS_PORTAL e STATUS_PORTAL como constantes globais
    const idxIdCotacaoPortal = cabecalhosPortal.indexOf(CABECALHOS_PORTAL[0]);
    const idxFornecedorPortal = cabecalhosPortal.indexOf(CABECALHOS_PORTAL[1]);
    const idxTokenLinkPortal = cabecalhosPortal.indexOf(CABECALHOS_PORTAL[2]);
    const idxLinkAcessoPortal = cabecalhosPortal.indexOf(CABECALHOS_PORTAL[3]);
    const idxStatusRespostaPortal = cabecalhosPortal.indexOf(CABECALHOS_PORTAL[4]);
    const idxDataEnvioPortal = cabecalhosPortal.indexOf(CABECALHOS_PORTAL[5]);
    const idxTextoPersonalizadoPortal = cabecalhosPortal.indexOf(NOME_COLUNA_TEXTO_PERSONALIZADO);

    if ([idxIdCotacaoPortal, idxFornecedorPortal, idxTokenLinkPortal, idxLinkAcessoPortal, idxStatusRespostaPortal, idxDataEnvioPortal].some(idx => idx === -1)) {
      resultado.message = `Coluna(s) essencial(is) não encontrada(s) na aba "${NOME_ABA_PORTAL}". Verifique Constantes.gs.`;
      return resultado;
    }
    if (idxTextoPersonalizadoPortal === -1) {
        Logger.log(`Aviso EtapasCRUD: Coluna "${NOME_COLUNA_TEXTO_PERSONALIZADO}" não encontrada na aba Portal. Não será preenchida.`);
    }

    const scriptUrlBase = ScriptApp.getService().getUrl().replace('/dev', '/exec');
    let linksProcessados = 0;

    for (const nomeFornecedor of fornecedoresDaCotacao) {
      let linhaExistentePortalNum = -1; // 1-based index for sheet, -1 if not found
      for (let i = 1; i < dadosPortal.length; i++) { // Start from 1 to skip header
        if (String(dadosPortal[i][idxIdCotacaoPortal]).trim() === idCotacao &&
          String(dadosPortal[i][idxFornecedorPortal]).trim() === nomeFornecedor) {
          linhaExistentePortalNum = i + 1;
          break;
        }
      }

      let token = "";
      let linkCompleto = "";

      if (linhaExistentePortalNum !== -1) { // Fornecedor já existe no portal para esta cotação
        token = String(dadosPortal[linhaExistentePortalNum - 1][idxTokenLinkPortal]).trim();
        if (!token) {
          token = Utilities.getUuid();
          abaPortal.getRange(linhaExistentePortalNum, idxTokenLinkPortal + 1).setValue(token);
        }
        linkCompleto = String(dadosPortal[linhaExistentePortalNum - 1][idxLinkAcessoPortal] || "").trim();
        if (!linkCompleto.startsWith('http') || !linkCompleto.includes(token)) {
          linkCompleto = `${scriptUrlBase}?view=PortalFornecedorView&token=${token}`; // Adicionado view
          abaPortal.getRange(linhaExistentePortalNum, idxLinkAcessoPortal + 1).setValue(linkCompleto);
        }
        const statusAtual = String(dadosPortal[linhaExistentePortalNum-1][idxStatusRespostaPortal]).trim().toLowerCase();
        if(statusAtual !== STATUS_PORTAL.RESPONDIDO.toLowerCase()){
             abaPortal.getRange(linhaExistentePortalNum, idxStatusRespostaPortal + 1).setValue(STATUS_PORTAL.LINK_GERADO);
        }
        if (!dadosPortal[linhaExistentePortalNum - 1][idxDataEnvioPortal]) {
          abaPortal.getRange(linhaExistentePortalNum, idxDataEnvioPortal + 1).setValue(new Date());
        }
      } else { // Nova entrada no portal
        token = Utilities.getUuid();
        linkCompleto = `${scriptUrlBase}?view=PortalFornecedorView&token=${token}`; // Adicionado view
        const novaLinhaArray = new Array(cabecalhosPortal.length).fill("");
        novaLinhaArray[idxIdCotacaoPortal] = idCotacao;
        novaLinhaArray[idxFornecedorPortal] = nomeFornecedor;
        novaLinhaArray[idxTokenLinkPortal] = token;
        novaLinhaArray[idxLinkAcessoPortal] = linkCompleto;
        novaLinhaArray[idxStatusRespostaPortal] = STATUS_PORTAL.LINK_GERADO;
        if(idxDataEnvioPortal !== -1) novaLinhaArray[idxDataEnvioPortal] = new Date();
        if(idxTextoPersonalizadoPortal !== -1) novaLinhaArray[idxTextoPersonalizadoPortal] = ""; // Default

        abaPortal.appendRow(novaLinhaArray);
        dadosPortal.push(novaLinhaArray); // Adiciona à cópia em memória para consistência no loop
      }
      resultado.detalhesLinks.push({ fornecedor: nomeFornecedor, link: linkCompleto });
      linksProcessados++;
    }

    resultado.success = true;
    resultado.message = `${linksProcessados} link(s) de fornecedor(es) processado(s) para a cotação ${idCotacao} no portal.`;
  } catch (error) {
    resultado.message = `Erro ao gerar/atualizar links (EtapasCRUD): ${error.message}`;
    Logger.log(`ERRO em EtapasCRUD_gerarOuAtualizarLinksPortalParaEtapaEnvio: ${error.toString()} Stack: ${error.stack}`);
  }
  return resultado;
}

/**
 * Exclui linhas específicas da aba 'Cotações' com base em uma lista de identificadores de subproduto.
 * @param {string} idCotacao O ID da cotação.
 * @param {Array<object>} subProdutosParaExcluir Array de objetos {Produto, SubProdutoChave, Fornecedor}.
 * @return {object} Resultado da operação { success, message, linhasExcluidas }.
 */
function EtapasCRUD_excluirLinhasDaCotacaoPorSubProduto(idCotacao, subProdutosParaExcluir) {
  console.log(`EtapasCRUD_excluirLinhasDaCotacaoPorSubProduto: ID '${idCotacao}'. Excluindo ${subProdutosParaExcluir.length} subproduto(s).`);
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaCotacoes = planilha.getSheetByName(ABA_COTACOES);
  let linhasExcluidasCount = 0;

  if (!abaCotacoes) return { success: false, message: `Aba "${ABA_COTACOES}" não encontrada.`, linhasExcluidas: 0 };
  if (subProdutosParaExcluir.length === 0) return { success: true, message: "Nenhum subproduto para exclusão.", linhasExcluidas: 0 };

  try {
    const ultimaLinha = abaCotacoes.getLastRow();
    if (ultimaLinha <= 1) return { success: true, message: "Aba de cotações vazia.", linhasExcluidas: 0 };

    const rangeCompleto = abaCotacoes.getDataRange();
    const todosOsValores = rangeCompleto.getValues();
    const cabecalhos = todosOsValores[0];

    const colIndexIdCotacao = cabecalhos.indexOf("ID da Cotação");
    const colIndexProduto = cabecalhos.indexOf("Produto");
    const colIndexSubProduto = cabecalhos.indexOf("SubProduto");
    const colIndexFornecedor = cabecalhos.indexOf("Fornecedor");

    if ([colIndexIdCotacao, colIndexProduto, colIndexSubProduto, colIndexFornecedor].some(idx => idx === -1)) {
        return { success: false, message: `Colunas chave não encontradas na aba "${ABA_COTACOES}".`, linhasExcluidas: 0 };
    }

    const indicesLinhasParaExcluir = [];
    // Itera de baixo para cima para evitar problemas com a reindexação das linhas após uma exclusão.
    for (let i = todosOsValores.length - 1; i >= 1; i--) {
        const linhaAtual = todosOsValores[i];
        const idCotacaoLinha = String(linhaAtual[colIndexIdCotacao]);

        if (idCotacaoLinha === String(idCotacao)) {
            const produtoLinha = String(linhaAtual[colIndexProduto]).trim();
            const subProdutoLinha = String(linhaAtual[colIndexSubProduto]).trim();
            const fornecedorLinha = String(linhaAtual[colIndexFornecedor]).trim();

            const deveExcluir = subProdutosParaExcluir.some(item =>
                item.Produto.trim() === produtoLinha &&
                item.SubProdutoChave.trim() === subProdutoLinha &&
                item.Fornecedor.trim() === fornecedorLinha
            );

            if (deveExcluir) {
                abaCotacoes.deleteRow(i + 1); // i é 0-based, número da linha é 1-based
                linhasExcluidasCount++;
            }
        }
    }

    return {
      success: true,
      message: `${linhasExcluidasCount} subproduto(s) foram removidos da cotação.`,
      linhasExcluidas: linhasExcluidasCount
    };

  } catch (error) {
    console.error(`ERRO em EtapasCRUD_excluirLinhasDaCotacaoPorSubProduto: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no CRUD ao excluir subprodutos: " + error.message, linhasExcluidas: 0 };
  }
}

/**
 * Busca a lista de empresas cadastradas para serem usadas no faturamento.
 * @returns {{success: boolean, message: string, empresas?: string[]}}
 */
function EtapasCRUD_obterEmpresasParaFaturamento() {
  Logger.log("EtapasCRUD_obterEmpresasParaFaturamento: Buscando empresas na aba Cadastros.");
  try {
    const abaCadastros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CADASTROS);
    if (!abaCadastros) {
      return { success: false, message: `Aba "${ABA_CADASTROS}" não encontrada.` };
    }
    const ultimaLinha = abaCadastros.getLastRow();
    if (ultimaLinha < 2) {
      return { success: true, empresas: [] }; // Nenhuma empresa cadastrada
    }
    const cabecalhos = abaCadastros.getRange(1, 1, 1, abaCadastros.getLastColumn()).getValues()[0];
    const colIndexEmpresa = cabecalhos.indexOf(CABECALHOS_CADASTROS[0]); // "Empresas"

    if (colIndexEmpresa === -1) {
      return { success: false, message: `Coluna "${CABECALHOS_CADASTROS[0]}" não encontrada na aba "${ABA_CADASTROS}".` };
    }

    const empresasData = abaCadastros.getRange(2, colIndexEmpresa + 1, ultimaLinha - 1, 1).getValues();
    const empresas = empresasData.map(row => row[0]).filter(String); // Filtra vazios

    Logger.log(`EtapasCRUD_obterEmpresasParaFaturamento: ${empresas.length} empresas encontradas.`);
    return { success: true, empresas: empresas };

  } catch (error) {
    Logger.log(`ERRO em EtapasCRUD_obterEmpresasParaFaturamento: ${error.toString()}`);
    return { success: false, message: "Erro ao buscar empresas para faturamento." };
  }
}

/**
 * Busca os valores de pedido mínimo para cada fornecedor.
 * USA AS CONSTANTES GLOBAIS PARA ENCONTRAR A ABA E AS COLUNAS CORRETAS.
 * @returns {{success: boolean, message: string, pedidosMinimos?: {[key: string]: number}}}
 */
function EtapasCRUD_obterPedidosMinimosFornecedores() {
  Logger.log("EtapasCRUD_obterPedidosMinimosFornecedores: Buscando pedidos mínimos.");
  
  // Usando as constantes globais fornecidas
  const NOME_ABA_FORNECEDORES = ABA_FORNECEDORES;
  const NOME_COL_FORNECEDOR = CABECALHOS_FORNECEDORES[2];     // "Fornecedor"
  const NOME_COL_PEDIDO_MINIMO = CABECALHOS_FORNECEDORES[12]; // "Pedido Mínimo (R$)"

  try {
    const abaFornecedores = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_FORNECEDORES);
    if (!abaFornecedores) {
      Logger.log(`Aviso: Aba "${NOME_ABA_FORNECEDORES}" não encontrada. Pedido mínimo não será validado.`);
      return { success: true, pedidosMinimos: {} };
    }
    const ultimaLinha = abaFornecedores.getLastRow();
    if (ultimaLinha < 2) {
      return { success: true, pedidosMinimos: {} };
    }

    const cabecalhos = abaFornecedores.getRange(1, 1, 1, abaFornecedores.getLastColumn()).getValues()[0];
    const colFornecedor = cabecalhos.indexOf(NOME_COL_FORNECEDOR);
    const colPedidoMinimo = cabecalhos.indexOf(NOME_COL_PEDIDO_MINIMO);

    if (colFornecedor === -1 || colPedidoMinimo === -1) {
      Logger.log(`Aviso: Colunas "${NOME_COL_FORNECEDOR}" ou "${NOME_COL_PEDIDO_MINIMO}" não encontradas na aba "${NOME_ABA_FORNECEDORES}". Pedido mínimo não será validado.`);
      return { success: true, pedidosMinimos: {} };
    }

    const data = abaFornecedores.getRange(2, 1, ultimaLinha - 1, abaFornecedores.getLastColumn()).getValues();
    const pedidosMinimos = {};

    data.forEach(row => {
      const fornecedor = row[colFornecedor];
      // A célula de valor pode ser string ou número, então tratamos ambos os casos.
      const valorCelula = row[colPedidoMinimo];
      let pedidoMinimo = 0;
      if (typeof valorCelula === 'string') {
        pedidoMinimo = parseFloat(valorCelula.replace("R$", "").replace(/\./g, '').replace(',', '.').trim());
      } else if (typeof valorCelula === 'number') {
        pedidoMinimo = valorCelula;
      }
      
      if (fornecedor && !isNaN(pedidoMinimo) && pedidoMinimo > 0) {
        pedidosMinimos[String(fornecedor).trim()] = pedidoMinimo;
      }
    });
    
    Logger.log(`Pedidos mínimos encontrados: ${JSON.stringify(pedidosMinimos)}`);
    return { success: true, pedidosMinimos: pedidosMinimos };

  } catch (error) {
    Logger.log(`ERRO em EtapasCRUD_obterPedidosMinimosFornecedores: ${error.toString()}`);
    return { success: false, message: "Erro ao buscar pedidos mínimos dos fornecedores." };
  }
}

/**
 * Busca as condições de pagamento para cada fornecedor na aba "Fornecedores".
 * @returns {{success: boolean, message: string, condicoes?: {[key: string]: string}}}
 */
function EtapasCRUD_obterCondicoesPagamentoFornecedores() {
  Logger.log("EtapasCRUD_obterCondicoesPagamentoFornecedores: Buscando condições de pagamento.");

  const NOME_ABA_FORNECEDORES = ABA_FORNECEDORES;
  const NOME_COL_FORNECEDOR = CABECALHOS_FORNECEDORES[2]; // "Fornecedor"
  const NOME_COL_CONDICOES = CABECALHOS_FORNECEDORES[10]; // "Condições de Pagamento"

  try {
    const abaFornecedores = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_FORNECEDORES);
    if (!abaFornecedores) {
      Logger.log(`Aviso: Aba "${NOME_ABA_FORNECEDORES}" não encontrada.`);
      return { success: true, condicoes: {} };
    }
    const ultimaLinha = abaFornecedores.getLastRow();
    if (ultimaLinha < 2) {
      return { success: true, condicoes: {} };
    }

    const cabecalhos = abaFornecedores.getRange(1, 1, 1, abaFornecedores.getLastColumn()).getValues()[0];
    const colFornecedor = cabecalhos.indexOf(NOME_COL_FORNECEDOR);
    const colCondicoes = cabecalhos.indexOf(NOME_COL_CONDICOES);

    if (colFornecedor === -1 || colCondicoes === -1) {
      Logger.log(`Aviso: Colunas "${NOME_COL_FORNECEDOR}" ou "${NOME_COL_CONDICOES}" não encontradas.`);
      return { success: true, condicoes: {} };
    }

    const data = abaFornecedores.getRange(2, 1, ultimaLinha - 1, abaFornecedores.getLastColumn()).getValues();
    const condicoes = {};

    data.forEach(row => {
      const fornecedor = row[colFornecedor];
      const condicao = row[colCondicoes];
      if (fornecedor && condicao) {
        condicoes[String(fornecedor).trim()] = String(condicao);
      }
    });

    Logger.log(`Condições de pagamento encontradas: ${Object.keys(condicoes).length} fornecedores.`);
    return { success: true, condicoes: condicoes };

  } catch (error) {
    Logger.log(`ERRO em EtapasCRUD_obterCondicoesPagamentoFornecedores: ${error.toString()}`);
    return { success: false, message: "Erro ao buscar condições de pagamento." };
  }
}

/**
 * Salva a condição de pagamento escolhida na coluna "Condição de Pagamento" da aba 'Cotações'
 * para cada um dos itens que compõem aquele pedido (agrupamento de fornecedor/empresa).
 * @param {string} idCotacao O ID da cotação.
 * @param {Array<object>} dadosPagamento Array de objetos {fornecedor, empresa, condicao}.
 * @returns {object} Resultado da operação.
 */
function EtapasCRUD_salvarCondicoesPagamentoNaCotacao(idCotacao, dadosPagamento) {
  Logger.log(`EtapasCRUD_salvarCondicoesPagamentoNaCotacao: Iniciando para cotação ${idCotacao}`);
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaCotacoes = planilha.getSheetByName(ABA_COTACOES);
  let linhasAtualizadas = 0;

  if (!abaCotacoes) {
    return { success: false, message: `Aba "${ABA_COTACOES}" não encontrada.` };
  }

  try {
    const range = abaCotacoes.getDataRange();
    const values = range.getValues();
    const headers = values[0];

    const colIdCotacao = headers.indexOf("ID da Cotação");
    const colFornecedor = headers.indexOf("Fornecedor");
    const colEmpresa = headers.indexOf("Empresa Faturada");
    const colCondicaoPagamento = headers.indexOf("Condição de Pagamento");
    const colComprar = headers.indexOf("Comprar");

    if ([colIdCotacao, colFornecedor, colEmpresa, colCondicaoPagamento, colComprar].some(idx => idx === -1)) {
        return { success: false, message: "Uma ou mais colunas necessárias não foram encontradas na aba Cotações." };
    }

    // Criar um mapa para acesso rápido aos dados de pagamento
    const mapaPagamentos = dadosPagamento.reduce((acc, item) => {
        const chave = `${item.fornecedor}__${item.empresa}`;
        acc[chave] = item.condicao;
        return acc;
    }, {});
    
    let algumaAlteracaoFeita = false;

    // Itera pelas linhas de dados (começando em 1 para pular o cabeçalho)
    for (let i = 1; i < values.length; i++) {
        const linha = values[i];

        // Verifica se a linha pertence à cotação e se é um item a ser comprado
        const comprarValor = parseFloat(String(linha[colComprar] || '0').replace(',', '.'));
        if (String(linha[colIdCotacao]) === String(idCotacao) && comprarValor > 0) {
            const fornecedorLinha = String(linha[colFornecedor]);
            const empresaLinha = String(linha[colEmpresa]);
            
            const chaveLinha = `${fornecedorLinha}__${empresaLinha}`;
            
            // Se houver uma condição de pagamento definida para essa combinação de fornecedor/empresa
            if (mapaPagamentos.hasOwnProperty(chaveLinha)) {
                const novaCondicao = mapaPagamentos[chaveLinha];
                // Se a nova condição for diferente da que já está na planilha
                if (String(linha[colCondicaoPagamento]) !== novaCondicao) {
                    values[i][colCondicaoPagamento] = novaCondicao; // Atualiza o valor no array
                    linhasAtualizadas++;
                    algumaAlteracaoFeita = true;
                }
            }
        }
    }

    if (algumaAlteracaoFeita) {
      // Escreve todo o array de valores de volta na planilha de uma vez
      range.setValues(values);
      Logger.log(`${linhasAtualizadas} linhas tiveram a condição de pagamento atualizada.`);
      return { success: true, message: `Condições de pagamento salvas com sucesso para ${linhasAtualizadas} itens.` };
    } else {
      Logger.log(`Nenhuma alteração necessária nas condições de pagamento.`);
      return { success: true, message: "Nenhuma alteração nas condições de pagamento foi necessária." };
    }

  } catch (error) {
    Logger.log(`ERRO em EtapasCRUD_salvarCondicoesPagamentoNaCotacao: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no CRUD ao salvar as condições de pagamento: " + error.message };
  }
}

/**
 * Busca dados na aba 'Cotações' e agrupa os resultados por Fornecedor. Cada fornecedor
 * terá uma lista de pedidos, um para cada Empresa Faturada distinta.
 * (VERSÃO CORRIGIDA - Agrupamento por Fornecedor)
 * @param {string} idCotacao O ID da cotação a ser processada.
 * @returns {{success: boolean, message: string, dados?: object}}
 */
function EtapasCRUD_buscarDadosAgrupadosParaImpressao(idCotacao) {
  Logger.log(`EtapasCRUD_buscarDadosAgrupadosParaImpressao: Iniciando busca para ID '${idCotacao}' com agrupamento por Fornecedor.`);
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaCotacoes = planilha.getSheetByName(ABA_COTACOES);
  const abaCadastros = planilha.getSheetByName(ABA_CADASTROS);
  
  if (!abaCotacoes || !abaCadastros) {
    return { success: false, message: `Uma das abas necessárias (${ABA_COTACOES} ou ${ABA_CADASTROS}) não foi encontrada.` };
  }

  try {
    // 1. Mapear CNPJs das Empresas de Faturamento
    const dadosCadastros = abaCadastros.getDataRange().getValues();
    const cabecalhosCadastros = dadosCadastros.shift();
    const colEmpresaCad = cabecalhosCadastros.indexOf(CABECALHOS_CADASTROS[0]); // "Empresas"
    const colCnpjCad = cabecalhosCadastros.indexOf(CABECALHOS_CADASTROS[1]); // "CNPJ"
    const mapaCnpj = {};
    if (colEmpresaCad > -1 && colCnpjCad > -1) {
      dadosCadastros.forEach(linha => {
        if (linha[colEmpresaCad]) {
          mapaCnpj[linha[colEmpresaCad].toString().trim()] = linha[colCnpjCad] || 'Não informado';
        }
      });
    }

    // 2. Processar a aba de Cotações
    const dadosCotacoes = abaCotacoes.getDataRange().getValues();
    const cabecalhosCotacoes = dadosCotacoes.shift();
    
    const colMap = {
      id: cabecalhosCotacoes.indexOf("ID da Cotação"),
      empresa: cabecalhosCotacoes.indexOf("Empresa Faturada"),
      condicao: cabecalhosCotacoes.indexOf("Condição de Pagamento"),
      fornecedor: cabecalhosCotacoes.indexOf("Fornecedor"),
      subProduto: cabecalhosCotacoes.indexOf("SubProduto"),
      un: cabecalhosCotacoes.indexOf("UN"),
      comprar: cabecalhosCotacoes.indexOf("Comprar"),
      preco: cabecalhosCotacoes.indexOf("Preço"),
      valorTotal: cabecalhosCotacoes.indexOf("Valor Total")
    };
    
    // Validar se todas as colunas foram encontradas
    for (const key in colMap) {
      if (colMap[key] === -1) {
        return { success: false, message: `Coluna obrigatória "${key}" não foi encontrada na aba ${ABA_COTACOES}.` };
      }
    }
    
    const pedidosTemporarios = {};

    // 3. Filtrar e agrupar os dados temporariamente por Fornecedor e Empresa
    dadosCotacoes.forEach(linha => {
      const idLinha = linha[colMap.id];
      const comprarQtd = parseFloat(String(linha[colMap.comprar] || '0').replace(',', '.'));
      
      if (String(idLinha) === String(idCotacao) && comprarQtd > 0) {
        const nomeEmpresa = linha[colMap.empresa];
        if (!nomeEmpresa) return;

        const nomeFornecedor = linha[colMap.fornecedor];
        const chaveUnica = `${nomeFornecedor}__${nomeEmpresa}`; // Chave combinada

        if (!pedidosTemporarios[chaveUnica]) {
          pedidosTemporarios[chaveUnica] = {
            fornecedor: nomeFornecedor,
            empresaFaturada: nomeEmpresa,
            cnpj: mapaCnpj[nomeEmpresa.trim()] || 'Não informado',
            condicaoPagamento: linha[colMap.condicao] || 'Não informada',
            itens: [],
            totalPedido: 0
          };
        }
        
        const item = {
          subProduto: linha[colMap.subProduto],
          un: linha[colMap.un],
          qtd: comprarQtd,
          valorUnit: parseFloat(String(linha[colMap.preco] || '0').replace(',', '.')),
          valorTotal: parseFloat(String(linha[colMap.valorTotal] || '0').replace(',', '.'))
        };
        pedidosTemporarios[chaveUnica].itens.push(item);
        pedidosTemporarios[chaveUnica].totalPedido += item.valorTotal;
      }
    });

    // 4. Estruturar o resultado final agrupado por Fornecedor
    const dadosFinaisAgrupados = {};
    for (const chave in pedidosTemporarios) {
        const pedido = pedidosTemporarios[chave];
        const fornecedor = pedido.fornecedor;

        if (!dadosFinaisAgrupados[fornecedor]) {
            dadosFinaisAgrupados[fornecedor] = [];
        }
        dadosFinaisAgrupados[fornecedor].push(pedido);
    }

    Logger.log(`EtapasCRUD_buscarDadosAgrupadosParaImpressao: Processamento concluído. Encontrados pedidos para ${Object.keys(dadosFinaisAgrupados).length} fornecedor(es).`);
    return { success: true, dados: dadosFinaisAgrupados };

  } catch (error) {
    Logger.log(`ERRO em EtapasCRUD_buscarDadosAgrupadosParaImpressao: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: "Erro no CRUD ao buscar e agrupar dados para impressão: " + error.message };
  }
}

/**
 * NOVA FUNÇÃO: Salva um lote de alterações da coluna "Empresa Faturada" na aba 'Cotações'.
 * @param {string} idCotacao O ID da cotação.
 * @param {Array<object>} alteracoes Array de objetos {Produto, SubProdutoChave, Fornecedor, "Empresa Faturada"}.
 * @returns {object} Resultado da operação { success, message, linhasAtualizadas }.
 */
function EtapasCRUD_salvarEmpresasFaturadasEmLote(idCotacao, alteracoes) {
  Logger.log(`EtapasCRUD_salvarEmpresasFaturadasEmLote: Iniciando salvamento para cotação ${idCotacao}. ${alteracoes.length} alterações.`);
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaCotacoes = planilha.getSheetByName(ABA_COTACOES);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Aguarda até 30s
  
  let linhasAtualizadas = 0;

  try {
    if (!abaCotacoes) {
      throw new Error(`Aba "${ABA_COTACOES}" não encontrada.`);
    }

    const range = abaCotacoes.getDataRange();
    const values = range.getValues();
    const headers = values[0];

    const colMap = {
      id: headers.indexOf("ID da Cotação"),
      produto: headers.indexOf("Produto"),
      subProduto: headers.indexOf("SubProduto"),
      fornecedor: headers.indexOf("Fornecedor"),
      empresaFaturada: headers.indexOf("Empresa Faturada")
    };

    // Valida se todas as colunas necessárias foram encontradas
    for (const key in colMap) {
      if (colMap[key] === -1) {
        throw new Error(`Coluna obrigatória para faturamento ("${key}") não foi encontrada na aba ${ABA_COTACOES}.`);
      }
    }

    // Criar um mapa das alterações para busca rápida
    const mapaAlteracoes = alteracoes.reduce((acc, item) => {
      const chave = `${item.Produto.trim()}__${item.SubProdutoChave.trim()}__${item.Fornecedor.trim()}`;
      acc[chave] = item["Empresa Faturada"];
      return acc;
    }, {});

    // Itera pelas linhas da planilha (a partir da segunda linha de dados)
    for (let i = 1; i < values.length; i++) {
      const linha = values[i];
      if (String(linha[colMap.id]).trim() === String(idCotacao).trim()) {
        const chaveLinha = `${String(linha[colMap.produto]).trim()}__${String(linha[colMap.subProduto]).trim()}__${String(linha[colMap.fornecedor]).trim()}`;
        
        if (mapaAlteracoes.hasOwnProperty(chaveLinha)) {
          const novaEmpresa = mapaAlteracoes[chaveLinha];
          if (String(linha[colMap.empresaFaturada]) !== novaEmpresa) {
            values[i][colMap.empresaFaturada] = novaEmpresa; // Atualiza o valor no array
            linhasAtualizadas++;
          }
        }
      }
    }

    if (linhasAtualizadas > 0) {
      range.setValues(values); // Escreve todo o array de valores de volta na planilha de uma vez
      Logger.log(`${linhasAtualizadas} linha(s) tiveram a "Empresa Faturada" atualizada.`);
      return { 
        success: true, 
        message: `${linhasAtualizadas} item(ns) tiveram a empresa de faturamento atualizada com sucesso.`,
        linhasAtualizadas: linhasAtualizadas
      };
    } else {
      return { 
        success: true, 
        message: "Nenhuma alteração de faturamento foi necessária.",
        linhasAtualizadas: 0 
      };
    }

  } catch (error) {
    Logger.log(`ERRO em EtapasCRUD_salvarEmpresasFaturadasEmLote: ${error.toString()} Stack: ${error.stack}`);
    return { 
      success: false, 
      message: "Erro no CRUD ao salvar faturamento em lote: " + error.message,
      linhasAtualizadas: 0
    };
  } finally {
    lock.releaseLock();
  }
}