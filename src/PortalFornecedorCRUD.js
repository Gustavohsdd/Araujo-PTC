// =====================================================================================
// CRUD Functions - PortalFornecedorCRUD.gs
// =====================================================================================

/**
 * @file PortalFornecedorCRUD.gs
 * @description Funções CRUD para as operações do Portal do Fornecedor,
 * interagindo diretamente com as planilhas.
 */

/**
 * Valida um token de acesso na ABA_PORTAL e retorna os dados do log associados.
 * @param {string} token O token de acesso a ser validado.
 * @return {object} Um objeto com:
 * valido: boolean,
 * mensagemErro?: string,
 * idCotacao?: string, 
 * nomeFornecedor?: string,
 * status?: string,
 * linhaNoLog?: number (1-indexed)
 */
/**
 * Valida um token de acesso na ABA_PORTAL e retorna os dados do log associados.
 * @param {string} token O token de acesso a ser validado.
 * @return {object} Um objeto com:
 * valido: boolean,
 * mensagemErro?: string,
 * idCotacao?: string, 
 * nomeFornecedor?: string,
 * status?: string,
 * linhaNoLog?: number (1-indexed),
 * dataEnvio?: Date
 */
function PortalCRUD_validarTokenEObterDadosLog(token) {
  const tokenLimpo = String(token || "").trim().replace(/^"|"$/g, ''); 
  const resultado = { valido: false, mensagemErro: "Token não fornecido.", idCotacao: null, nomeFornecedor: null, status: null, linhaNoLog: -1, dataEnvio: null };
  
  if (!tokenLimpo) return resultado;

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaPortal = planilha.getSheetByName(ABA_PORTAL);

    if (!abaPortal) {
      resultado.mensagemErro = `Aba de controle de portal "${ABA_PORTAL}" não encontrada.`;
      return resultado;
    }

    const ultimaLinha = abaPortal.getLastRow();
    if (ultimaLinha < 1) { 
      resultado.mensagemErro = `Aba "${ABA_PORTAL}" está vazia.`;
      return resultado;
    }
    
    const dadosAbaPortal = abaPortal.getDataRange().getValues();
    const cabecalhosPlanilha = dadosAbaPortal[0]; 

    const indices = {};
    CABECALHOS_PORTAL.forEach(nomeConstante => {
        const indiceReal = cabecalhosPlanilha.indexOf(nomeConstante);
        if (indiceReal !== -1) {
            indices[nomeConstante] = indiceReal;
        }
    });

    const idxToken = indices["Token Acesso"];
    const idxIdCotacao = indices["ID da Cotação"];
    const idxNomeFornecedor = indices["Nome Fornecedor"];
    const idxStatus = indices["Status"];
    const idxDataEnvio = indices["Data Envio"]; // Índice para a nova data

    if (idxToken === undefined || idxIdCotacao === undefined || idxNomeFornecedor === undefined || idxStatus === undefined || idxDataEnvio === undefined) {
      resultado.mensagemErro = `Configuração de colunas da aba "${ABA_PORTAL}" está incorreta ou cabeçalhos não correspondem à constante CABECALHOS_PORTAL. Verifique se a coluna 'Data Envio' existe.`;
      return resultado;
    }

    for (let i = 1; i < dadosAbaPortal.length; i++) { 
      const tokenPlanilha = String(dadosAbaPortal[i][idxToken] || "").trim();
      if (tokenPlanilha === tokenLimpo) {
        resultado.valido = true;
        resultado.idCotacao = dadosAbaPortal[i][idxIdCotacao]; 
        resultado.nomeFornecedor = String(dadosAbaPortal[i][idxNomeFornecedor] || "").trim();
        resultado.status = String(dadosAbaPortal[i][idxStatus] || "").trim();
        resultado.linhaNoLog = i + 1; 
        resultado.dataEnvio = dadosAbaPortal[i][idxDataEnvio] instanceof Date ? dadosAbaPortal[i][idxDataEnvio] : null; // Captura a data
        resultado.mensagemErro = null;
        return resultado;
      }
    }

    resultado.mensagemErro = "Token de acesso inválido ou não encontrado na aba Portal.";
    return resultado;

  } catch (error) {
    console.error(`PortalCRUD_validarTokenEObterDadosLog: Erro para token '${tokenLimpo}': ${error.toString()}`);
    resultado.valido = false;
    resultado.mensagemErro = "Erro interno ao validar o token de acesso.";
    return resultado;
  }
}

/**
 * Busca os produtos de uma cotação específica para um determinado fornecedor na ABA_COTACOES.
 * @param {string|number} idCotacao O ID da cotação.
 * @param {string} nomeFornecedor O nome do fornecedor.
 * @return {Array<object>|null} Um array de objetos de produto ou null em caso de erro.
 */
function PortalCRUD_buscarProdutosFornecedorDaCotacao(idCotacao, nomeFornecedor) {
  // Logger.log(`PortalCRUD_buscarProdutosFornecedorDaCotacao: Iniciando busca para Cotação ID '${idCotacao}' (Tipo: ${typeof idCotacao}), Fornecedor '${nomeFornecedor}'.`); // Log removido
  const produtosDoFornecedor = [];

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaCotacoes = planilha.getSheetByName(ABA_COTACOES);

    if (!abaCotacoes) {
      // Logger.log(`PortalCRUD_buscarProdutosFornecedorDaCotacao: ERRO - Aba "${ABA_COTACOES}" não encontrada.`); // Log removido
      console.error(`PortalCRUD_buscarProdutosFornecedorDaCotacao: Aba "${ABA_COTACOES}" não encontrada.`);
      return null;
    }

    const ultimaLinha = abaCotacoes.getLastRow();
    if (ultimaLinha < 1) { 
      // Logger.log(`PortalCRUD_buscarProdutosFornecedorDaCotacao: Aba "${ABA_COTACOES}" vazia ou sem cabeçalhos.`); // Log removido
      return produtosDoFornecedor; // Retorna array vazio se a aba estiver vazia
    }

    const dadosCompletos = abaCotacoes.getDataRange().getValues();
    const cabecalhosPlanilha = dadosCompletos[0];
    const indicesMapeados = {};

    const colunasParaPortalCliente = {
        idCotacao: "ID da Cotação", 
        fornecedor: "Fornecedor",
        subproduto: "SubProduto",
        tamanho: "Tamanho",
        un: "UN",
        fator: "Fator",
        preco: "Preço"
    };

    let todasRequeridasEncontradas = true;
    for (const chaveCliente in colunasParaPortalCliente) {
        const nomeColunaConstante = colunasParaPortalCliente[chaveCliente];
        const indiceReal = cabecalhosPlanilha.indexOf(nomeColunaConstante);
        if (indiceReal !== -1) {
            indicesMapeados[chaveCliente] = indiceReal;
        } else {
            // Logger.log(`PortalCRUD_buscarProdutosFornecedorDaCotacao: ERRO CRÍTICO - Cabeçalho essencial "${nomeColunaConstante}" (para chave cliente '${chaveCliente}') não encontrado na planilha "${ABA_COTACOES}".`); // Log removido
            console.error(`PortalCRUD_buscarProdutosFornecedorDaCotacao: Cabeçalho essencial "${nomeColunaConstante}" não encontrado na planilha "${ABA_COTACOES}".`);
            todasRequeridasEncontradas = false;
        }
    }
    // Logger.log(`PortalCRUD_buscarProdutosFornecedorDaCotacao: Mapeamento de índices (chaveCliente: indiceReal): ${JSON.stringify(indicesMapeados)}`); // Log removido

    if (!todasRequeridasEncontradas) {
        return null; 
    }
    
    for (let i = 1; i < dadosCompletos.length; i++) { 
      const linhaAtual = dadosCompletos[i];
      const idCotacaoLinha = String(linhaAtual[indicesMapeados.idCotacao]).trim();
      const fornecedorLinha = String(linhaAtual[indicesMapeados.fornecedor]).trim();

      if (idCotacaoLinha === String(idCotacao) && fornecedorLinha === String(nomeFornecedor)) {
        
        let precoValor = linhaAtual[indicesMapeados.preco];
        if (typeof precoValor === 'number') {
             try {
                 precoValor = Utilities.formatString('%.2f', precoValor).replace('.', ',');
             } catch(e){ 
                // Logger.log(`Erro ao formatar preço ${precoValor} da linha ${i+1}: ${e}`); // Log removido
                // Mantém o valor original se a formatação falhar, ou trata como string.
                precoValor = String(precoValor); 
             }
        } else {
            precoValor = precoValor || ''; 
        }
        
        const produtoParaPortal = {
          idLinha: i + 1, 
          subproduto: String(linhaAtual[indicesMapeados.subproduto] || ''),
          tamanho:    String(linhaAtual[indicesMapeados.tamanho] || ''),
          un:         String(linhaAtual[indicesMapeados.un] || ''),
          fator:      String(linhaAtual[indicesMapeados.fator] === null || linhaAtual[indicesMapeados.fator] === undefined ? '' : linhaAtual[indicesMapeados.fator]), 
          preco:      String(precoValor) 
        };
        // Logger.log(`PortalCRUD_buscarProdutosFornecedorDaCotacao: Produto encontrado e processado para linha ${i+1}: ${JSON.stringify(produtoParaPortal)}`); // Log removido
        produtosDoFornecedor.push(produtoParaPortal);
      }
    }
    // Logger.log(`PortalCRUD_buscarProdutosFornecedorDaCotacao: Finalizado. Encontrados ${produtosDoFornecedor.length} produtos para Fornecedor '${nomeFornecedor}' na Cotação ID '${idCotacao}'.`); // Log removido
    return produtosDoFornecedor;

  } catch (error) {
    // Logger.log(`ERRO CRÍTICO em PortalCRUD_buscarProdutosFornecedorDaCotacao para Cotação ID '${idCotacao}', Fornecedor '${nomeFornecedor}': ${error.toString()} Stack: ${error.stack}`); // Log removido
    console.error(`PortalCRUD_buscarProdutosFornecedorDaCotacao: Erro para Cotação ID '${idCotacao}', Fornecedor '${nomeFornecedor}': ${error.toString()}`);
    return null;
  }
}


/**
 * Salva a alteração de uma célula individual na ABA_COTACOES e, se aplicável, na ABA_SUBPRODUTOS.
 * Esta função é chamada pelo PortalFornecedorController quando o fornecedor edita no portal.
 * @param {string} idCotacao O ID da cotação.
 * @param {string} nomeFornecedor O nome do fornecedor (para validação da linha).
 * @param {number} idLinhaNaAbaCotacoes O número da linha física (1-indexed) na ABA_COTACOES.
 * @param {string} colunaAlteradaCliente O nome da coluna como veio do cliente (ex: "tamanho", "preco").
 * @param {string|number|null} novoValor O novo valor para a célula.
 * @return {object} Um objeto { success: boolean, message: string, updatedInCotacoes: boolean, updatedInSubProdutos: boolean, novoSubProdutoNomeSeAlterado?: string }.
 */
function PortalCRUD_salvarAlteracaoCelulaIndividual(idCotacao, nomeFornecedor, idLinhaNaAbaCotacoes, colunaAlteradaCliente, novoValor) {
  // Logger.log(`PortalCRUD_salvarAlteracaoCelulaIndividual (Portal Fornecedor): ID Cotação '${idCotacao}', Fornecedor '${nomeFornecedor}', Linha ${idLinhaNaAbaCotacoes}, Coluna Cliente '${colunaAlteradaCliente}', Novo Valor '${novoValor}' (Tipo: ${typeof novoValor})`); // Log removido
  
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaCotacoes = planilha.getSheetByName(ABA_COTACOES);
  const abaSubProdutos = planilha.getSheetByName(ABA_SUBPRODUTOS); // Pode ser null se não existir

  const COLUNAS_SINCRONIZAVEIS_ABA_SUBPRODUTOS = (typeof COLUNAS_PARA_ABA_SUBPRODUTOS !== 'undefined') ? COLUNAS_PARA_ABA_SUBPRODUTOS : ["SubProduto", "Tamanho", "UN", "Fator"];

  let updatedInCotacoes = false;
  let updatedInSubProdutos = false;
  let nomeSubProdutoOriginalDaLinhaCotacao = null; 
  let nomeProdutoPrincipalDaLinhaCotacao = null;  

  const resultado = { 
    success: false, 
    message: "Nenhuma alteração realizada.", 
    updatedInCotacoes: false, 
    updatedInSubProdutos: false 
  };

  if (!abaCotacoes) {
    resultado.message = `Aba "${ABA_COTACOES}" não encontrada.`;
    // Logger.log(`PortalCRUD_salvarAlteracaoCelulaIndividual: ${resultado.message}`); // Log removido
    console.error(resultado.message);
    return resultado;
  }

  try {
    const cabecalhosPlanilhaCot = abaCotacoes.getRange(1, 1, 1, abaCotacoes.getLastColumn()).getValues()[0];
    const indicesPlanilhaCot = {};
    cabecalhosPlanilhaCot.forEach((c, i) => indicesPlanilhaCot[c] = i);

    const mapaClienteParaCabecalhoReal = {
        "subproduto": "SubProduto",
        "tamanho": "Tamanho",
        "un": "UN",
        "fator": "Fator",
        "preco": "Preço"
    };
    const nomeCorretoColunaAlterada = mapaClienteParaCabecalhoReal[String(colunaAlteradaCliente).toLowerCase()];

    if (!nomeCorretoColunaAlterada) {
      // Logger.log(`PortalCRUD_salvarAlteracaoCelulaIndividual: Não foi possível mapear a coluna cliente '${colunaAlteradaCliente}' para um cabeçalho conhecido. Colunas mapeadas: ${Object.keys(mapaClienteParaCabecalhoReal).join(', ')}`); // Log removido
      throw new Error(`Coluna "${colunaAlteradaCliente}" (enviada pelo cliente) não é uma coluna editável reconhecida ou não pôde ser mapeada.`);
    }
    // Logger.log(`PortalCRUD_salvarAlteracaoCelulaIndividual: Coluna cliente '${colunaAlteradaCliente}' mapeada para cabeçalho constante '${nomeCorretoColunaAlterada}'.`); // Log removido

    const idxColunaAlteradaCot = indicesPlanilhaCot[nomeCorretoColunaAlterada]; 
    
    const nomeConstIdCotacao = CABECALHOS_COTACOES.find(h => h === "ID da Cotação");
    const nomeConstFornecedor = CABECALHOS_COTACOES.find(h => h === "Fornecedor");
    const nomeConstSubProduto = CABECALHOS_COTACOES.find(h => h === "SubProduto");
    const nomeConstProduto = CABECALHOS_COTACOES.find(h => h === "Produto");

    const idxIdCotacaoCot = indicesPlanilhaCot[nomeConstIdCotacao];
    const idxFornecedorCot = indicesPlanilhaCot[nomeConstFornecedor];
    const idxSubProdutoCot = indicesPlanilhaCot[nomeConstSubProduto];
    const idxProdutoCot = indicesPlanilhaCot[nomeConstProduto]; 

    if (idxColunaAlteradaCot === undefined) {
      throw new Error(`Coluna "${nomeCorretoColunaAlterada}" (mapeada de "${colunaAlteradaCliente}") não encontrada no cabeçalho real da aba "${ABA_COTACOES}". Verifique se CABECALHOS_COTACOES está sincronizado com a planilha.`);
    }
    if ([idxIdCotacaoCot, idxFornecedorCot, idxSubProdutoCot, idxProdutoCot].some(idx => idx === undefined)) {
        throw new Error(`Uma ou mais colunas chave ("${nomeConstIdCotacao}", "${nomeConstFornecedor}", "${nomeConstSubProduto}", "${nomeConstProduto}") não encontradas na ABA_COTACOES. Verifique os nomes em CABECALHOS_COTACOES e na planilha.`);
    }
    
    const linhaDadosCot = abaCotacoes.getRange(idLinhaNaAbaCotacoes, 1, 1, abaCotacoes.getLastColumn()).getValues()[0];
    if (String(linhaDadosCot[idxIdCotacaoCot]).trim() !== String(idCotacao).trim() || 
        String(linhaDadosCot[idxFornecedorCot]).trim() !== String(nomeFornecedor).trim()) {
        resultado.message = `Falha de validação de segurança: A linha ${idLinhaNaAbaCotacoes} na aba de cotações não corresponde à cotação/fornecedor esperado.`;
        // Logger.log(`PortalCRUD_salvarAlteracaoCelulaIndividual: ${resultado.message}. Esperado: CotID=${idCotacao}, Forn=${nomeFornecedor}. Encontrado na linha: CotID=${linhaDadosCot[idxIdCotacaoCot]}, Forn=${linhaDadosCot[idxFornecedorCot]}`); // Log removido
        console.error(resultado.message);
        return resultado;
    }
    
    nomeSubProdutoOriginalDaLinhaCotacao = String(linhaDadosCot[idxSubProdutoCot]).trim(); 
    nomeProdutoPrincipalDaLinhaCotacao = String(linhaDadosCot[idxProdutoCot]).trim();

    let valorFinalParaPlanilha = novoValor;
    const nomeColunaPrecoConst = CABECALHOS_COTACOES.find(h => h === "Preço");
    const nomeColunaFatorConst = CABECALHOS_COTACOES.find(h => h === "Fator");

    if (nomeCorretoColunaAlterada === nomeColunaPrecoConst || nomeCorretoColunaAlterada === nomeColunaFatorConst) {
        if (novoValor === null || String(novoValor).trim() === "") {
            valorFinalParaPlanilha = null; 
        } else {
            const num = Number(novoValor); 
            if (!isNaN(num)) {
                valorFinalParaPlanilha = num;
            } else {
                // Logger.log(`PortalCRUD_salvarAlteracaoCelulaIndividual: Valor '${novoValor}' para coluna '${nomeCorretoColunaAlterada}' não é um número válido. Será salvo como texto.`); // Log removido
                // Se não for número, mas a coluna espera um, pode causar problemas.
                // A validação no cliente já deveria pegar isso, mas como fallback, salva como texto.
                valorFinalParaPlanilha = novoValor; 
            }
        }
    }

    abaCotacoes.getRange(idLinhaNaAbaCotacoes, idxColunaAlteradaCot + 1).setValue(valorFinalParaPlanilha);
    updatedInCotacoes = true;
    // Logger.log(`PortalCRUD_salvarAlteracaoCelulaIndividual: ABA_COTACOES - Linha ${idLinhaNaAbaCotacoes}, Coluna "${nomeCorretoColunaAlterada}" atualizada para: ${valorFinalParaPlanilha} (Tipo: ${typeof valorFinalParaPlanilha})`); // Log removido

    if (nomeCorretoColunaAlterada === CABECALHOS_COTACOES.find(h => h === "SubProduto")) { 
        resultado.novoSubProdutoNomeSeAlterado = String(novoValor); 
    }

    if (COLUNAS_SINCRONIZAVEIS_ABA_SUBPRODUTOS.includes(nomeCorretoColunaAlterada)) { 
      if (!abaSubProdutos) {
        // console.warn(`PortalCRUD_salvarAlteracaoCelulaIndividual: Aba "${ABA_SUBPRODUTOS}" não encontrada. Não foi possível atualizar: ${nomeCorretoColunaAlterada}`); // Log removido (era console.warn)
        // A mensagem de resultado já indicará isso.
      } else {
        const ultimaLinhaSub = abaSubProdutos.getLastRow();
        if (ultimaLinhaSub > 0) { 
          const dadosSub = abaSubProdutos.getRange(1, 1, ultimaLinhaSub, abaSubProdutos.getLastColumn()).getValues();
          const cabecalhosSub = dadosSub[0];
          const indicesSub = {};
          cabecalhosSub.forEach((c, i) => indicesSub[c] = i);

          const NOME_COLUNA_PRODUTO_VINCULADO_SUB = "Produto Vinculado"; 
          const NOME_COLUNA_SUBPRODUTO_SUB = "SubProduto";
          const NOME_COLUNA_FORNECEDOR_SUB = "Fornecedor"; 

          const idxProdutoVinculadoSub = indicesSub[NOME_COLUNA_PRODUTO_VINCULADO_SUB];
          const idxSubProdutoSub = indicesSub[NOME_COLUNA_SUBPRODUTO_SUB];
          const idxFornecedorSub = indicesSub[NOME_COLUNA_FORNECEDOR_SUB]; 
          const idxColunaAlteradaSub = indicesSub[nomeCorretoColunaAlterada]; 

          if (idxProdutoVinculadoSub === undefined || idxSubProdutoSub === undefined) {
            // console.warn(`PortalCRUD_salvarAlteracaoCelulaIndividual: Colunas chave ("${NOME_COLUNA_PRODUTO_VINCULADO_SUB}" ou "${NOME_COLUNA_SUBPRODUTO_SUB}") não encontradas na aba "${ABA_SUBPRODUTOS}". Atualização ignorada.`); // Log removido
          } else if (idxColunaAlteradaSub === undefined) {
            // console.warn(`PortalCRUD_salvarAlteracaoCelulaIndividual: Coluna "${nomeCorretoColunaAlterada}" a ser atualizada não encontrada na aba "${ABA_SUBPRODUTOS}". Atualização ignorada.`); // Log removido
          } else {
            for (let i = 1; i < dadosSub.length; i++) {
              const linhaSub = dadosSub[i];
              const produtoVinculadoPlanilha = String(linhaSub[idxProdutoVinculadoSub]).trim();
              const subProdutoPlanilha = String(linhaSub[idxSubProdutoSub]).trim();
              const fornecedorPlanilha = idxFornecedorSub !== undefined ? String(linhaSub[idxFornecedorSub]).trim() : null;

              if (produtoVinculadoPlanilha === nomeProdutoPrincipalDaLinhaCotacao &&
                  subProdutoPlanilha === nomeSubProdutoOriginalDaLinhaCotacao && 
                  (fornecedorPlanilha === null || fornecedorPlanilha === String(nomeFornecedor).trim()) 
              ) {
                abaSubProdutos.getRange(i + 1, idxColunaAlteradaSub + 1).setValue(valorFinalParaPlanilha); 
                updatedInSubProdutos = true;
                // Logger.log(`PortalCRUD_salvarAlteracaoCelulaIndividual: ABA_SUBPRODUTOS - Linha ${i+1}, Coluna "${nomeCorretoColunaAlterada}" atualizada para: ${valorFinalParaPlanilha}`); // Log removido
                break; 
              }
            }
          }
        }
      }
    }

    if (updatedInCotacoes || updatedInSubProdutos) {
      resultado.success = true;
      let msgPartCotacoes = updatedInCotacoes ? "Cotações: Sim" : "Cotações: Não";
      let msgPartSubProdutos = updatedInSubProdutos ? "SubProdutos: Sim" : "SubProdutos: Não";
      if (COLUNAS_SINCRONIZAVEIS_ABA_SUBPRODUTOS.includes(nomeCorretoColunaAlterada) && !abaSubProdutos) {
          msgPartSubProdutos = "SubProdutos: Aba não encontrada";
      } else if (COLUNAS_SINCRONIZAVEIS_ABA_SUBPRODUTOS.includes(nomeCorretoColunaAlterada) && !updatedInSubProdutos && abaSubProdutos) {
          msgPartSubProdutos = `SubProdutos: Linha não encontrada (ProdPrinc: ${nomeProdutoPrincipalDaLinhaCotacao}, SubProd Orig: ${nomeSubProdutoOriginalDaLinhaCotacao}, Forn: ${nomeFornecedor})`;
      }
      resultado.message = `Célula "${nomeCorretoColunaAlterada}" atualizada. ${msgPartCotacoes}, ${msgPartSubProdutos}.`;
    } else if (updatedInCotacoes && !COLUNAS_SINCRONIZAVEIS_ABA_SUBPRODUTOS.includes(nomeCorretoColunaAlterada)) {
      resultado.success = true;
      resultado.message = `Célula "${nomeCorretoColunaAlterada}" atualizada na cotação.`;
    }
    
    resultado.updatedInCotacoes = updatedInCotacoes;
    resultado.updatedInSubProdutos = updatedInSubProdutos;
    return resultado;

  } catch (error) {
    // Logger.log(`ERRO CRÍTICO em PortalCRUD_salvarAlteracaoCelulaIndividual: ${error.toString()} Stack: ${error.stack}`); // Log removido
    console.error(`PortalCRUD_salvarAlteracaoCelulaIndividual: Erro: ${error.toString()}`);
    resultado.success = false;
    resultado.message = `Erro ao salvar alteração da célula: ${error.message}`;
    return resultado;
  }
}


/**
 * Atualiza o status e a data de resposta de um link na ABA_PORTAL.
 * @param {string} token O token de acesso do link.
 * @param {string} novoStatus O novo status a ser definido (ex: STATUS_PORTAL.RESPONDIDO).
 * @param {Date|null} dataResposta A data da resposta (null para limpar).
 * @return {object} Um objeto { success: boolean, message: string }.
 */
function PortalCRUD_atualizarStatusLinkPortal(token, novoStatus, dataResposta) {
  const tokenLimpo = String(token || "").trim().replace(/^"|"$/g, ''); 
  // Logger.log(`PortalCRUD_atualizarStatusLinkPortal: Token (limpo) '${tokenLimpo}', Novo Status '${novoStatus}', Data Resposta '${dataResposta}'`); // Log removido
  
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaPortal = planilha.getSheetByName(ABA_PORTAL);

    if (!abaPortal) {
      // Logger.log(`PortalCRUD_atualizarStatusLinkPortal: Aba "${ABA_PORTAL}" não encontrada.`); // Log removido
      console.error(`PortalCRUD_atualizarStatusLinkPortal: Aba "${ABA_PORTAL}" não encontrada.`);
      return { success: false, message: `Aba de controle de portal "${ABA_PORTAL}" não encontrada.` };
    }

    const dadosAbaPortal = abaPortal.getDataRange().getValues();
    const cabecalhos = dadosAbaPortal[0];
    const indices = {};
    CABECALHOS_PORTAL.forEach(nomeConstante => {
        const indiceReal = cabecalhos.indexOf(nomeConstante);
        if (indiceReal !== -1) indices[nomeConstante] = indiceReal;
    });

    const idxToken = indices["Token Acesso"]; 
    const idxStatus = indices["Status"]; 
    const idxDataResposta = indices["Data Resposta"]; 

    if (idxToken === undefined || idxStatus === undefined || idxDataResposta === undefined) {
      // Logger.log(`PortalCRUD_atualizarStatusLinkPortal: Configuração de colunas da aba "${ABA_PORTAL}" está incorreta.`); // Log removido
      console.error(`PortalCRUD_atualizarStatusLinkPortal: Configuração de colunas da aba "${ABA_PORTAL}" está incorreta.`);
      return { success: false, message: "Configuração de colunas da aba Portal incorreta." };
    }

    for (let i = 1; i < dadosAbaPortal.length; i++) {
      const tokenPlanilha = String(dadosAbaPortal[i][idxToken] || "").trim();
      if (tokenPlanilha === tokenLimpo) {
        abaPortal.getRange(i + 1, idxStatus + 1).setValue(novoStatus);
        if (dataResposta instanceof Date) {
          abaPortal.getRange(i + 1, idxDataResposta + 1).setValue(dataResposta);
        } else {
          abaPortal.getRange(i + 1, idxDataResposta + 1).clearContent();
        }
        // Logger.log(`PortalCRUD_atualizarStatusLinkPortal: Status do token '${tokenLimpo}' atualizado para '${novoStatus}' na linha ${i+1}.`); // Log removido
        return { success: true, message: "Status do link atualizado com sucesso." };
      }
    }
    // Logger.log(`PortalCRUD_atualizarStatusLinkPortal: Token '${tokenLimpo}' não encontrado para atualização de status.`); // Log removido
    return { success: false, message: "Token não encontrado para atualização de status." };

  } catch (error) {
    // Logger.log(`ERRO CRÍTICO em PortalCRUD_atualizarStatusLinkPortal: ${error.toString()} Stack: ${error.stack}`); // Log removido
    console.error(`PortalCRUD_atualizarStatusLinkPortal: Erro: ${error.toString()}`);
    return { success: false, message: "Erro interno ao atualizar status do link." };
  }
}

/**
 * Gera ou atualiza um link de acesso para um fornecedor específico em uma cotação.
 * Registra/atualiza a entrada na ABA_PORTAL.
 * @param {string} idCotacao O ID da cotação.
 * @param {string} nomeFornecedor O nome do fornecedor.
 * @param {string} webAppUrlBase A URL base do Web App implantado.
 * @return {object} Um objeto { success: boolean, message: string, link?: string, token?: string, statusAnterior?: string }.
 */
function PortalCRUD_gerarOuAtualizarLinkFornecedor(idCotacao, nomeFornecedor, webAppUrlBase) {
  // Logger.log(`PortalCRUD_gerarOuAtualizarLinkFornecedor: Cotação '${idCotacao}', Fornecedor '${nomeFornecedor}'. URL Base: ${webAppUrlBase}`); // Log removido
  const resultado = { success: false, message: "Não foi possível gerar/atualizar o link." };

  if (!idCotacao || !nomeFornecedor) {
    resultado.message = "ID da Cotação e Nome do Fornecedor são obrigatórios.";
    // Logger.log(`PortalCRUD_gerarOuAtualizarLinkFornecedor: ${resultado.message}`); // Log removido
    return resultado;
  }
  if (!webAppUrlBase || !webAppUrlBase.includes("/exec")) {
    resultado.message = "URL base do Web App inválida ou não fornecida.";
    // Logger.log(`PortalCRUD_gerarOuAtualizarLinkFornecedor: ${resultado.message}`); // Log removido
    return resultado;
  }

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaPortal = planilha.getSheetByName(ABA_PORTAL);

    if (!abaPortal) {
      resultado.message = `Aba de controle de portal "${ABA_PORTAL}" não encontrada.`;
      // Logger.log(`PortalCRUD_gerarOuAtualizarLinkFornecedor: ${resultado.message}`); // Log removido
      console.error(resultado.message);
      return resultado;
    }

    const dadosAbaPortal = abaPortal.getDataRange().getValues();
    const cabecalhos = dadosAbaPortal[0];
    const indices = {};
    CABECALHOS_PORTAL.forEach(nomeConstante => {
        const indiceReal = cabecalhos.indexOf(nomeConstante);
        if (indiceReal !== -1) indices[nomeConstante] = indiceReal;
    });
    
    const idxIdCotacao = indices["ID da Cotação"];
    const idxNomeFornecedor = indices["Nome Fornecedor"];
    const idxToken = indices["Token Acesso"];
    const idxLink = indices["Link Acesso"];
    const idxStatus = indices["Status"];
    const idxDataEnvio = indices["Data Envio"];
    const idxDataResposta = indices["Data Resposta"];

    if ([idxIdCotacao, idxNomeFornecedor, idxToken, idxLink, idxStatus, idxDataEnvio, idxDataResposta].some(idx => idx === undefined)) {
      resultado.message = "Configuração de colunas da aba Portal está incorreta para geração de links.";
      // Logger.log(`PortalCRUD_gerarOuAtualizarLinkFornecedor: ${resultado.message}. Verifique CABECALHOS_PORTAL e a planilha.`); // Log removido
      console.error(resultado.message);
      return resultado;
    }

    let linhaExistente = -1;
    let statusAnterior = null;

    for (let i = 1; i < dadosAbaPortal.length; i++) {
      if (String(dadosAbaPortal[i][idxIdCotacao]).trim() === String(idCotacao).trim() &&
          String(dadosAbaPortal[i][idxNomeFornecedor]).trim() === String(nomeFornecedor).trim()) {
        linhaExistente = i + 1; 
        statusAnterior = String(dadosAbaPortal[i][idxStatus]).trim();
        resultado.statusAnterior = statusAnterior;
        break;
      }
    }

    const novoToken = Utilities.getUuid();
    const novoLink = `${webAppUrlBase}?token=${novoToken}`;
    const dataAtual = new Date();

    if (linhaExistente !== -1) {
      if (statusAnterior === STATUS_PORTAL.RESPONDIDO || statusAnterior === STATUS_PORTAL.FECHADO) {
        resultado.success = false; 
        resultado.message = `Link para ${nomeFornecedor} na cotação ${idCotacao} já está como '${statusAnterior}'. Não foi regerado.`;
        resultado.link = dadosAbaPortal[linhaExistente-1][idxLink]; 
        resultado.token = dadosAbaPortal[linhaExistente-1][idxToken];
        // Logger.log(`PortalCRUD_gerarOuAtualizarLinkFornecedor: ${resultado.message}`); // Log removido
        return resultado;
      }
      abaPortal.getRange(linhaExistente, idxToken + 1).setValue(novoToken);
      abaPortal.getRange(linhaExistente, idxLink + 1).setValue(novoLink);
      abaPortal.getRange(linhaExistente, idxStatus + 1).setValue(STATUS_PORTAL.LINK_GERADO);
      abaPortal.getRange(linhaExistente, idxDataEnvio + 1).setValue(dataAtual);
      abaPortal.getRange(linhaExistente, idxDataResposta + 1).clearContent(); 
      resultado.message = `Link atualizado para ${nomeFornecedor} na cotação ${idCotacao}.`;
      // Logger.log(`PortalCRUD_gerarOuAtualizarLinkFornecedor: ${resultado.message} (Linha ${linhaExistente})`); // Log removido
    } else {
      const novaLinhaDados = [];
      for(let c=0; c < cabecalhos.length; c++) { 
        if (c === idxIdCotacao) novaLinhaDados.push(idCotacao);
        else if (c === idxNomeFornecedor) novaLinhaDados.push(nomeFornecedor);
        else if (c === idxToken) novaLinhaDados.push(novoToken);
        else if (c === idxLink) novaLinhaDados.push(novoLink);
        else if (c === idxStatus) novaLinhaDados.push(STATUS_PORTAL.LINK_GERADO);
        else if (c === idxDataEnvio) novaLinhaDados.push(dataAtual);
        else if (c === idxDataResposta) novaLinhaDados.push(""); // Data Resposta vazia inicialmente
        else novaLinhaDados.push(""); 
      }
      abaPortal.appendRow(novaLinhaDados);
      resultado.message = `Novo link gerado para ${nomeFornecedor} na cotação ${idCotacao}.`;
      // Logger.log(`PortalCRUD_gerarOuAtualizarLinkFornecedor: ${resultado.message}`); // Log removido
    }

    resultado.success = true;
    resultado.link = novoLink;
    resultado.token = novoToken;
    return resultado;

  } catch (error) {
    // Logger.log(`ERRO CRÍTICO em PortalCRUD_gerarOuAtualizarLinkFornecedor: ${error.toString()} Stack: ${error.stack}`); // Log removido
    console.error(`PortalCRUD_gerarOuAtualizarLinkFornecedor: Erro: ${error.toString()}`);
    resultado.success = false;
    resultado.message = "Erro interno ao gerar/atualizar o link do fornecedor.";
    return resultado;
  }
}

/**
 * Busca os dados de um pedido finalizado para um fornecedor específico em uma cotação.
 * Procura por itens que tenham uma quantidade definida na coluna "Comprar".
 * @param {string|number} idCotacao O ID da cotação.
 * @param {string} nomeFornecedor O nome do fornecedor.
 * @return {object|null} Um objeto com os dados do pedido ou null se ocorrer um erro.
 */
function PortalCRUD_buscarDadosDoPedidoFinalizado(idCotacao, nomeFornecedor) {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaCotacoes = planilha.getSheetByName(ABA_COTACOES);

    if (!abaCotacoes) {
      console.error(`PortalCRUD_buscarDadosDoPedidoFinalizado: Aba "${ABA_COTACOES}" não encontrada.`);
      return null;
    }

    const dadosCompletos = abaCotacoes.getDataRange().getValues();
    const cabecalhosPlanilha = dadosCompletos[0];

    // ALTERAÇÃO: Removida a busca pela coluna "Quantidade a Comprar".
    const colunasNecessarias = {
      idCotacao: "ID da Cotação",
      fornecedor: "Fornecedor",
      comprar: "Comprar", // Esta coluna agora contém a QUANTIDADE.
      empresaFaturada: "Empresa Faturada",
      condicaoPagamento: "Condição de Pagamento",
      subproduto: "SubProduto",
      un: "UN",
      preco: "Preço"
    };

    const indicesMapeados = {};
    for (const chave in colunasNecessarias) {
      const nomeColuna = colunasNecessarias[chave];
      const indice = cabecalhosPlanilha.indexOf(nomeColuna);
      if (indice === -1) {
        console.error(`PortalCRUD_buscarDadosDoPedidoFinalizado: Cabeçalho essencial "${nomeColuna}" não encontrado na aba "${ABA_COTACOES}". Verifique se a coluna existe.`);
        return null;
      }
      indicesMapeados[chave] = indice;
    }

    const itensComprados = [];
    let pedidoInfo = {
      pedidoExiste: false,
      empresaFaturada: "Não informado",
      cnpj: "Não informado",
      condicaoPagamento: "Não informado"
    };

    const idCotacaoStr = String(idCotacao);
    const nomeFornecedorStr = String(nomeFornecedor);

    for (let i = 1; i < dadosCompletos.length; i++) {
      const linhaAtual = dadosCompletos[i];
      const idCotacaoLinha = String(linhaAtual[indicesMapeados.idCotacao]).trim();
      const fornecedorLinha = String(linhaAtual[indicesMapeados.fornecedor]).trim();

      if (idCotacaoLinha === idCotacaoStr && fornecedorLinha === nomeFornecedorStr) {
        // ALTERAÇÃO: A lógica agora verifica se a coluna "Comprar" contém um número válido e maior que zero.
        const valorDaColunaComprar = linhaAtual[indicesMapeados.comprar];
        const quantidade = parseInt(valorDaColunaComprar, 10);

        if (quantidade && !isNaN(quantidade) && quantidade > 0) {
            
          if (!pedidoInfo.pedidoExiste) {
            pedidoInfo.pedidoExiste = true;
            pedidoInfo.empresaFaturada = linhaAtual[indicesMapeados.empresaFaturada] || "Não informado";
            pedidoInfo.condicaoPagamento = linhaAtual[indicesMapeados.condicaoPagamento] || "Não informado";
          }

          const preco = parseFloat(linhaAtual[indicesMapeados.preco]) || 0;
          const valorTotal = preco * quantidade;

          itensComprados.push({
            subproduto: linhaAtual[indicesMapeados.subproduto],
            un: linhaAtual[indicesMapeados.un],
            precoUnit: preco.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }),
            quantidade: quantidade, // A quantidade agora vem diretamente da coluna "Comprar"
            valorTotal: valorTotal.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })
          });
        }
      }
    }

    if (pedidoInfo.pedidoExiste) {
      pedidoInfo.itensComprados = itensComprados;
      return pedidoInfo;
    } else {
      return { pedidoExiste: false }; 
    }

  } catch (error) {
    console.error(`PortalCRUD_buscarDadosDoPedidoFinalizado: Erro para Cotação ID '${idCotacao}', Fornecedor '${nomeFornecedor}': ${error.toString()}`);
    return null;
  }
}