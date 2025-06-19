// @ts-nocheck

//####################################################################################################
// MÓDULO: FUNCOES (SERVER-SIDE CRUD)
// Funções CRUD para as opções do menu "Funções".
//####################################################################################################

/**
 * @file FuncoesCRUD.gs
 * @description Funções CRUD para as funcionalidades do menu "Funções", como "Gerenciar Cotações (Portal)".
 */

/**
 * Obtém dados das cotações e seus fornecedores da aba Portal para a funcionalidade "Gerenciar Cotações".
 * (Originada de PortalCRUD_getDadosGerenciarCotacoes)
 * @return {object} { success: boolean, dados: Array<object>, message?: string }
 */
function FuncoesCRUD_getDadosGerenciarCotacoes() {
  Logger.log("FuncoesCRUD_getDadosGerenciarCotacoes: Iniciando.");
  const resultado = { success: false, dados: [], message: "" };
  const NOME_ABA_PORTAL = ABA_PORTAL; // Constante global
  const NOME_COLUNA_TEXTO_PERSONALIZADO = "Texto Personalizado Link"; // Esta coluna deve existir na aba Portal

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaPortal = planilha.getSheetByName(NOME_ABA_PORTAL);

    if (!abaPortal) {
      resultado.message = `Aba "${NOME_ABA_PORTAL}" não encontrada.`;
      Logger.log(resultado.message);
      return resultado;
    }

    const ultimaLinha = abaPortal.getLastRow();
    if (ultimaLinha < 1) {
        resultado.message = `Aba "${NOME_ABA_PORTAL}" está completamente vazia.`;
        resultado.success = true; 
        Logger.log(resultado.message);
        return resultado;
    }
    if (ultimaLinha < 2 && ultimaLinha >=1) { // Apenas cabeçalho
      resultado.message = `Aba "${NOME_ABA_PORTAL}" contém apenas cabeçalho.`;
      resultado.success = true; 
      Logger.log(resultado.message);
      return resultado;
    }

    const dadosPlanilha = abaPortal.getRange(1, 1, ultimaLinha, abaPortal.getLastColumn()).getValues();
    const cabecalhos = dadosPlanilha[0].map(c => String(c).trim());

    // Assumindo CABECALHOS_PORTAL e STATUS_PORTAL como constantes globais
    const idxIdCotacao = cabecalhos.indexOf(CABECALHOS_PORTAL[0]); // "ID da Cotação"
    const idxFornecedor = cabecalhos.indexOf(CABECALHOS_PORTAL[1]); // "Nome Fornecedor"
    const idxTokenLink = cabecalhos.indexOf(CABECALHOS_PORTAL[2]); // "Token Acesso"
    const idxLinkAcesso = cabecalhos.indexOf(CABECALHOS_PORTAL[3]); // "Link Acesso"
    const idxStatusResposta = cabecalhos.indexOf(CABECALHOS_PORTAL[4]); // "Status"
    const idxTextoPersonalizado = cabecalhos.indexOf(NOME_COLUNA_TEXTO_PERSONALIZADO);


    if ([idxIdCotacao, idxFornecedor, idxTokenLink, idxLinkAcesso, idxStatusResposta].some(idx => idx === -1)) {
      let colunasFaltantes = [];
      if(idxIdCotacao === -1) colunasFaltantes.push(`"${CABECALHOS_PORTAL[0]}"`);
      if(idxFornecedor === -1) colunasFaltantes.push(`"${CABECALHOS_PORTAL[1]}"`);
      if(idxTokenLink === -1) colunasFaltantes.push(`"${CABECALHOS_PORTAL[2]}"`);
      if(idxLinkAcesso === -1) colunasFaltantes.push(`"${CABECALHOS_PORTAL[3]}"`);
      if(idxStatusResposta === -1) colunasFaltantes.push(`"${CABECALHOS_PORTAL[4]}"`);
      resultado.message = `Coluna(s) essencial(is) ${colunasFaltantes.join(', ')} não encontrada(s) na aba "${NOME_ABA_PORTAL}". Verifique os nomes em Constantes.gs.`;
      Logger.log(resultado.message + ` Índices encontrados: ID=${idxIdCotacao}, Forn=${idxFornecedor}, Token=${idxTokenLink}, Link=${idxLinkAcesso}, Status=${idxStatusResposta}, TextoPers=${idxTextoPersonalizado}`);
      return resultado;
    }
    if (idxTextoPersonalizado === -1) {
        console.warn(`FuncoesCRUD_getDadosGerenciarCotacoes: Coluna "${NOME_COLUNA_TEXTO_PERSONALIZADO}" não encontrada na aba "${NOME_ABA_PORTAL}". Textos personalizados não serão carregados/salvos corretamente. Adicione esta coluna à aba e à constante CABECALHOS_PORTAL.`);
    }
    
    const scriptUrlBase = ScriptApp.getService().getUrl().replace('/dev', '/exec'); 
    const cotacoesAgrupadas = {};

    for (let i = 1; i < dadosPlanilha.length; i++) { // Começa em 1 para pular o cabeçalho
      const linha = dadosPlanilha[i];
      const idCotacao = String(linha[idxIdCotacao]).trim();
      if (!idCotacao) continue; // Pula linhas sem ID de cotação

      if (!cotacoesAgrupadas[idCotacao]) {
        cotacoesAgrupadas[idCotacao] = {
          idCotacao: idCotacao,
          fornecedores: [],
          totalFornecedores: 0,
          respondidos: 0,
          textoPersonalizadoCotacao: null // Será preenchido com o texto do primeiro fornecedor ou o texto global salvo
        };
      }

      const fornecedorNome = String(linha[idxFornecedor]).trim();
      const token = String(linha[idxTokenLink]).trim();
      let linkCompleto = "";
      if (linha[idxLinkAcesso] && String(linha[idxLinkAcesso]).trim().startsWith('http')) {
          linkCompleto = String(linha[idxLinkAcesso]).trim();
      } else if (token) {
          linkCompleto = `${scriptUrlBase}?view=PortalFornecedorView&token=${token}`; // Adicionado view=PortalFornecedorView
      }
      
      let textoPersonalizadoValor = "";
      if (idxTextoPersonalizado !== -1 && linha[idxTextoPersonalizado] !== null && linha[idxTextoPersonalizado] !== undefined) {
          textoPersonalizadoValor = String(linha[idxTextoPersonalizado]);
      }

      if (cotacoesAgrupadas[idCotacao].textoPersonalizadoCotacao === null) {
          cotacoesAgrupadas[idCotacao].textoPersonalizadoCotacao = textoPersonalizadoValor;
      }

      cotacoesAgrupadas[idCotacao].fornecedores.push({
        nome: fornecedorNome,
        link: linkCompleto,
        statusResposta: String(linha[idxStatusResposta]).trim()
      });

      cotacoesAgrupadas[idCotacao].totalFornecedores++;
      if (String(linha[idxStatusResposta]).trim().toLowerCase() === STATUS_PORTAL.RESPONDIDO.toLowerCase()) { 
        cotacoesAgrupadas[idCotacao].respondidos++;
      }
    }

    resultado.dados = Object.values(cotacoesAgrupadas).map(cot => {
      cot.percentualRespondido = (cot.totalFornecedores > 0) ? (cot.respondidos / cot.totalFornecedores) * 100 : 0;
      return cot;
    });

    resultado.success = true;
    resultado.message = `${resultado.dados.length} cotações carregadas do portal (FuncoesCRUD).`;
    Logger.log(resultado.message);

  } catch (error) {
    resultado.message = `Erro ao buscar dados do portal (FuncoesCRUD): ${error.message}`;
    Logger.log(`ERRO em FuncoesCRUD_getDadosGerenciarCotacoes: ${error.toString()} Stack: ${error.stack}`);
  }
  return resultado;
}


/**
 * Exclui um fornecedor de uma cotação específica na aba Portal.
 * (Originada de PortalCRUD_excluirFornecedorDaCotacao, focada apenas na aba Portal)
 * @param {string} idCotacao O ID da cotação.
 * @param {string} nomeFornecedor O nome do fornecedor a ser excluído.
 * @return {object} { success: boolean, message: string }
 */
function FuncoesCRUD_excluirFornecedorDaCotacaoPortal(idCotacao, nomeFornecedor) {
  Logger.log(`FuncoesCRUD_excluirFornecedorDaCotacaoPortal: ID Cotação '${idCotacao}', Fornecedor '${nomeFornecedor}'.`);
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaPortal = planilha.getSheetByName(ABA_PORTAL); // Constante global
  let linhaExcluidaPortal = false;

  try {
    if (!abaPortal) {
      Logger.log(`FuncoesCRUD_excluirFornecedorDaCotacaoPortal: Aba "${ABA_PORTAL}" não encontrada.`);
      return { success: false, message: `Aba "${ABA_PORTAL}" não encontrada.` };
    }

    const ultimaLinhaPortal = abaPortal.getLastRow();
    if (ultimaLinhaPortal > 1) { // Precisa de cabeçalho + dados
      const dadosPortal = abaPortal.getRange(1, 1, ultimaLinhaPortal, abaPortal.getLastColumn()).getValues();
      const cabecalhosPortal = dadosPortal[0].map(c => String(c).trim());
      const idxIdCotacaoPortal = cabecalhosPortal.indexOf(CABECALHOS_PORTAL[0]); // "ID da Cotação"
      const idxFornecedorPortal = cabecalhosPortal.indexOf(CABECALHOS_PORTAL[1]); // "Nome Fornecedor"

      if (idxIdCotacaoPortal !== -1 && idxFornecedorPortal !== -1) {
        for (let i = dadosPortal.length - 1; i >= 1; i--) { // Itera de baixo para cima para exclusão segura
          if (String(dadosPortal[i][idxIdCotacaoPortal]).trim() === idCotacao &&
              String(dadosPortal[i][idxFornecedorPortal]).trim() === nomeFornecedor) {
            abaPortal.deleteRow(i + 1);
            linhaExcluidaPortal = true;
            Logger.log(`FuncoesCRUD_excluirFornecedorDaCotacaoPortal: Linha ${i+1} excluída da aba Portal.`);
            break; 
          }
        }
      } else {
        Logger.log(`FuncoesCRUD_excluirFornecedorDaCotacaoPortal: Colunas chave não encontradas em ABA_PORTAL.`);
        return { success: false, message: "Colunas chave não encontradas na aba Portal." };
      }
    }

    if (linhaExcluidaPortal) {
      return { success: true, message: `Fornecedor '${nomeFornecedor}' excluído da cotação '${idCotacao}' no portal.` };
    } else {
      return { success: false, message: `Fornecedor '${nomeFornecedor}' não encontrado na cotação '${idCotacao}' no portal para exclusão.` };
    }

  } catch (error) {
    Logger.log(`ERRO em FuncoesCRUD_excluirFornecedorDaCotacaoPortal: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: `Erro ao excluir fornecedor do portal (FuncoesCRUD): ${error.message}` };
  }
}


/**
 * Salva o texto personalizado GLOBAL para uma cotação na aba Portal.
 * (Originada de PortalCRUD_salvarTextoGlobalCotacao)
 * @param {string} idCotacao O ID da cotação.
 * @param {string} textoGlobal O texto global a ser salvo para a cotação.
 * @return {object} { success: boolean, message: string }
 */
function FuncoesCRUD_salvarTextoGlobalCotacaoPortal(idCotacao, textoGlobal) {
  Logger.log(`FuncoesCRUD_salvarTextoGlobalCotacaoPortal: ID Cotação '${idCotacao}'.`);
  const NOME_ABA_PORTAL = ABA_PORTAL; // Constante global
  const NOME_COLUNA_TEXTO_PERSONALIZADO = "Texto Personalizado Link"; // Certifique-se que esta coluna existe

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaPortal = planilha.getSheetByName(NOME_ABA_PORTAL);

    if (!abaPortal) {
      return { success: false, message: `Aba "${NOME_ABA_PORTAL}" não encontrada.` };
    }

    const ultimaLinha = abaPortal.getLastRow();
    if (ultimaLinha < 2) { // Precisa de cabeçalho + dados
      return { success: false, message: `Aba "${NOME_ABA_PORTAL}" com dados insuficientes ou vazia.` };
    }

    const dadosPlanilha = abaPortal.getRange(1, 1, ultimaLinha, abaPortal.getLastColumn()).getValues();
    const cabecalhos = dadosPlanilha[0].map(c => String(c).trim());

    const idxIdCotacao = cabecalhos.indexOf(CABECALHOS_PORTAL[0]); // "ID da Cotação"
    const idxTextoPersonalizado = cabecalhos.indexOf(NOME_COLUNA_TEXTO_PERSONALIZADO);

    if (idxIdCotacao === -1) {
      return { success: false, message: `Coluna "${CABECALHOS_PORTAL[0]}" não encontrada na aba "${NOME_ABA_PORTAL}".` };
    }
    if (idxTextoPersonalizado === -1) {
        return { success: false, message: `Coluna "${NOME_COLUNA_TEXTO_PERSONALIZADO}" não encontrada na aba "${NOME_ABA_PORTAL}". Adicione-a à planilha.` };
    }

    let linhasAtualizadas = 0;
    for (let i = 1; i < dadosPlanilha.length; i++) { // Começa em 1 para pular cabeçalho
      if (String(dadosPlanilha[i][idxIdCotacao]).trim() === idCotacao) {
        abaPortal.getRange(i + 1, idxTextoPersonalizado + 1).setValue(textoGlobal);
        linhasAtualizadas++;
      }
    }

    if (linhasAtualizadas > 0) {
      Logger.log(`FuncoesCRUD_salvarTextoGlobalCotacaoPortal: Texto global salvo para ${linhasAtualizadas} fornecedor(es) da Cotação '${idCotacao}'.`);
      return { success: true, message: "Texto global da cotação salvo com sucesso para todos os fornecedores." };
    } else {
      // Se a cotação não existe no portal, não é necessariamente um erro, mas nada foi atualizado.
      // Poderia-se optar por adicionar as linhas aqui se fosse desejado.
      Logger.log(`FuncoesCRUD_salvarTextoGlobalCotacaoPortal: Nenhuma entrada encontrada para a cotação '${idCotacao}' na aba Portal para atualizar o texto global.`);
      return { success: true, message: `Nenhuma entrada para cotação '${idCotacao}' encontrada no portal para atualizar texto. (Isso pode ser normal se a cotação ainda não foi enviada para fornecedores)` };
    }

  } catch (error) {
    Logger.log(`ERRO em FuncoesCRUD_salvarTextoGlobalCotacaoPortal: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: `Erro ao salvar texto global da cotação (FuncoesCRUD): ${error.message}` };
  }
}

/**
 * MELHORADO: Busca na aba "Cotacoes" os últimos dados (Preço, Tamanho, UN, Fator) para cada combinação de SubProduto e Fornecedor
 * e preenche os dados em itens da cotação atual que não possuem preço. Também recalcula os campos dependentes.
 * @param {string} idCotacaoAlvo O ID da cotação cujos dados devem ser preenchidos.
 * @return {object} Um objeto com { success: boolean, numItens: number, message: string }.
 */
function FuncoesCRUD_preencherUltimosPrecos(idCotacaoAlvo) {
  Logger.log(`FuncoesCRUD_preencherUltimosPrecos (MELHORADO): Iniciando para ID '${idCotacaoAlvo}'.`);
  
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaCotacoes = planilha.getSheetByName(ABA_COTACOES);
    if (!abaCotacoes) {
      throw new Error(`A aba "${ABA_COTACOES}" não foi encontrada.`);
    }
    
    const range = abaCotacoes.getDataRange();
    const valores = range.getValues();
    const cabecalhos = valores[0].map(String);

    // Mapear todos os índices de coluna necessários
    const idxIdCotacao = cabecalhos.indexOf("ID da Cotação");
    const idxSubProduto = cabecalhos.indexOf("SubProduto");
    const idxFornecedor = cabecalhos.indexOf("Fornecedor");
    const idxPreco = cabecalhos.indexOf("Preço");
    const idxTamanho = cabecalhos.indexOf("Tamanho");
    const idxUN = cabecalhos.indexOf("UN");
    const idxFator = cabecalhos.indexOf("Fator");
    const idxComprar = cabecalhos.indexOf("Comprar");
    const idxValorTotal = cabecalhos.indexOf("Valor Total");
    const idxPrecoPorFator = cabecalhos.indexOf("Preço por Fator");

    // Validar se todas as colunas essenciais existem
    const requiredColumns = {
        "ID da Cotação": idxIdCotacao, "SubProduto": idxSubProduto, "Fornecedor": idxFornecedor, 
        "Preço": idxPreco, "Tamanho": idxTamanho, "UN": idxUN, "Fator": idxFator,
        "Comprar": idxComprar, "Valor Total": idxValorTotal, "Preço por Fator": idxPrecoPorFator
    };

    for(const col in requiredColumns) {
        if(requiredColumns[col] === -1) {
            throw new Error(`A coluna essencial "${col}" não foi encontrada na aba 'Cotacoes'.`);
        }
    }

    const ultimosDadosMap = {};
    // 1. Construir o mapa de últimos dados, iterando do mais recente para o mais antigo.
    Logger.log("Construindo mapa de últimos dados (Preço, Tamanho, UN, Fator)...");
    for (let i = valores.length - 1; i > 0; i--) {
      const linha = valores[i];
      const idCotacaoLinha = String(linha[idxIdCotacao]).trim();
      
      if (idCotacaoLinha === idCotacaoAlvo) continue;

      const subProduto = String(linha[idxSubProduto]).trim();
      const fornecedor = String(linha[idxFornecedor]).trim();
      const preco = parseFloat(String(linha[idxPreco]).replace(",", "."));

      if (subProduto && fornecedor) {
        const chave = `${subProduto}__${fornecedor}`;
        // Se a chave ainda não existe no mapa e o preço é válido e maior que zero, guarde todos os dados.
        if (!ultimosDadosMap.hasOwnProperty(chave) && !isNaN(preco) && preco > 0) {
          ultimosDadosMap[chave] = {
            preco: preco,
            tamanho: linha[idxTamanho],
            un: linha[idxUN],
            fator: linha[idxFator]
          };
        }
      }
    }
    Logger.log(`Mapa de dados históricos construído com ${Object.keys(ultimosDadosMap).length} entradas.`);

    // 2. Aplicar os dados e recalcular campos na matriz de valores em memória
    let itensAtualizadosCount = 0;
    let foiModificado = false;
    Logger.log(`Aplicando dados históricos na cotação alvo: ${idCotacaoAlvo}`);
    for (let i = 1; i < valores.length; i++) {
      const linha = valores[i];
      const idCotacaoLinha = String(linha[idxIdCotacao]).trim();

      if (idCotacaoLinha === idCotacaoAlvo) {
        const precoAtual = parseFloat(String(linha[idxPreco]).replace(",", "."));
        let dadosForamAtualizadosNestaLinha = false;

        // A condição principal para atualização continua sendo o preço vazio/zerado
        if (isNaN(precoAtual) || precoAtual === 0) {
          const subProduto = String(linha[idxSubProduto]).trim();
          const fornecedor = String(linha[idxFornecedor]).trim();
          const chave = `${subProduto}__${fornecedor}`;

          if (ultimosDadosMap.hasOwnProperty(chave)) {
            const dadosHistoricos = ultimosDadosMap[chave];
            
            // Atualiza os quatro campos na matriz em memória
            valores[i][idxPreco] = dadosHistoricos.preco;
            valores[i][idxTamanho] = dadosHistoricos.tamanho;
            valores[i][idxUN] = dadosHistoricos.un;
            valores[i][idxFator] = dadosHistoricos.fator;
            
            itensAtualizadosCount++;
            dadosForamAtualizadosNestaLinha = true;
            foiModificado = true;
            Logger.log(`Dados para '${subProduto}' do fornecedor '${fornecedor}' atualizados na linha ${i + 1} (em memória).`);
          }
        }
        
        // Se os dados foram atualizados, precisamos recalcular os campos dependentes
        if (dadosForamAtualizadosNestaLinha) {
            const preco = parseFloat(valores[i][idxPreco]) || 0;
            const comprar = parseFloat(String(valores[i][idxComprar]).replace(",", ".")) || 0;
            const fator = parseFloat(String(valores[i][idxFator]).replace(",", ".")) || 0;

            const valorTotalCalculado = preco * comprar;
            const precoPorFatorCalculado = fator !== 0 ? preco / fator : 0;
            
            valores[i][idxValorTotal] = valorTotalCalculado;
            valores[i][idxPrecoPorFator] = precoPorFatorCalculado;
        }
      }
    }
    
    // 3. Salvar todas as alterações de volta na planilha de uma só vez, se houver modificações
    if (foiModificado) {
        range.setValues(valores);
        Logger.log("Matriz de valores modificada foi salva de volta na planilha.");
    }

    if (itensAtualizadosCount > 0) {
      return { success: true, numItens: itensAtualizadosCount, message: `Dados de ${itensAtualizadosCount} item(ns) foram atualizados com base no histórico.` };
    } else {
      return { success: true, numItens: 0, message: "Nenhum dado a ser atualizado foi encontrado no histórico para os itens desta cotação." };
    }

  } catch (error) {
    Logger.log(`ERRO CRÍTICO em FuncoesCRUD_preencherUltimosPrecos: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, numItens: 0, message: "Erro no servidor ao buscar últimos dados: " + error.message };
  } finally {
    lock.releaseLock();
  }
}