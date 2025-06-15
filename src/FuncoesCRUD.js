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
