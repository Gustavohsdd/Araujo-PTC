// @ts-nocheck
/**
 * @file RelatoriosController.gs
 * @description Controlador para a geração de relatórios.
 */

/**
 * Ponto de entrada para obter os dados do relatório de Análise de Compra.
 * @param {string} idCotacao O ID da cotação para a qual o relatório está sendo gerado.
 * @return {object} Um objeto com { success: boolean, dados: Array<object>, message?: string }.
 */
function RelatoriosController_obterDadosAnaliseCompra(idCotacao) {
  Logger.log(`RelatoriosController: Solicitado relatório de Análise de Compra para Cotação ID '${idCotacao}'.`);
  if (!idCotacao) {
    return { success: false, message: "O ID da Cotação é obrigatório para gerar o relatório." };
  }
  
  try {
    const dadosRelatorio = RelatoriosCRUD_gerarDadosAnaliseCompra(idCotacao);
    
    if (dadosRelatorio === null) {
        return { success: false, message: "Falha ao processar os dados do relatório na camada CRUD." };
    }

    return { success: true, dados: dadosRelatorio };

  } catch (error) {
    Logger.log(`ERRO CRÍTICO em RelatoriosController_obterDadosAnaliseCompra: ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: `Erro no servidor ao gerar relatório: ${error.message}` };
  }
}
