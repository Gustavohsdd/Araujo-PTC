/**
 * @file NotasFiscaisController.js
 * @description Orquestra as operações do módulo de gerenciamento de Notas Fiscais.
 */

/**
 * Lista todas as notas fiscais com base nos filtros fornecidos.
 * @param {object} filtros Objeto contendo os filtros (status, fornecedor, dataInicio, dataFim, busca).
 * @returns {{success: boolean, dados?: object[], message?: string}}
 */
function NotasFiscaisController_listarNotas(filtros) {
  try {
    Logger.log('Controller_listarNotas: início. Filtros=' + JSON.stringify(filtros || {}));

    // Validação básica dos IDs/Abas (sem acesso pesado)
    if (!ID_PLANILHA_NF) {
      return { success: false, message: 'ID da planilha de NF não definido (ID_PLANILHA_NF).' };
    }
    // Tenta abrir e checar pelo menos a aba base
    const planilhaNF = SpreadsheetApp.openById(ID_PLANILHA_NF);
    const abaNF = planilhaNF.getSheetByName(ABA_NF_NOTAS_FISCAIS);
    if (!abaNF) {
      return { success: false, message: 'Aba de Notas Fiscais não encontrada: ' + ABA_NF_NOTAS_FISCAIS };
    }

    const dados = NotasFiscaisCRUD_obterTodasAsNotas(filtros);
    Logger.log('Controller_listarNotas: sucesso. Qtde=' + (dados ? dados.length : 0));
    return { success: true, dados: dados };
  } catch (e) {
    Logger.log('ERRO em NotasFiscaisController_listarNotas: ' + e.toString() + '\n' + e.stack);
    return { success: false, message: 'Erro ao listar notas: ' + e.message };
  }
}

/**
 * Atualiza o status de uma NF. Trata o retorno da função de conciliação.
 * @param {string} chaveAcesso A chave de acesso da NF a ser atualizada.
 * @param {string} novoStatus O novo status para a NF (ex: "Bonificação", "NF Tipo B", "Pendente").
 * @returns {{success: boolean, message: string}}
 */
function NotasFiscaisController_atualizarStatusNF(chaveAcesso, novoStatus) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: 'Outro processo está em execução. Tente novamente em alguns instantes.' };
  }
  try {
    Logger.log(`Controller: Atualizando status da NF ${chaveAcesso} para ${novoStatus}`);
    // A função ConciliacaoNFCrud_atualizarStatusNF deve ser robusta.
    // Assumimos que ela lança um erro em caso de falha.
    ConciliacaoNFCrud_atualizarStatusNF([chaveAcesso], null, novoStatus);
    return { success: true, message: `Status da NF atualizado para "${novoStatus}" com sucesso.` };
    
  } catch (e) {
    Logger.log(`ERRO em NotasFiscaisController_atualizarStatusNF: ${e.toString()}\n${e.stack}`);
    return { success: false, message: `Erro ao atualizar status: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Desfaz a conciliação de uma NF.
 * 1) Reseta a NF para "Pendente" (e "Status do Rateio" para "Pendente")
 * 2) Limpa campos da aba Cotações, localizando por "Número da Nota"
 * 3) Ajusta "Status da Cotação" para "Aguardando Faturamento" (somente dos itens dessa nota)
 * 4) Remove lançamentos de Contas a Pagar da chave
 *
 * @param {string} chaveAcesso
 * @param {string|number} numeroNF
 * @returns {{success:boolean, message:string}}
 */
function NotasFiscaisController_desfazerConciliacao(chaveAcesso, numeroNF) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: 'Outro processo está em execução. Tente novamente.' };
  }
  try {
    if (!chaveAcesso) throw new Error('Chave de acesso não informada.');
    if (numeroNF === undefined || numeroNF === null || String(numeroNF).trim() === '') {
      throw new Error('Número da Nota não informado.');
    }

    Logger.log(`[Desfazer] Iniciando para chave=${chaveAcesso}, numeroNF=${numeroNF}`);

    // 4) Remove Contas a Pagar
    NotasFiscaisCRUD_excluirContasAPagarPorChave(chaveAcesso);

    // 2) Limpa campos da aba Cotações por Número da Nota
    NotasFiscaisCRUD_limparDadosCotacoesPorNumeroNota(numeroNF);

    // 3) Ajusta status da cotação para "Aguardando Faturamento"
    NotasFiscaisCRUD_atualizarStatusCotacoesPorNumeroNota(numeroNF, 'Aguardando Faturamento');

    // 1) Reseta NF para "Pendente" (e rateio p/ "Pendente")
    NotasFiscaisCRUD_resetarStatusNF(chaveAcesso);

    // (Opcional) se existir rotina de conciliação central, tenta marcar como Pendente também
    try {
      if (typeof ConciliacaoNFCrud_atualizarStatusNF === 'function') {
        ConciliacaoNFCrud_atualizarStatusNF([chaveAcesso], null, 'Pendente');
      }
    } catch (eOpc) {
      Logger.log('[Desfazer] API ConciliacaoNFCrud_atualizarStatusNF não disponível. Prosseguindo mesmo assim.');
    }

    return {
      success: true,
      message: 'Conciliação desfeita: NF voltou a "Pendente", Cotações limpas/ajustadas e Contas a Pagar removidas.'
    };
  } catch (e) {
    Logger.log('ERRO em NotasFiscaisController_desfazerConciliacao: ' + e.toString() + '\n' + e.stack);
    return { success: false, message: 'Erro ao desfazer conciliação: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Obtém o resumo financeiro de uma NF, incluindo faturas e contas a pagar.
 * @param {string} chaveAcesso A chave de acesso da NF.
 * @returns {{success: boolean, dados?: {faturas: object[], contasAPagar: object[]}, message?: string}}
 */
function NotasFiscaisController_obterResumoFinanceiroDaNF(chaveAcesso) {
  try {
    const dados = NotasFiscaisCRUD_obterResumoFinanceiroDaNF(chaveAcesso);
    return { success: true, dados: dados };
  } catch (e) {
    Logger.log(`ERRO em NotasFiscaisController_obterResumoFinanceiroDaNF: ${e.toString()}\n${e.stack}`);
    return { success: false, message: `Erro ao obter resumo financeiro: ${e.message}` };
  }
}

/**
 * Salva (substitui) as faturas de uma NF.
 * @param {string} chaveAcesso A chave de acesso da NF.
 * @param {object[]} faturas Array de objetos de fatura a serem salvos.
 * @returns {{success: boolean, message: string}}
 */
function NotasFiscaisController_salvarFaturas(chaveAcesso, faturas) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: 'Outro processo está em execução. Tente novamente.' };
  }
  try {
    NotasFiscaisCRUD_substituirFaturasDaNF(chaveAcesso, faturas);
    return { success: true, message: "Faturas salvas com sucesso!" };
  } catch (e) {
    Logger.log(`ERRO em NotasFiscaisController_salvarFaturas: ${e.toString()}\n${e.stack}`);
    return { success: false, message: `Erro ao salvar faturas: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Salva (substitui) os lançamentos de contas a pagar de uma NF.
 * @param {string} chaveAcesso A chave de acesso da NF.
 * @param {object[]} linhas Array de objetos de contas a pagar a serem salvos.
 * @returns {{success: boolean, message: string}}
 */
function NotasFiscaisController_salvarContasAPagar(chaveAcesso, linhas) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: 'Outro processo está em execução. Tente novamente.' };
  }
  try {
    NotasFiscaisCRUD_substituirContasAPagarDaNF(chaveAcesso, linhas);
    return { success: true, message: "Contas a Pagar salvas com sucesso!" };
  } catch (e) {
    Logger.log(`ERRO em NotasFiscaisController_salvarContasAPagar: ${e.toString()}\n${e.stack}`);
    return { success: false, message: `Erro ao salvar contas a pagar: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

function NotasFiscaisController_ping() {
  try {
    Logger.log('Ping controller OK');
    return { success: true, message: 'Ping OK (Controller ativo)' };
  } catch (e) {
    Logger.log('Erro em NotasFiscaisController_ping: ' + e.toString() + '\n' + e.stack);
    return { success: false, message: 'Falha no ping: ' + e.message };
  }
}

/**
 * NotasFiscaisController_listarSetoresRegrasRateio
 * Controller para expor a lista de setores do rateio ao frontend.
 * @returns {{success:boolean, dados?: string[], message?: string}}
 */
function NotasFiscaisController_listarSetoresRegrasRateio() {
  try {
    var setores = NotasFiscaisCRUD_listarSetoresRegrasRateio();
    return { success: true, dados: setores };
  } catch (e) {
    Logger.log('ERRO em NotasFiscaisController_listarSetoresRegrasRateio: ' + e.toString() + '\n' + e.stack);
    return { success: false, message: 'Erro ao listar setores do rateio: ' + e.message };
  }
}
