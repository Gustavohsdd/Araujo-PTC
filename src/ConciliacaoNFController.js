/**
 * @file ConciliacaoNFController.gs
 * @description Orquestra o processo de leitura dos arquivos XML de NF-e da pasta do Drive
 * e chama as funções CRUD para popular a planilha de dados.
 */

/**
 * Função principal para processar todos os arquivos XML da pasta de entrada.
 * Esta função deve ser executada manualmente ou por um acionador de tempo (trigger).
 */
function ConciliacaoNFController_processarXmlsDaPasta() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('Não foi possível obter o lock. Outro processo já está em execução.');
    return;
  }

  Logger.log('INICIANDO EXECUÇÃO OTIMIZADA...');

  // MUDANÇA: Arrays para acumular todos os dados antes de escrever na planilha.
  const todosOsDados = {
    notasFiscais: [],
    itensNf: [],
    faturasNf: [],
    transporteNf: [],
    tributosTotaisNf: []
  };
  
  try {
    const pastaProcessados = ConciliacaoNFCrud_garantirPastaProcessados();
    if (!pastaProcessados) {
      Logger.log('FALHA: Não foi possível encontrar ou criar a pasta de destino "Processados". Abortando.');
      return;
    }

    const pastaXml = DriveApp.getFolderById(ID_PASTA_XML);
    const todosOsArquivos = pastaXml.getFiles();
    
    const arquivosXml = [];
    while(todosOsArquivos.hasNext()) {
      const arquivo = todosOsArquivos.next();
      if (arquivo.getName().toLowerCase().endsWith('.xml')) {
        arquivosXml.push(arquivo);
      }
    }

    if (arquivosXml.length === 0) {
      Logger.log('Nenhum arquivo .xml para processar.');
      return;
    }
    
    const chavesExistentes = ConciliacaoNFCrud_obterChavesDeAcessoExistentes();
    Logger.log(`Encontradas ${chavesExistentes.size} chaves existentes.`);

    let arquivosProcessados = 0;
    let arquivosSkipped = 0;

    for (const arquivo of arquivosXml) {
      const nomeArquivo = arquivo.getName();
      Logger.log(`--- Processando arquivo em memória: ${nomeArquivo} ---`);
      
      try {
        const conteudoXml = arquivo.getBlob().getDataAsString('UTF-8');
        const dadosNf = ConciliacaoNFCrud_parsearConteudoXml(conteudoXml);

        if (!dadosNf || !dadosNf.notasFiscais.chaveAcesso) {
          Logger.log(`AVISO: Chave de acesso não encontrada no XML do arquivo ${nomeArquivo}. Pulando.`);
          arquivosSkipped++;
          arquivo.moveTo(pastaProcessados);
          continue;
        }
        
        const chaveAtual = dadosNf.notasFiscais.chaveAcesso;

        if (chavesExistentes.has(chaveAtual)) {
          Logger.log(`AVISO: A NF-e com chave ${chaveAtual} já existe na planilha. Pulando.`);
          arquivosSkipped++;
          arquivo.moveTo(pastaProcessados);
          continue;
        }
        
        // MUDANÇA: Em vez de salvar na planilha, acumula os dados nos arrays.
        todosOsDados.notasFiscais.push(dadosNf.notasFiscais);
        todosOsDados.itensNf.push(...dadosNf.itensNf);
        todosOsDados.faturasNf.push(...dadosNf.faturasNf);
        todosOsDados.transporteNf.push(...dadosNf.transporteNf);
        todosOsDados.tributosTotaisNf.push(...dadosNf.tributosTotaisNf);

        chavesExistentes.add(chaveAtual);
        Logger.log(`Dados do arquivo ${nomeArquivo} processados e acumulados em memória.`);
        
        arquivo.moveTo(pastaProcessados);
        arquivosProcessados++;

      } catch (e) {
        Logger.log(`ERRO CRÍTICO ao processar o arquivo ${nomeArquivo}. Erro: ${e.message}. Stack: ${e.stack}`);
        arquivo.moveTo(pastaProcessados);
        arquivosSkipped++;
      }
    }
    
    // MUDANÇA: Após o loop, se houver dados para salvar, chama a nova função de salvamento em lote.
    if (todosOsDados.notasFiscais.length > 0) {
        Logger.log(`Iniciando salvamento em lote de ${todosOsDados.notasFiscais.length} nota(s) fiscal(is).`);
        ConciliacaoNFCrud_salvarDadosEmLote(todosOsDados);
        Logger.log(`Salvamento em lote finalizado.`);
    } else {
        Logger.log("Nenhuma nova nota fiscal para adicionar à planilha.");
    }
    
    Logger.log(`--- Processamento de todos os arquivos finalizado. Resumo: ${arquivosProcessados} processados com sucesso, ${arquivosSkipped} pulados/com erro. ---`);

  } catch (error) {
    Logger.log(`ERRO GERAL na função ConciliacaoNFController_processarXmlsDaPasta: ${error.toString()}. Stack: ${error.stack}`);
  } finally {
    lock.releaseLock();
    Logger.log('FINALIZANDO EXECUÇÃO. Lock liberado.');
  }
}

/**
 * @file ConciliacaoNFController.gs
 * @description Funções adicionadas para o processo de conciliação.
 */

/**
 * Obtém os dados necessários para popular a página de conciliação.
 * Retorna uma lista de cotações abertas e notas fiscais não conciliadas.
 * @returns {object} Objeto com { success: boolean, dados: object|null, message: string|null }.
 */
function ConciliacaoNFController_obterDadosParaPagina() {
  try {
    Logger.log("ConciliacaoNFController: Obtendo TODOS os dados para a página de conciliação.");
    
    // 1. Obtém as listas principais (isso você já faz)
    const cotacoes = ConciliacaoNFCrud_obterCotacoesAbertas();
    const notasFiscais = ConciliacaoNFCrud_obterNFsNaoConciliadas();

    if (cotacoes === null || notasFiscais === null) {
      throw new Error("Falha ao buscar dados das planilhas (Cotações ou Notas Fiscais).");
    }

    // 2. OTIMIZAÇÃO: Busca todos os dados secundários de uma só vez
    const chavesNFsNaoConciliadas = notasFiscais.map(nf => nf.chaveAcesso);
    const chavesCotacoesAbertas = cotacoes.map(c => ({ idCotacao: c.idCotacao, fornecedor: c.fornecedor }));

    // Busca todos os itens e totais em lote, usando as chaves que coletamos
    const todosOsItensNF = ConciliacaoNFCrud_obterItensDasNFs(chavesNFsNaoConciliadas);
    const todosOsDadosGeraisNF = ConciliacaoNFCrud_obterDadosGeraisDasNFs(chavesNFsNaoConciliadas);
    const todosOsItensCotacao = ConciliacaoNFCrud_obterTodosItensCotacoesAbertas(chavesCotacoesAbertas); // Nova função que vamos criar

    Logger.log(`Dados carregados: ${cotacoes.length} cotações, ${notasFiscais.length} NFs.`);
    
    // 3. Retorna tudo para o frontend em um único objeto
    return {
      success: true,
      dados: {
        cotacoes: cotacoes,
        notasFiscais: notasFiscais,
        // --- DADOS ADICIONAIS PARA PERFORMANCE ---
        itensNF: todosOsItensNF,
        itensCotacao: todosOsItensCotacao,
        dadosGeraisNF: todosOsDadosGeraisNF
      },
      message: null
    };

  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_obterDadosParaPagina: ${e.toString()}\n${e.stack}`);
    return { success: false, dados: null, message: e.message };
  }
}

/**
 * [MODIFICADO] Apenas coleta os dados da Cotação e das NFs para a conciliação manual na interface.
 * @param {string} compositeKeyCotacao - A chave composta da cotação (ex: "54-Ambev").
 * @param {Array<string>} chavesAcessoNF - Um array de chaves de acesso das NFs.
 * @returns {object} Objeto com { success: boolean, dados: object|null, message: string|null }.
 */
function ConciliacaoNFController_realizarComparacao(compositeKeyCotacao, chavesAcessoNF) {
  try {
    if (!compositeKeyCotacao || !chavesAcessoNF || chavesAcessoNF.length === 0) {
      throw new Error("Chave da Cotação e Chaves de Acesso das NFs são obrigatórios.");
    }
    
    // Interpreta a chave composta
    const parts = compositeKeyCotacao.split('-');
    const idCotacao = parts.shift();
    const nomeFornecedor = parts.join('-');

    Logger.log(`Coletando dados para conciliação manual. Cotação ID: ${idCotacao}, Fornecedor: ${nomeFornecedor}`);
    
    // 1. Obter itens da Cotação
    const itensCotacao = ConciliacaoNFCrud_obterItensDaCotacao(idCotacao, nomeFornecedor);
    if (itensCotacao.length === 0) {
      return { success: false, message: "Nenhum item marcado para comprar nesta cotação para este fornecedor." };
    }

    // 2. Obter itens das NFs
    const itensNF = ConciliacaoNFCrud_obterItensDasNFs(chavesAcessoNF);

    // 3. Obter dados gerais (continuam úteis)
    const dadosGeraisNF = ConciliacaoNFCrud_obterDadosGeraisDasNFs(chavesAcessoNF);
    const analisePrazo = {}; // Lógica de prazo pode ser mantida ou simplificada se necessário

    // 4. Montar o objeto de dados brutos para a interface
    const dadosParaPagina = {
      idCotacao: idCotacao,
      nomeFornecedor: nomeFornecedor,
      chavesAcessoNF: chavesAcessoNF,
      itensCotacao: itensCotacao, // Lista de itens da cotação
      itensNF: itensNF,           // Lista de itens da NF
      dadosGeraisNF: dadosGeraisNF,
      analisePrazo: analisePrazo
    };

    return { success: true, dados: dadosParaPagina };

  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_realizarComparacao: ${e.toString()}\n${e.stack}`);
    return { success: false, dados: null, message: e.message };
  }
}

/**
 * [CORRIGIDO] Salva o resultado da conciliação. Usa o ID da Cotação e o Nome do Fornecedor.
 * @param {object} dadosConciliacao - O objeto de resultado gerado por realizarComparacao.
 * @returns {object} Objeto com { success: boolean, message: string|null }.
 */
function ConciliacaoNFController_salvarConciliacao(dadosConciliacao) {
  try {
    const { idCotacao, nomeFornecedor, chavesAcessoNF, itensConciliados, itensSomenteCotacao } = dadosConciliacao;
    Logger.log(`Salvando conciliação para Cotação ID: ${idCotacao}, Fornecedor: ${nomeFornecedor}`);

    // Atualiza Planilha de Cotações usando id e fornecedor
    const sucessoCotacoes = ConciliacaoNFCrud_atualizarStatusCotacao(idCotacao, nomeFornecedor, itensConciliados, itensSomenteCotacao);
    if (!sucessoCotacoes) {
        throw new Error("Falha ao atualizar a planilha de cotações.");
    }
    
    // Atualiza Planilha de Notas Fiscais
    const sucessoNF = ConciliacaoNFCrud_atualizarStatusNF(chavesAcessoNF, idCotacao, "Conciliada");
    if (!sucessoNF) {
        throw new Error("Falha ao atualizar a planilha de notas fiscais.");
    }

    Logger.log("Conciliação salva com sucesso.");
    return { success: true, message: "Conciliação salva com sucesso!" };

  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_salvarConciliacao: ${e.toString()}\n${e.stack}`);
    return { success: false, message: e.message };
  }
}

/**
 * [NOVA FUNÇÃO] Atualiza o status de uma ou mais NFs para "Sem Pedido".
 * @param {Array<string>} chavesAcessoNF - Um array de chaves de acesso das NFs a serem atualizadas.
 * @returns {object} Objeto com { success: boolean, message: string|null }.
 */
function ConciliacaoNFController_marcarNFsSemPedido(chavesAcessoNF) {
  try {
    if (!chavesAcessoNF || !Array.isArray(chavesAcessoNF) || chavesAcessoNF.length === 0) {
      throw new Error("Nenhuma chave de acesso foi fornecida.");
    }
    
    Logger.log(`Marcando ${chavesAcessoNF.length} NF(s) como 'Sem Pedido'.`);

    // A função de atualização no CRUD foi generalizada para aceitar qualquer status.
    const sucesso = ConciliacaoNFCrud_atualizarStatusNF(chavesAcessoNF, null, "Sem Pedido");
    
    if (!sucesso) {
      throw new Error("Falha ao atualizar o status das notas fiscais na planilha.");
    }

    Logger.log("Notas fiscais marcadas como 'Sem Pedido' com sucesso.");
    return { success: true, message: `${chavesAcessoNF.length} nota(s) fiscal(is) marcada(s) como 'Sem Pedido' com sucesso!` };

  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_marcarNFsSemPedido: ${e.toString()}\n${e.stack}`);
    return { success: false, message: e.message };
  }
}