/**
 * @file ConciliacaoNFController.gs
 * @description Orquestra o processo de leitura dos arquivos XML de NF-e da pasta do Drive,
 * o upload de novos arquivos e chama as funções CRUD para popular a planilha de dados.
 */

/**
 * [FUNÇÃO ALTERADA]
 * Recebe arquivos da interface, decodifica, processa em memória e salva na planilha.
 * Suporta upload de múltiplos arquivos XML e ZIP. Os arquivos não são salvos no Drive.
 * @param {Array<object>} arquivos - Array de objetos com {fileName, mimeType, content (base64)}
 * @returns {object} Objeto com { success: boolean, message: string }.
 */
function ConciliacaoNFController_uploadArquivos(arquivos) {
  if (!arquivos || arquivos.length === 0) {
    return { success: false, message: "Nenhum arquivo recebido." };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: 'Não foi possível obter o lock. Outro processo já está em execução.' };
  }

  Logger.log('INICIANDO EXECUÇÃO DE UPLOAD E PROCESSAMENTO EM MEMÓRIA...');

  const todosOsDados = {
    notasFiscais: [],
    itensNf: [],
    faturasNf: [],
    transporteNf: [],
    tributosTotaisNf: []
  };

  let arquivosProcessados = 0;
  let arquivosDuplicados = 0;
  let arquivosComErro = 0;
  let arquivosIgnorados = [];

  try {
    const chavesExistentes = ConciliacaoNFCrud_obterChavesDeAcessoExistentes();
    Logger.log(`Encontradas ${chavesExistentes.size} chaves existentes na planilha.`);

    const processarBlobXml = (blob, nomeArquivoFonte) => {
      try {
        if (!blob.getName().toLowerCase().endsWith('.xml')) {
            Logger.log(`Arquivo "${blob.getName()}" dentro de um ZIP não é .xml e foi ignorado.`);
            return;
        }

        const conteudoXml = blob.getDataAsString('UTF-8');
        const dadosNf = ConciliacaoNFCrud_parsearConteudoXml(conteudoXml);

        if (!dadosNf || !dadosNf.notasFiscais.chaveAcesso) {
          Logger.log(`AVISO: Chave de acesso não encontrada no XML do arquivo ${nomeArquivoFonte}. Pulando.`);
          arquivosComErro++;
          return;
        }

        const chaveAtual = dadosNf.notasFiscais.chaveAcesso;
        if (chavesExistentes.has(chaveAtual)) {
          Logger.log(`AVISO: A NF-e com chave ${chaveAtual} (arquivo: ${nomeArquivoFonte}) já existe. Pulando.`);
          arquivosDuplicados++;
          return;
        }

        todosOsDados.notasFiscais.push(dadosNf.notasFiscais);
        todosOsDados.itensNf.push(...dadosNf.itensNf);
        todosOsDados.faturasNf.push(...dadosNf.faturasNf);
        todosOsDados.transporteNf.push(...dadosNf.transporteNf);
        todosOsDados.tributosTotaisNf.push(...dadosNf.tributosTotaisNf);

        chavesExistentes.add(chaveAtual); // Adiciona à lista para evitar duplicidade no mesmo lote
        arquivosProcessados++;
        Logger.log(`Dados do arquivo ${nomeArquivoFonte} (chave: ${chaveAtual}) processados e acumulados em memória.`);

      } catch (e) {
        Logger.log(`ERRO CRÍTICO ao processar o arquivo em memória ${nomeArquivoFonte}. Erro: ${e.message}. Stack: ${e.stack}`);
        arquivosComErro++;
      }
    };

    for (const arquivo of arquivos) {
      const { fileName, mimeType, content } = arquivo;
      const decodedContent = Utilities.base64Decode(content);
      const blob = Utilities.newBlob(decodedContent, mimeType, fileName);
      const nomeArquivoLower = fileName.toLowerCase();

      if (mimeType === 'application/zip' || mimeType === 'application/x-zip-compressed' || nomeArquivoLower.endsWith('.zip')) {
        try {
          const arquivosDescompactados = Utilities.unzip(blob);
          for (const arquivoDescompactado of arquivosDescompactados) {
            processarBlobXml(arquivoDescompactado, `${fileName}/${arquivoDescompactado.getName()}`);
          }
        } catch (e) {
          Logger.log(`Erro ao descompactar o arquivo ZIP "${fileName}": ${e.message}`);
          arquivosIgnorados.push(fileName);
        }
      } else if (mimeType === 'text/xml' || mimeType === 'application/xml' || nomeArquivoLower.endsWith('.xml')) {
        processarBlobXml(blob, fileName);
      } else {
        Logger.log(`Arquivo ${fileName} com tipo ${mimeType} não é suportado e foi ignorado.`);
        arquivosIgnorados.push(fileName);
      }
    }

    if (todosOsDados.notasFiscais.length > 0) {
      Logger.log(`Iniciando salvamento em lote de ${todosOsDados.notasFiscais.length} nova(s) nota(s) fiscal(is).`);
      ConciliacaoNFCrud_salvarDadosEmLote(todosOsDados);
      Logger.log(`Salvamento em lote finalizado.`);
    }

    let message = `Processamento concluído.\n\n- ${arquivosProcessados} nova(s) NF-e processada(s) com sucesso.\n- ${arquivosDuplicados} NF-e duplicada(s) ignorada(s).\n- ${arquivosComErro} arquivo(s) com erro de leitura/parsing.`;
    if (arquivosIgnorados.length > 0) {
      message += `\n\nArquivos ignorados (tipo não suportado ou erro de descompactação): ${arquivosIgnorados.join(', ')}.`;
    }
    
    return { success: true, message: message };

  } catch (e) {
    Logger.log(`ERRO GERAL em ConciliacaoNFController_uploadArquivos: ${e.toString()}\n${e.stack}`);
    return { success: false, message: `Erro fatal no servidor durante o processamento: ${e.message}` };
  } finally {
    lock.releaseLock();
    Logger.log('FINALIZANDO EXECUÇÃO DE UPLOAD. Lock liberado.');
  }
}


function ConciliacaoNFController_obterDadosParaPagina() {
  try {
    Logger.log("ConciliacaoNFController: Obtendo TODOS os dados para a página de conciliação.");
    
    const cotacoes = ConciliacaoNFCrud_obterCotacoesAbertas();
    const notasFiscais = ConciliacaoNFCrud_obterNFsNaoConciliadas();
    const mapeamentoConciliacao = ConciliacaoNFCrud_obterMapeamentoConciliacao(); // Busca os mapeamentos da aba "Conciliacao"

    if (cotacoes === null || notasFiscais === null) {
      throw new Error("Falha ao buscar dados das planilhas (Cotações ou Notas Fiscais).");
    }

    const chavesNFsNaoConciliadas = notasFiscais.map(nf => nf.chaveAcesso);
    const chavesCotacoesAbertas = cotacoes.map(c => ({ idCotacao: c.idCotacao, fornecedor: c.fornecedor }));

    const todosOsItensNF = ConciliacaoNFCrud_obterItensDasNFs(chavesNFsNaoConciliadas);
    const todosOsDadosGeraisNF = ConciliacaoNFCrud_obterDadosGeraisDasNFs(chavesNFsNaoConciliadas);
    const todosOsItensCotacao = ConciliacaoNFCrud_obterTodosItensCotacoesAbertas(chavesCotacoesAbertas);

    Logger.log(`Dados carregados: ${cotacoes.length} cotações, ${notasFiscais.length} NFs.`);
    
    return {
      success: true,
      dados: {
        cotacoes: cotacoes,
        notasFiscais: notasFiscais,
        itensNF: todosOsItensNF,
        itensCotacao: todosOsItensCotacao,
        dadosGeraisNF: todosOsDadosGeraisNF,
        mapeamentoConciliacao: mapeamentoConciliacao // Envia os mapeamentos para a interface
      },
      message: null
    };

  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_obterDadosParaPagina: ${e.toString()}\n${e.stack}`);
    return { success: false, dados: null, message: e.message };
  }
}

function ConciliacaoNFController_realizarComparacao(compositeKeyCotacao, chavesAcessoNF) {
  try {
    if (!compositeKeyCotacao || !chavesAcessoNF || chavesAcessoNF.length === 0) {
      throw new Error("Chave da Cotação e Chaves de Acesso das NFs são obrigatórios.");
    }
    
    const parts = compositeKeyCotacao.split('-');
    const idCotacao = parts.shift();
    const nomeFornecedor = parts.join('-');

    Logger.log(`Coletando dados para conciliação manual. Cotação ID: ${idCotacao}, Fornecedor: ${nomeFornecedor}`);
    
    const itensCotacao = ConciliacaoNFCrud_obterItensDaCotacao(idCotacao, nomeFornecedor);
    if (itensCotacao.length === 0) {
      return { success: false, message: "Nenhum item marcado para comprar nesta cotação para este fornecedor." };
    }

    const itensNF = ConciliacaoNFCrud_obterItensDasNFs(chavesAcessoNF);
    const dadosGeraisNF = ConciliacaoNFCrud_obterDadosGeraisDasNFs(chavesAcessoNF);
    const analisePrazo = {}; 

    const dadosParaPagina = {
      idCotacao: idCotacao,
      nomeFornecedor: nomeFornecedor,
      chavesAcessoNF: chavesAcessoNF,
      itensCotacao: itensCotacao,
      itensNF: itensNF,
      dadosGeraisNF: dadosGeraisNF,
      analisePrazo: analisePrazo
    };

    return { success: true, dados: dadosParaPagina };

  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_realizarComparacao: ${e.toString()}\n${e.stack}`);
    return { success: false, dados: null, message: e.message };
  }
}


function ConciliacaoNFController_salvarConciliacaoEmLote(dadosLote) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: 'Outro processo de salvamento já está em execução. Tente novamente em alguns instantes.' };
  }

  try {
    Logger.log("Recebidos dados para salvamento em lote.");
    const { conciliacoes, itensCortados, novosMapeamentos, statusUpdates } = dadosLote; 

    if (!Array.isArray(conciliacoes) || !Array.isArray(itensCortados) || !Array.isArray(novosMapeamentos) || !Array.isArray(statusUpdates)) {
      throw new Error("Formato de dados inválido para salvamento em lote.");
    }

    // [NOVO] Processa as atualizações de status (Sem Pedido, Bonificação, etc.)
    if (statusUpdates && statusUpdates.length > 0) {
      Logger.log(`Processando ${statusUpdates.length} atualização(ões) de status.`);
      // Agrupa as chaves de acesso por novo status para otimizar chamadas
      const updatesByStatus = statusUpdates.reduce((acc, update) => {
        if (!acc[update.novoStatus]) {
          acc[update.novoStatus] = [];
        }
        acc[update.novoStatus].push(update.chaveAcesso);
        return acc;
      }, {});

      for (const status in updatesByStatus) {
        ConciliacaoNFCrud_atualizarStatusNF(updatesByStatus[status], null, status);
        Logger.log(`Status de ${updatesByStatus[status].length} NF(s) atualizado para '${status}'.`);
      }
    }

    // Processa o resto do lote (conciliações, itens cortados, mapeamentos)
    const sucesso = ConciliacaoNFCrud_salvarAlteracoesEmLote(conciliacoes, itensCortados, novosMapeamentos);

    if (!sucesso) {
      throw new Error("Ocorreu uma falha no backend ao tentar salvar os dados nas planilhas.");
    }

    Logger.log("Salvamento em lote concluído com sucesso.");
    return { success: true, message: "Todas as alterações foram salvas com sucesso!" };

  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_salvarConciliacaoEmLote: ${e.toString()}\n${e.stack}`);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}