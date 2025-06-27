/**
 * @file ConciliacaoNFController.gs
 * @description Orquestra o processo de leitura dos arquivos XML de NF-e da pasta do Drive,
 * o upload de novos arquivos e chama as funções CRUD para popular a planilha de dados.
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
    let arquivosExcluidos = 0;

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

        // ===== ALTERAÇÃO PRINCIPAL: EXCLUIR ARQUIVO DUPLICADO =====
        if (chavesExistentes.has(chaveAtual)) {
          Logger.log(`AVISO: A NF-e com chave ${chaveAtual} já existe. Excluindo arquivo duplicado: ${nomeArquivo}.`);
          arquivo.setTrashed(true); // Move o arquivo para a lixeira
          arquivosExcluidos++;
          continue;
        }
        
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
        arquivo.moveTo(pastaProcessados); // Move para processados mesmo com erro para não travar o processo
        arquivosSkipped++;
      }
    }
    
    if (todosOsDados.notasFiscais.length > 0) {
        Logger.log(`Iniciando salvamento em lote de ${todosOsDados.notasFiscais.length} nota(s) fiscal(is).`);
        ConciliacaoNFCrud_salvarDadosEmLote(todosOsDados);
        Logger.log(`Salvamento em lote finalizado.`);
    } else {
        Logger.log("Nenhuma nova nota fiscal para adicionar à planilha.");
    }
    
    Logger.log(`--- Processamento finalizado. Resumo: ${arquivosProcessados} processados, ${arquivosExcluidos} duplicados excluídos, ${arquivosSkipped} pulados/com erro. ---`);

  } catch (error) {
    Logger.log(`ERRO GERAL na função ConciliacaoNFController_processarXmlsDaPasta: ${error.toString()}. Stack: ${error.stack}`);
  } finally {
    lock.releaseLock();
    Logger.log('FINALIZANDO EXECUÇÃO. Lock liberado.');
  }
}

/**
 * [NOVA FUNÇÃO]
 * Recebe arquivos da interface, decodifica e salva na pasta de XMLs.
 * Suporta arquivos XML e ZIP. Arquivos RAR não são suportados.
 * @param {Array<object>} arquivos - Array de objetos com {fileName, mimeType, content (base64)}
 * @returns {object} Objeto com { success: boolean, message: string }.
 */
function ConciliacaoNFController_uploadArquivos(arquivos) {
  if (!arquivos || arquivos.length === 0) {
    return { success: false, message: "Nenhum arquivo recebido." };
  }
  
  try {
    const pastaXml = DriveApp.getFolderById(ID_PASTA_XML);
    let arquivosSalvos = 0;
    let arquivosXmlExtraidos = 0;
    let arquivosIgnorados = [];

    for (const arquivo of arquivos) {
      const { fileName, mimeType, content } = arquivo;
      const decodedContent = Utilities.base64Decode(content);
      const blob = Utilities.newBlob(decodedContent, mimeType, fileName);

      const nomeArquivoLower = fileName.toLowerCase();

      // Handle ZIP files
      if (mimeType === 'application/zip' || mimeType === 'application/x-zip-compressed' || nomeArquivoLower.endsWith('.zip')) {
        try {
          const arquivosDescompactados = Utilities.unzip(blob);
          for (const arquivoDescompactado of arquivosDescompactados) {
            if (arquivoDescompactado.getName().toLowerCase().endsWith('.xml')) {
              pastaXml.createFile(arquivoDescompactado);
              arquivosXmlExtraidos++;
            }
          }
           Logger.log(`Arquivo ZIP "${fileName}" processado.`);
        } catch(e) {
           Logger.log(`Erro ao descompactar o arquivo ZIP "${fileName}": ${e.message}`);
           arquivosIgnorados.push(fileName);
        }
      }
      // Handle individual XML files
      else if (mimeType === 'text/xml' || mimeType === 'application/xml' || nomeArquivoLower.endsWith('.xml')) {
        pastaXml.createFile(blob);
        arquivosSalvos++;
        Logger.log(`Arquivo XML "${fileName}" salvo.`);
      }
      // Log unsupported types like RAR
      else {
        Logger.log(`Arquivo ${fileName} com tipo ${mimeType} não é suportado e foi ignorado.`);
        arquivosIgnorados.push(fileName);
      }
    }

    let message = '';
    if (arquivosSalvos > 0) message += `${arquivosSalvos} arquivo(s) XML salvo(s) com sucesso.\n`;
    if (arquivosXmlExtraidos > 0) message += `${arquivosXmlExtraidos} arquivo(s) XML extraído(s) de arquivos ZIP.\n`;
    
    if (message === '') {
      message = 'Nenhum arquivo XML válido foi encontrado para upload. Lembre-se que arquivos .RAR não são suportados.';
    } else {
      if (arquivosIgnorados.length > 0) {
        message += `\nOs seguintes arquivos foram ignorados (tipo não suportado): ${arquivosIgnorados.join(', ')}.`;
      }
    }
    
    return { success: true, message: message };

  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_uploadArquivos: ${e.toString()}\n${e.stack}`);
    return { success: false, message: `Erro no servidor durante o upload: ${e.message}` };
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
    // Extrai os novos mapeamentos do payload recebido da interface
    const { conciliacoes, itensCortados, novosMapeamentos } = dadosLote; 

    if (!Array.isArray(conciliacoes) || !Array.isArray(itensCortados) || !Array.isArray(novosMapeamentos)) {
      throw new Error("Formato de dados inválido para salvamento em lote.");
    }

    // Passa os novos mapeamentos para a função de CRUD que salva nas planilhas
    const sucesso = ConciliacaoNFCrud_salvarAlteracoesEmLote(conciliacoes, itensCortados, novosMapeamentos);

    if (!sucesso) {
      throw new Error("Ocorreu uma falha no backend ao tentar salvar os dados nas planilhas.");
    }

    Logger.log("Salvamento em lote concluído com sucesso.");
    return { success: true, message: "Todas as conciliações e itens cortados foram salvos com sucesso!" };

  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_salvarConciliacaoEmLote: ${e.toString()}\n${e.stack}`);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function ConciliacaoNFController_marcarNFsSemPedido(chavesAcessoNF) {
  try {
    if (!chavesAcessoNF || !Array.isArray(chavesAcessoNF) || chavesAcessoNF.length === 0) {
      throw new Error("Nenhuma chave de acesso foi fornecida.");
    }
    
    Logger.log(`Marcando ${chavesAcessoNF.length} NF(s) como 'Sem Pedido'.`);

    // A função ConciliacaoNFCrud_atualizarStatusNF foi adicionada ao CRUD para esta chamada funcionar
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
