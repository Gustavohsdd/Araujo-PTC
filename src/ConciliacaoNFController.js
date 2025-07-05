/**
 * @file ConciliacaoNFController.gs
 * @description [UNIFICADO] Orquestra o processo de conciliação e rateio, 
 * buscando e salvando dados para ambos os módulos.
 */

/**
 * Recebe arquivos da interface, decodifica, processa em memória e salva na planilha.
 * Nenhuma alteração necessária nesta função.
 * @param {Array<object>} arquivos - Array de objetos com {fileName, content (base64)}
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

  Logger.log('INICIANDO EXECUÇÃO DE UPLOAD...');

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
  
  try {
    const chavesExistentes = ConciliacaoNFCrud_obterChavesDeAcessoExistentes();
    Logger.log(`Encontradas ${chavesExistentes.size} chaves existentes.`);

    for (const arquivo of arquivos) {
      const { fileName, content } = arquivo;
      try {
        const decodedContent = Utilities.base64Decode(content);
        const blob = Utilities.newBlob(decodedContent, 'application/xml', fileName);
        const conteudoXml = blob.getDataAsString('UTF-8');
        const dadosNf = ConciliacaoNFCrud_parsearConteudoXml(conteudoXml);
        
        const chaveAtual = dadosNf.notasFiscais.chaveAcesso;
        if (chavesExistentes.has(chaveAtual)) {
          Logger.log(`AVISO: NF ${chaveAtual} já existe. Pulando.`);
          arquivosDuplicados++;
          continue;
        }

        todosOsDados.notasFiscais.push(dadosNf.notasFiscais);
        todosOsDados.itensNf.push(...dadosNf.itensNf);
        todosOsDados.faturasNf.push(...dadosNf.faturasNf);
        todosOsDados.transporteNf.push(...dadosNf.transporteNf);
        todosOsDados.tributosTotaisNf.push(...dadosNf.tributosTotaisNf);

        chavesExistentes.add(chaveAtual);
        arquivosProcessados++;
      } catch (e) {
         Logger.log(`ERRO ao processar o arquivo ${fileName}. Erro: ${e.message}`);
         arquivosComErro++;
      }
    }

    if (todosOsDados.notasFiscais.length > 0) {
      ConciliacaoNFCrud_salvarDadosEmLote(todosOsDados);
    }

    let message = `Processamento concluído.\n- ${arquivosProcessados} nova(s) NF-e processada(s).\n- ${arquivosDuplicados} NF-e duplicada(s) ignorada(s).\n- ${arquivosComErro} arquivo(s) com erro.`;
    return { success: true, message: message };

  } catch (e) {
    Logger.log(`ERRO GERAL em ConciliacaoNFController_uploadArquivos: ${e.toString()}\n${e.stack}`);
    return { success: false, message: `Erro fatal no servidor: ${e.message}` };
  } finally {
    lock.releaseLock();
    Logger.log('FINALIZANDO EXECUÇÃO DE UPLOAD. Lock liberado.');
  }
}


/**
 * Obtém todos os dados para a página unificada.
 */
function ConciliacaoNFController_obterDadosParaPagina() {
  try {
    Logger.log("Controller Unificado: Obtendo TODOS os dados para a página.");
    
    // Dados da Conciliação
    const cotacoes = ConciliacaoNFCrud_obterCotacoesAbertas();
    const notasFiscais = ConciliacaoNFCrud_obterNFsNaoConciliadas();
    const mapeamentoConciliacao = ConciliacaoNFCrud_obterMapeamentoConciliacao();
    const chavesNFsNaoConciliadas = notasFiscais.map(nf => nf.chaveAcesso);
    const chavesCotacoesAbertas = cotacoes.map(c => ({ idCotacao: c.idCotacao, fornecedor: c.fornecedor }));
    const todosOsItensNF = ConciliacaoNFCrud_obterItensDasNFs(chavesNFsNaoConciliadas);
    const todosOsDadosGeraisNF = ConciliacaoNFCrud_obterDadosGeraisDasNFs(chavesNFsNaoConciliadas); 
    const todosOsItensCotacao = ConciliacaoNFCrud_obterTodosItensCotacoesAbertas(chavesCotacoesAbertas);
    
    // Dados do Rateio
    const regrasRateio = RateioCrud_obterRegrasRateio();
    const setoresUnicos = RateioCrud_obterSetoresUnicos(); // ADICIONADO

    Logger.log(`Dados carregados: ${cotacoes.length} cotações, ${notasFiscais.length} NFs, ${regrasRateio.length} regras, ${setoresUnicos.length} setores.`);
    
    return {
      success: true,
      dados: {
        cotacoes,
        notasFiscais,
        itensNF: todosOsItensNF,
        itensCotacao: todosOsItensCotacao,
        dadosGeraisNF: todosOsDadosGeraisNF,
        mapeamentoConciliacao,
        regrasRateio,
        setoresUnicos // ADICIONADO
      }
    };

  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_obterDadosParaPagina: ${e.toString()}\n${e.stack}`);
    return { success: false, message: e.message };
  }
}


/**
 * [ALTERADO] Salva um lote unificado de alterações de conciliação e rateio.
 * Esta função substitui a antiga 'salvarConciliacaoEmLote' e a 'salvarRateioEmLote'.
 * A lógica de rateio agora inclui um tratamento para pagamentos "À Vista" (sem faturas).
 * @param {object} dadosLote - O objeto contendo todos os dados a serem salvos.
 */
function ConciliacaoNFController_salvarLoteUnificado(dadosLote) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: 'Outro processo de salvamento já está em execução.' };
  }

  try {
    Logger.log("Iniciando salvamento de lote unificado (Conciliação + Rateio).");
    const { conciliacoes, itensCortados, novosMapeamentos, statusUpdates, rateios } = dadosLote;

    // 1. Salvar alterações da Conciliação
    if (conciliacoes && conciliacoes.length > 0) {
        ConciliacaoNFCrud_salvarAlteracoesEmLote(conciliacoes, itensCortados, novosMapeamentos);
        Logger.log(`${conciliacoes.length} conciliação(ões) salva(s).`);
    }

    // 2. Salvar atualizações de Status (Ex: Sem Pedido, Bonificação, NF Tipo B)
    if (statusUpdates && statusUpdates.length > 0) {
      const updatesByStatus = statusUpdates.reduce((acc, update) => {
        if (!acc[update.novoStatus]) acc[update.novoStatus] = [];
        acc[update.novoStatus].push(update.chaveAcesso);
        return acc;
      }, {});
      for (const status in updatesByStatus) {
        ConciliacaoNFCrud_atualizarStatusNF(updatesByStatus[status], null, status);
        Logger.log(`Status de ${updatesByStatus[status].length} NF(s) atualizado para '${status}'.`);
      }
    }
    
    // 3. Salvar dados do Rateio
    if (rateios && rateios.length > 0) {
        const todasAsLinhasContasAPagar = [];
        const todasAsNovasRegras = [];
        const todasAsChavesParaAtualizar = new Set();
        
        for (const dadosRateio of rateios) {
            const { chaveAcesso, totaisPorSetor, novasRegras, numeroNF, mapaSetorParaItens } = dadosRateio;
            todasAsChavesParaAtualizar.add(chaveAcesso);

            if (novasRegras && novasRegras.length > 0) {
                todasAsNovasRegras.push(...novasRegras);
            }
            
            const faturas = RateioCrud_obterFaturasDaNF(chaveAcesso);
            let valorTotalRateadoNota = Object.values(totaisPorSetor).reduce((s, v) => s + v, 0);

            // [LÓGICA ALTERADA] Trata faturas existentes ou cria lançamento "À Vista"
            if (faturas && faturas.length > 0) {
                // Lógica original para notas com faturas
                const numFaturasOriginais = faturas.length;
                const numSetores = Object.keys(totaisPorSetor).length;
                const totalNovosTitulosNota = numFaturasOriginais * numSetores;
                let contadorParcelaNota = 1;

                faturas.forEach(fatura => {
                    for (const setor in totaisPorSetor) {
                        const resumoItens = mapaSetorParaItens[setor] ? mapaSetorParaItens[setor].join(', ') : `NF ${numeroNF}`;
                        const numeroParcelaFormatado = `${contadorParcelaNota++}/${totalNovosTitulosNota}(Ref: ${fatura.numeroParcela})`;
                        
                        todasAsLinhasContasAPagar.push({
                            'ChavedeAcesso': chaveAcesso,
                            'NúmerodaFatura': fatura.numeroFatura,
                            'NúmerodaParcela': numeroParcelaFormatado,
                            'ResumodosItens': resumoItens,
                            'DatadeVencimento': new Date(fatura.dataVencimento),
                            'ValordaParcela': fatura.valorParcela,
                            'Setor': setor,
                            'ValorporSetor': (totaisPorSetor[setor] / valorTotalRateadoNota) * fatura.valorParcela
                        });
                    }
                });
            } else { 
                // [NOVO] Lógica para pagamento à vista (sem faturas explícitas)
                const numSetores = Object.keys(totaisPorSetor).length;
                let contadorParcelaNota = 1;

                for (const setor in totaisPorSetor) {
                    const resumoItens = mapaSetorParaItens[setor] ? mapaSetorParaItens[setor].join(', ') : `NF ${numeroNF}`;
                    const numeroParcelaFormatado = `${contadorParcelaNota++}/${numSetores}(À Vista)`;
                    
                    todasAsLinhasContasAPagar.push({
                        'ChavedeAcesso': chaveAcesso,
                        'NúmerodaFatura': numeroNF, // Usa o número da NF como referência
                        'NúmerodaParcela': numeroParcelaFormatado,
                        'ResumodosItens': resumoItens,
                        'DatadeVencimento': new Date(), // Data de hoje como vencimento
                        'ValordaParcela': valorTotalRateadoNota, // O valor da "parcela única" é o total da nota
                        'Setor': setor,
                        'ValorporSetor': totaisPorSetor[setor] // O valor por setor já é o valor final
                    });
                }
            }
        }
        
        RateioCrud_salvarNovasRegrasDeRateio(todasAsNovasRegras);
        RateioCrud_salvarContasAPagar(todasAsLinhasContasAPagar);
        Array.from(todasAsChavesParaAtualizar).forEach(chave => {
            // O status para rateio sem pedido já é setado no payload vindo do frontend,
            // aqui apenas garantimos que o status do MÓDULO DE RATEIO seja concluído.
            RateioCrud_atualizarStatusRateio(chave, "Concluído");
        });
        Logger.log(`${rateios.length} rateio(s) salvos e status atualizados.`);
    }

    Logger.log("Salvamento em lote unificado concluído com sucesso.");
    return { success: true, message: "Todas as alterações foram salvas com sucesso!" };

  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_salvarLoteUnificado: ${e.toString()}\n${e.stack}`);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Ponto de entrada para a interface buscar os dados para o Relatório de Rateio.
 * @param {string[]} termosDeBusca - Um array de números de NF ou Chaves de Acesso.
 * @returns {object} Objeto com status e os dados do relatório.
 */
function ConciliacaoNFController_obterDadosParaRelatorio(termosDeBusca) {
  try {
    const dados = RateioCrud_obterDadosParaRelatorio(termosDeBusca);
    return { success: true, dados: dados };
  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_obterDadosParaRelatorio: ${e.toString()}\n${e.stack}`);
    return { success: false, message: e.message };
  }
}

/**
 * Ponto de entrada para a interface buscar os dados para o Relatório de Rateio SINTÉTICO.
 * @param {string[]} termosDeBusca - Um array de números de NF ou Chaves de Acesso.
 * @returns {object} Objeto com status e os dados do relatório.
 */
function ConciliacaoNFController_obterDadosParaRelatorioSintetico(termosDeBusca) {
  try {
    const dados = RateioCrud_obterDadosParaRelatorioSintetico(termosDeBusca);
    return { success: true, dados: dados };
  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_obterDadosParaRelatorioSintetico: ${e.toString()}\n${e.stack}`);
    return { success: false, message: e.message };
  }
}