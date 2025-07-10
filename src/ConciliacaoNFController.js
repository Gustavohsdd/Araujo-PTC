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

/**
 * [VERSÃO DE DIAGNÓSTICO] Busca os dados de um fornecedor específico na aba 'Fornecedores'
 * usando o CNPJ, com logs detalhados para depuração.
 * @param {string} cnpj O CNPJ do fornecedor a ser buscado.
 * @returns {object|null} Um objeto com os dados do fornecedor se encontrado, ou null caso contrário.
 */
function ConciliacaoNFController_buscarFornecedorPorCnpj(cnpj) {
  Logger.log('--- INICIANDO BUSCA DE FORNECEDOR POR CNPJ ---');
  if (!cnpj) {
    Logger.log('AVISO: A função foi chamada com um CNPJ nulo ou vazio. Retornando null.');
    return null;
  }

  try {
    const cnpjNormalizado = String(cnpj).replace(/\D/g, '');
    Logger.log(`CNPJ recebido da NF: "${cnpj}" | Normalizado para busca: "${cnpjNormalizado}"`);

    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaFornecedores = planilha.getSheetByName(ABA_FORNECEDORES);
    if (!abaFornecedores) {
      throw new Error(`Aba de fornecedores '${ABA_FORNECEDORES}' não encontrada.`);
    }

    const dados = abaFornecedores.getDataRange().getValues();
    const cabecalhos = dados[0];
    const indiceCnpj = cabecalhos.indexOf("CNPJ");

    if (indiceCnpj === -1) {
      throw new Error("A coluna 'CNPJ' não foi encontrada na aba de Fornecedores.");
    }

    Logger.log('Iniciando varredura na planilha de Fornecedores...');
    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      const cnpjLinhaOriginal = linha[indiceCnpj];
      
      // Pula linhas em branco para não poluir o log
      if (!cnpjLinhaOriginal || String(cnpjLinhaOriginal).trim() === '') {
        continue;
      }

      const cnpjLinhaNormalizado = String(cnpjLinhaOriginal).replace(/\D/g, '');
      
      // Log para cada linha sendo verificada
      Logger.log(`Linha ${i + 1}: Lendo CNPJ da planilha: "${cnpjLinhaOriginal}" | Normalizado: "${cnpjLinhaNormalizado}"`);

      if (cnpjLinhaNormalizado === cnpjNormalizado) {
        Logger.log(`>>> SUCESSO! Fornecedor encontrado na linha ${i + 1}. Preparando objeto de retorno.`);
        const fornecedorObj = {};
        cabecalhos.forEach((cabecalho, index) => {
          // Tratamento para datas, para evitar erros de serialização
          if (linha[index] instanceof Date) {
             fornecedorObj[cabecalho] = linha[index].toLocaleDateString('pt-BR');
          } else {
             fornecedorObj[cabecalho] = linha[index];
          }
        });
        Logger.log('--- FIM DA BUSCA (ENCONTRADO) ---');
        return fornecedorObj;
      }
    }

    Logger.log('AVISO: Nenhum fornecedor correspondente encontrado após varrer toda a planilha.');
    Logger.log('--- FIM DA BUSCA (NÃO ENCONTRADO) ---');
    return null; // Retorna null se não encontrar o fornecedor

  } catch (e) {
    Logger.log(`ERRO CRÍTICO em ConciliacaoNFController_buscarFornecedorPorCnpj: ${e.message}\nStack: ${e.stack}`);
    throw new Error(`Falha ao buscar fornecedor por CNPJ: ${e.message}`);
  }
}


/**
 * [CORRIGIDO] Recebe os dados do fornecedor do modal da NF, e chama o controller de Fornecedores
 * para criar um novo ou atualizar um existente. Garante que um fornecedor existente seja sempre
 * atualizado, mesmo que o ID não venha do frontend.
 * @param {object} dadosFornecedor Objeto com os dados do fornecedor vindos do formulário.
 * @returns {object} O resultado da operação de criação ou atualização.
 */
function ConciliacaoNFController_salvarFornecedorViaNF(dadosFornecedor) {
  try {
    if (!dadosFornecedor) {
      throw new Error("Nenhum dado de fornecedor foi recebido.");
    }

    // --- LÓGICA DE DECISÃO MELHORADA ---

    // 1. Se um ID for fornecido explicitamente, é uma atualização.
    if (dadosFornecedor.ID && String(dadosFornecedor.ID).trim().length > 0) {
      Logger.log(`ID ${dadosFornecedor.ID} fornecido. Redirecionando para ATUALIZAR fornecedor.`);
      // A função FornecedoresController_atualizarFornecedor já propaga a mudança de nome.
      return FornecedoresController_atualizarFornecedor(dadosFornecedor);
    } 
    
    // 2. Se não houver ID, faz uma verificação defensiva no backend usando o CNPJ.
    if (dadosFornecedor.CNPJ) {
      const fornecedorExistente = ConciliacaoNFController_buscarFornecedorPorCnpj(dadosFornecedor.CNPJ);
      
      // Se encontrou um fornecedor com o mesmo CNPJ, deve ser uma ATUALIZAÇÃO.
      if (fornecedorExistente && fornecedorExistente.ID) {
        Logger.log(`ID não fornecido, mas CNPJ ${dadosFornecedor.CNPJ} já existe com ID ${fornecedorExistente.ID}. Redirecionando para ATUALIZAR.`);
        // Adiciona o ID encontrado ao objeto de dados para garantir que a função de atualização seja chamada corretamente.
        dadosFornecedor.ID = fornecedorExistente.ID;
        // A função FornecedoresController_atualizarFornecedor já propaga a mudança de nome.
        return FornecedoresController_atualizarFornecedor(dadosFornecedor);
      }
    }

    // 3. Se não encontrou ID nem CNPJ existente, então é uma CRIAÇÃO.
    Logger.log(`Nenhum ID ou CNPJ existente encontrado. Redirecionando para CRIAR novo fornecedor: ${dadosFornecedor.Fornecedor}`);
    // A função FornecedoresController_criarNovoFornecedor já previne duplicatas por nome ou CNPJ.
    return FornecedoresController_criarNovoFornecedor(dadosFornecedor);
    
  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_salvarFornecedorViaNF: ${e.toString()}\n${e.stack}`);
    return { success: false, message: `Erro no servidor ao salvar fornecedor: ${e.message}` };
  }
}

/**
 * Busca os dados necessários para popular o modal de cadastro de itens via NF.
 * Retorna os itens da NF que não possuem conciliação, o fornecedor da NF e a lista de todos os Produtos Principais.
 * @param {string} chaveAcesso A chave de acesso da NF selecionada.
 * @returns {object} Objeto com { success: boolean, dados?: { itensNF: Array, fornecedor: object, produtos: Array } }
 */
function ConciliacaoNFController_obterDadosParaCadastroItens(chaveAcesso) {
  try {
    if (!chaveAcesso) {
      throw new Error("A chave de acesso da NF é necessária.");
    }

    // 1. Obter todos os itens da NF
    const todosItensDaNF = ConciliacaoNFCrud_obterItensDasNFs([chaveAcesso]);

    // 2. Obter o Fornecedor (ID e Nome)
    const nfGeral = ConciliacaoNFCrud_obterNFsNaoConciliadas().find(nf => nf.chaveAcesso === chaveAcesso);
    if (!nfGeral) {
        throw new Error("Não foi possível encontrar os dados gerais da NF selecionada.");
    }
    const fornecedorDaNF = ConciliacaoNFController_buscarFornecedorPorCnpj(nfGeral.cnpjEmitente);
    if (!fornecedorDaNF || !fornecedorDaNF.ID) {
      throw new Error(`O fornecedor com CNPJ '${nfGeral.cnpjEmitente}' não está cadastrado. Cadastre o fornecedor primeiro.`);
    }

    // 3. Obter todos os Produtos principais para o dropdown
    const todosOsProdutos = SubProdutosCRUD_obterTodosProdutosParaDropdown();

    // NOTA: No futuro, podemos otimizar para retornar apenas os itens não conciliados.
    // Por enquanto, retornaremos todos para o frontend filtrar.
    return {
      success: true,
      dados: {
        itensNF: todosItensDaNF,
        fornecedor: {
          ID: fornecedorDaNF.ID,
          Fornecedor: fornecedorDaNF.Fornecedor
        },
        produtos: todosOsProdutos
      }
    };
  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_obterDadosParaCadastroItens: ${e.toString()}\n${e.stack}`);
    return { success: false, message: e.message };
  }
}

/**
 * [VERSÃO FINAL E CORRIGIDA] Cria um novo Produto Principal.
 * Esta função chama a função CRUD compartilhada e, em caso de sucesso,
 * TRANSFORMA a resposta { success: true, novoId: '...' } no formato que a
 * interface espera { success: true, produto: { ID: '...', Produto: '...' } },
 * resolvendo o problema de atualização da tela sem modificar a função CRUD.
 * @param {object} dadosProduto Objeto com os dados do novo produto (Nome, UN, etc.).
 * @returns {object} Objeto com o resultado da operação formatado para a tela de conciliação.
 */
function ConciliacaoNFController_salvarProdutoViaNF(dadosProduto) {
  try {
    Logger.log("ConciliacaoNFController: redirecionando para criar novo Produto com dados:", JSON.stringify(dadosProduto));
    
    // 1. Chama a sua função CRUD, que não será modificada.
    const resultadoCRUD = ProdutosCRUD_criarNovoProduto(dadosProduto);
    
    // 2. Se a função CRUD teve sucesso...
    if (resultadoCRUD && resultadoCRUD.success) {
      Logger.log("CRUD retornou sucesso. Transformando a resposta para a interface.");

      // 3. Montamos o objeto 'produto' que a interface precisa,
      // pegando o ID da resposta do CRUD e o Nome dos dados que já tínhamos.
      const produtoParaCliente = {
        ID: resultadoCRUD.novoId,
        Produto: dadosProduto.Produto 
      };

      // 4. Retornamos o objeto completo no formato correto.
      return { success: true, produto: produtoParaCliente };

    } else {
       // Se o CRUD falhou, apenas repassamos a mensagem de erro.
       Logger.log("ERRO: O CRUD retornou uma falha.");
       return resultadoCRUD;
    }

  } catch (e) {
    Logger.log(`ERRO CRÍTICO em ConciliacaoNFController_salvarProdutoViaNF: ${e.toString()}\n${e.stack}`);
    return { success: false, message: "Erro crítico no servidor ao criar produto: " + e.message };
  }
}

/**
 * Cadastra múltiplos SubProdutos em lote a partir do modal da tela de conciliação.
 * Reutiliza a função CRUD existente do módulo de SubProdutos.
 * @param {object} dadosLote Objeto contendo { fornecedorId, subProdutos: [...] }.
 * @returns {object} Objeto com o resultado da operação em lote.
 */
function ConciliacaoNFController_salvarSubProdutosViaNF(dadosLote) {
  try {
    Logger.log("ConciliacaoNFController: redirecionando para cadastrar múltiplos SubProdutos.");
    
    // A função SubProdutosCRUD_cadastrarMultiplosSubProdutos é ideal, pois já lida com a busca
    // de nomes a partir de IDs e validações de duplicidade.
    // Apenas ajustamos o payload para corresponder ao que a função espera.
    const payloadParaCRUD = {
      fornecedorGlobal: dadosLote.fornecedorId,
      subProdutos: dadosLote.subProdutos.map(sp => ({
        "ProdutoVinculadoID": sp["Produto Vinculado"], // Passa o ID do produto
        "SubProduto": sp.SubProduto,
        "UN": sp.UN,
        "Categoria": sp.Categoria,
        "Tamanho": sp.Tamanho,
        "Fator": sp.Fator,
        "NCM": sp.NCM,
        "CST": sp.CST,
        "CFOP": sp.CFOP,
        "Status": sp.Status || "Ativo"
      }))
    };
    
    return SubProdutosCRUD_cadastrarMultiplosSubProdutos(payloadParaCRUD);

  } catch (e) {
    Logger.log(`ERRO em ConciliacaoNFController_salvarSubProdutosViaNF: ${e.toString()}\n${e.stack}`);
    return { success: false, message: "Erro no servidor ao salvar subprodutos: " + e.message };
  }
}