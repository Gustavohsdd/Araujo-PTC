// @ts-nocheck

/**
 * @file CotacoesController.gs
 * @description Controlador para as operações relacionadas a cotações.
 */

/**
 * Obtém os resumos de cotações da planilha.
 * Cada resumo representa uma cotação única com informações agregadas.
 *
 * @return {object} Um objeto contendo:
 * - success {boolean}: true se a operação foi bem-sucedida, false caso contrário.
 * - dados {Array<object>|null}: Um array de objetos de resumo de cotação, ou null em caso de falha.
 * - message {string|null}: Uma mensagem de erro, se success for false.
 */
function CotacoesController_obterResumosDeCotacoes() {
  console.log("CotacoesController_obterResumosDeCotacoes: Iniciando execução.");
  try {
    console.log("CotacoesController: Prestes a chamar CotacoesCRUD_obterResumosDeCotacoes.");
    const resumosDeCotacoes = CotacoesCRUD_obterResumosDeCotacoes(); 

    if (resumosDeCotacoes === null) {
      console.warn("CotacoesController: CotacoesCRUD_obterResumosDeCotacoes retornou null. Falha no CRUD.");
      const resultadoFalha = {
        success: false,
        dados: null,
        message: "Não foi possível obter os resumos das cotações. Verifique os logs do servidor (CRUD)."
      };
      return resultadoFalha;
    }

    console.log(`CotacoesController: CotacoesCRUD_obterResumosDeCotacoes retornou com sucesso. Número de resumos de cotação: ${resumosDeCotacoes.length}.`);
    
    const resultadoSucesso = {
      success: true,
      dados: resumosDeCotacoes, 
      message: null
    };
    return resultadoSucesso;

  } catch (error) {
    console.error("!!!!!!!! ERRO CAPTURADO em CotacoesController_obterResumosDeCotacoes !!!!!!!!");
    console.error("Mensagem do Erro: " + error.toString());
    console.error("Stack Trace do Erro: " + error.stack);
    const resultadoErro = {
      success: false,
      dados: null,
      message: "Erro geral no controlador ao obter resumos das cotações: " + error.message
    };
    return resultadoErro;
  }
}


/**
 * Obtém as opções necessárias para criar uma nova cotação (listas de categorias, fornecedores, produtos).
 * @return {object} Um objeto com { success: boolean, dados: {categorias, fornecedores, produtos}|null, message: string|null }.
 */
function CotacoesController_obterOpcoesNovaCotacao() {
  console.log("CotacoesController_obterOpcoesNovaCotacao: Iniciando.");
  try {
    const categorias = CotacoesCRUD_obterListaCategoriasProdutos();
    const fornecedores = CotacoesCRUD_obterListaFornecedores(); // Retorna {id, nome}
    const produtos = CotacoesCRUD_obterListaProdutos();       // Retorna {id, nome}

    if (categorias === null || fornecedores === null || produtos === null) {
        return {
            success: false,
            dados: null,
            message: "Falha ao obter uma ou mais listas de opções do CRUD."
        };
    }
    
    return {
      success: true,
      dados: {
        categorias: categorias,
        fornecedores: fornecedores,
        produtos: produtos
      },
      message: null
    };
  } catch (e) {
    console.error("Erro em CotacoesController_obterOpcoesNovaCotacao: " + e.toString(), e.stack);
    return {
      success: false,
      dados: null,
      message: "Erro no servidor ao buscar opções para nova cotação: " + e.message
    };
  }
}

/**
 * Cria uma nova cotação com base nas opções fornecidas pelo cliente.
 * @param {object} opcoesCriacao Objeto contendo o 'tipo' de criação (categoria, fornecedor, etc.) 
 * e as 'selecoes' (array de IDs ou valores selecionados).
 * Ex: { tipo: 'categoria', selecoes: ['CAT-01', 'CAT-02'] }
 * @return {object} Um objeto com { success: boolean, idCotacao: string|null, numItens: int|null, message: string|null }.
 */
function CotacoesController_criarNovaCotacao(opcoesCriacao) {
  console.log("CotacoesController_criarNovaCotacao: Iniciando com opções:", JSON.stringify(opcoesCriacao));
  if (!opcoesCriacao || !opcoesCriacao.tipo || !opcoesCriacao.selecoes) {
    return { success: false, message: "Opções de criação inválidas ou incompletas." };
  }

  try {
    const resultadoCRUD = CotacoesCRUD_criarNovaCotacao(opcoesCriacao);
    
    if (resultadoCRUD && resultadoCRUD.success) {
      console.log(`CotacoesController: Nova cotação criada com sucesso: ID ${resultadoCRUD.idCotacao}, Itens: ${resultadoCRUD.numItens}`);
      return {
        success: true,
        idCotacao: resultadoCRUD.idCotacao,
        numItens: resultadoCRUD.numItens,
        message: "Nova cotação criada com sucesso."
      };
    } else {
      console.warn("CotacoesController: Falha ao criar nova cotação no CRUD.", resultadoCRUD ? resultadoCRUD.message : "Resultado CRUD nulo");
      return {
        success: false,
        message: resultadoCRUD ? resultadoCRUD.message : "Erro desconhecido ao criar cotação no CRUD."
      };
    }
  } catch (error) {
    console.error("!!!!!!!! ERRO CAPTURADO em CotacoesController_criarNovaCotacao !!!!!!!!");
    console.error("Mensagem do Erro: " + error.toString());
    console.error("Stack Trace do Erro: " + error.stack);
    return {
      success: false,
      dados: null,
      message: "Erro geral no controlador ao criar nova cotação: " + error.message
    };
  }
}

/**
 * Função interna para verificar se uma cotação está completa e, se estiver,
 * alterar seu status para "Finalizado".
 * Uma cotação é considerada completa quando todos os subprodutos marcados para
 * compra (coluna 'Comprar' > 0) têm um status de subproduto preenchido.
 * @param {string|number} idCotacao O ID da cotação a ser verificada.
 */
function _verificarEFinalizarCotacaoSeCompleta(idCotacao) {
  try {
    Logger.log(`Iniciando verificação de finalização para cotação ID: ${idCotacao}`);

    const abaCotacoes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_COTACOES);
    const ultimaLinha = abaCotacoes.getLastRow();
    if (ultimaLinha <= 1) return;

    const range = abaCotacoes.getRange(2, 1, ultimaLinha - 1, abaCotacoes.getLastColumn());
    const todosOsValores = range.getValues();
    const cabecalhos = Utilities_obterCabecalhos(ABA_COTACOES);

    const indiceIdCotacao = cabecalhos.indexOf("ID da Cotação");
    const indiceStatusCotacao = cabecalhos.indexOf("Status da Cotação");
    const indiceComprar = cabecalhos.indexOf("Comprar");
    const indiceStatusSubproduto = cabecalhos.indexOf("Status do SubProduto");

    // Filtra apenas as linhas da cotação relevante
    const linhasDaCotacao = todosOsValores.filter(linha => linha[indiceIdCotacao] == idCotacao);

    if (linhasDaCotacao.length === 0) {
      Logger.log(`Nenhuma linha encontrada para a cotação ${idCotacao}.`);
      return;
    }
    
    // Se a cotação já está Finalizada ou Cancelada, não faz nada.
    const statusAtual = linhasDaCotacao[0][indiceStatusCotacao];
    if (statusAtual === "Finalizado" || statusAtual === "Cancelado") {
      Logger.log(`Cotação ${idCotacao} já está em um estado final. Abortando.`);
      return;
    }

    // Pega todos os itens que deveriam ser comprados (Comprar > 0)
    const itensParaComprar = linhasDaCotacao.filter(linha => {
      const comprarValor = parseFloat(linha[indiceComprar]);
      return !isNaN(comprarValor) && comprarValor > 0;
    });

    // Se não há itens marcados para comprar, não há o que finalizar.
    if (itensParaComprar.length === 0) {
      Logger.log(`Cotação ${idCotacao} não possui itens marcados para comprar.`);
      return;
    }
    
    // Verifica se TODOS os itens para comprar já têm um status de subproduto
    const todosItensMarcados = itensParaComprar.every(linha => {
      const statusSub = linha[indiceStatusSubproduto];
      return statusSub !== "" && statusSub !== null && statusSub !== undefined;
    });

    if (todosItensMarcados) {
      Logger.log(`Todos os itens da cotação ${idCotacao} foram marcados. Atualizando status para "Finalizado".`);
      // Se todos estão marcados, atualiza o status principal da cotação para "Finalizado"
      // Itera sobre todas as linhas da planilha original para encontrar e marcar as linhas corretas.
      todosOsValores.forEach((linha, index) => {
        if (linha[indiceIdCotacao] == idCotacao) {
          // Atualiza o valor na matriz de dados
          todosOsValores[index][indiceStatusCotacao] = "Finalizado";
        }
      });
      
      // Escreve a matriz de dados inteira de volta na planilha
      range.setValues(todosOsValores);
      SpreadsheetApp.flush();
    } else {
      Logger.log(`Cotação ${idCotacao} ainda tem itens pendentes de marcação.`);
    }

  } catch (e) {
    Logger.log(`ERRO em _verificarEFinalizarCotacaoSeCompleta para ID ${idCotacao}: ${e.toString()}`);
  }
}

/**
 * Função de automação para cancelar cotações com mais de 3 meses de abertura
 * e que não possuam nenhum subproduto com status "Cancelado".
 * Esta função foi projetada para ser executada por um acionador de tempo (trigger).
 */
function automatizacao_cancelarCotacoesAntigas() {
  Logger.log("Iniciando rotina de cancelamento de cotações antigas.");
  try {
    const abaCotacoes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_COTACOES);
    const ultimaLinha = abaCotacoes.getLastRow();
    if (ultimaLinha <= 1) return;

    const range = abaCotacoes.getRange(2, 1, ultimaLinha - 1, abaCotacoes.getLastColumn());
    const todosOsValores = range.getValues();
    const cabecalhos = Utilities_obterCabecalhos(ABA_COTACOES);

    const indiceIdCotacao = cabecalhos.indexOf("ID da Cotação");
    const indiceDataAbertura = cabecalhos.indexOf("Data Abertura");
    const indiceStatusCotacao = cabecalhos.indexOf("Status da Cotação");
    const indiceStatusSubproduto = cabecalhos.indexOf("Status do SubProduto");

    const hoje = new Date();
    const tresMesesAtras = new Date(hoje.getFullYear(), hoje.getMonth() - 3, hoje.getDate());
    
    // Agrupa todas as linhas por ID de Cotação
    const cotacoesAgrupadas = todosOsValores.reduce((acc, linha) => {
      const id = linha[indiceIdCotacao];
      if (!id) return acc;
      if (!acc[id]) {
        acc[id] = {
          linhas: [],
          dataAbertura: new Date(linha[indiceDataAbertura]),
          statusAtual: linha[indiceStatusCotacao]
        };
      }
      acc[id].linhas.push(linha);
      return acc;
    }, {});

    const idsParaCancelar = [];

    // Itera sobre as cotações agrupadas para identificar quais devem ser canceladas
    for (const id in cotacoesAgrupadas) {
      const cotacao = cotacoesAgrupadas[id];

      // Pula cotações já finalizadas/canceladas ou que não são antigas o suficiente
      if (cotacao.statusAtual === "Finalizado" || cotacao.statusAtual === "Cancelado" || cotacao.dataAbertura > tresMesesAtras) {
        continue;
      }
      
      // Verifica se algum item já foi cancelado manualmente
      const temItemCancelado = cotacao.linhas.some(linha => linha[indiceStatusSubproduto] === "Cancelado");
      
      if (!temItemCancelado) {
        idsParaCancelar.push(id);
      }
    }

    if (idsParaCancelar.length > 0) {
      Logger.log(`Cotações a serem canceladas: ${idsParaCancelar.join(', ')}`);
      // Itera sobre a matriz de dados e atualiza o status
      todosOsValores.forEach((linha, index) => {
        const idLinha = linha[indiceIdCotacao].toString();
        if (idsParaCancelar.includes(idLinha)) {
          todosOsValores[index][indiceStatusCotacao] = "Cancelado";
        }
      });
      
      // Escreve a matriz atualizada de volta na planilha
      range.setValues(todosOsValores);
      SpreadsheetApp.flush();
    } else {
      Logger.log("Nenhuma cotação antiga encontrada para cancelar.");
    }

  } catch (e) {
    Logger.log(`ERRO na rotina automatizacao_cancelarCotacoesAntigas: ${e.toString()}`);
  }
}