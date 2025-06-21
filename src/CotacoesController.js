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