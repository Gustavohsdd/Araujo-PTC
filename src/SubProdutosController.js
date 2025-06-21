// @ts-nocheck
// Arquivo: SubProdutosController.gs

/**
 * @OnlyCurrentDoc
 */

const SubProdutosController_NOMES_CABECALHOS_PARA_EXIBIR_NA_TABELA = [
  "SubProduto",
  "Produto Vinculado",
  "Fornecedor",
  "Categoria",
  "UN",
  "Status"
  // Adicione ou remova conforme necessário para a visualização da tabela principal de subprodutos
];

/**
 * Normaliza o texto para comparação: remove acentos, converte para minúsculas e remove espaços extras.
 * @param {string} texto O texto a ser normalizado.
 * @return {string} O texto normalizado.
 */
function SubProdutosController_normalizarTextoComparacao(texto) {
  if (!texto || typeof texto !== 'string') return "";
  return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
}

/**
 * Obtém os dados dos subprodutos de forma paginada e com filtro de busca.
 * @param {object} options Objeto com { pagina: number, itensPorPagina: number, termoBusca?: string }.
 * @return {object} { cabecalhosParaExibicao: Array<string>, subProdutosPaginados: Array<Object>, totalItens: number, paginaAtual: number, totalPaginas: number, error?: string, message?: string }
 */
function SubProdutosController_obterListaSubProdutosPaginada(options) {
  console.log("SubProdutosController_obterListaSubProdutosPaginada: Iniciando com options:", JSON.stringify(options));

  const colunasParaExibicao = SubProdutosController_NOMES_CABECALHOS_PARA_EXIBIR_NA_TABELA;

  try {
    // 1. Validar e obter opções
    const pagina = (options && typeof options.pagina === 'number' && options.pagina > 0) ? parseInt(options.pagina, 10) : 1;
    const itensPorPagina = (options && typeof options.itensPorPagina === 'number' && options.itensPorPagina > 0) ? parseInt(options.itensPorPagina, 10) : 10;
    const termoBusca = (options && options.termoBusca && typeof options.termoBusca === 'string') ?
      SubProdutosController_normalizarTextoComparacao(options.termoBusca) : null;

    console.log(`SubProdutosController_obterListaSubProdutosPaginada: Termo de busca normalizado: '${termoBusca}'`);

    // 2. Verificar constantes globais (essas são definidas em Constantes.gs)
    if (typeof NOME_PLANILHA === 'undefined' || typeof ABA_SUBPRODUTOS === 'undefined' || typeof CABECALHOS_SUBPRODUTOS === 'undefined' || CABECALHOS_SUBPRODUTOS.length === 0) {
      console.error("SubProdutosController_obterListaSubProdutosPaginada: Constantes NOME_PLANILHA, ABA_SUBPRODUTOS ou CABECALHOS_SUBPRODUTOS não definidas ou vazias.");
      return { error: "Erro de configuração interna: Constantes essenciais não definidas.", cabecalhosParaExibicao: colunasParaExibicao, subProdutosPaginados: [], totalItens: 0, paginaAtual: 1, totalPaginas: 0 };
    }

    // 3. Acessar planilha e aba
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Planilha ativa não acessível.");
    const abaSubProdutos = ss.getSheetByName(ABA_SUBPRODUTOS);
    if (!abaSubProdutos) throw new Error(`Aba "${ABA_SUBPRODUTOS}" não encontrada.`);

    // 4. Obter dados
    const range = abaSubProdutos.getDataRange();
    const todasAsLinhas = range.getValues();

    if (todasAsLinhas.length <= 1) { // Apenas cabeçalho ou vazia
      console.log("SubProdutosController_obterListaSubProdutosPaginada: Nenhum subproduto na planilha.");
      return { cabecalhosParaExibicao: colunasParaExibicao, subProdutosPaginados: [], totalItens: 0, paginaAtual: 1, totalPaginas: 1, message: "Nenhum subproduto cadastrado." };
    }

    // 5. Mapear para objetos
    const cabecalhosCompletosDaPlanilha = todasAsLinhas[0].map(String); // Cabeçalhos reais da planilha
    const dadosBrutosDaPlanilha = todasAsLinhas.slice(1);
    let todosSubProdutosObj = dadosBrutosDaPlanilha.map(linha => {
      const subProdutoObj = {};
      cabecalhosCompletosDaPlanilha.forEach((nomeChavePlanilha, index) => {
        // Encontrar o nome do cabeçalho correspondente em CABECALHOS_SUBPRODUTOS (de Constantes.gs)
        // para garantir que o objeto use as chaves definidas em Constantes.gs
        const nomeChaveConstante = CABECALHOS_SUBPRODUTOS.find(chConst => SubProdutosController_normalizarTextoComparacao(chConst) === SubProdutosController_normalizarTextoComparacao(nomeChavePlanilha));
        
        if (nomeChaveConstante) { // Se encontrou um correspondente em Constantes.gs
          let valorCelula = linha[index];
          // Formatar data se for a coluna "Data de Cadastro"
          if (nomeChaveConstante === CABECALHOS_SUBPRODUTOS[0] && valorCelula instanceof Date) { // CABECALHOS_SUBPRODUTOS[0] é "Data de Cadastro"
            valorCelula = Utilities.formatDate(valorCelula, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
          } else if (valorCelula instanceof Date) { // Para outras possíveis colunas de data
            valorCelula = Utilities.formatDate(valorCelula, Session.getScriptTimeZone(), "dd/MM/yyyy");
          }
          subProdutoObj[nomeChaveConstante] = valorCelula !== null && valorCelula !== undefined ? String(valorCelula) : "";
        }
      });
      // Adicionar o ID explicitamente se não foi mapeado corretamente ou para garantir
      const idxIdConstante = CABECALHOS_SUBPRODUTOS.indexOf("ID");
      const idxIdPlanilha = cabecalhosCompletosDaPlanilha.indexOf("ID");
      if (idxIdConstante !== -1 && idxIdPlanilha !== -1 && !subProdutoObj["ID"]) {
          subProdutoObj["ID"] = linha[idxIdPlanilha] !== null && linha[idxIdPlanilha] !== undefined ? String(linha[idxIdPlanilha]) : "";
      }
      return subProdutoObj;
    });

    // 6. Aplicar filtro de busca
    let subProdutosFiltrados = todosSubProdutosObj;
    if (termoBusca && termoBusca.length > 0) {
      console.log(`SubProdutosController_obterListaSubProdutosPaginada: Aplicando filtro com termo: '${termoBusca}'`);
      subProdutosFiltrados = todosSubProdutosObj.filter(subProduto => {
        // Busca em todas as colunas definidas em CABECALHOS_SUBPRODUTOS
        return CABECALHOS_SUBPRODUTOS.some(nomeCabecalhoConstante => {
          const valorDoCampo = subProduto[nomeCabecalhoConstante];
          if (valorDoCampo) {
            return SubProdutosController_normalizarTextoComparacao(String(valorDoCampo)).includes(termoBusca);
          }
          return false;
        });
      });
      console.log(`SubProdutosController_obterListaSubProdutosPaginada: ${subProdutosFiltrados.length} subprodutos após filtro.`);
    }

    // 7. Calcular totais e offset para paginação (sobre os dados filtrados)
    const totalItens = subProdutosFiltrados.length;
    const totalPaginas = Math.ceil(totalItens / itensPorPagina) || 1;
    const paginaAjustada = Math.min(Math.max(1, pagina), totalPaginas);
    const offset = (paginaAjustada - 1) * itensPorPagina;

    // 8. Fatiar os objetos para a página atual
    const subProdutosPaginadosObj = subProdutosFiltrados.slice(offset, offset + itensPorPagina);

    console.log(`SubProdutosController_obterListaSubProdutosPaginada: Página Solicitada ${pagina}, Ajustada ${paginaAjustada}. Itens na página: ${subProdutosPaginadosObj.length}. Total Itens Filtrados: ${totalItens}. Total Páginas: ${totalPaginas}`);

    // 9. Construir o objeto de retorno
    const objetoDeRetorno = {
      cabecalhosParaExibicao: colunasParaExibicao, // Usar os cabeçalhos definidos para exibição
      subProdutosPaginados: subProdutosPaginadosObj,
      totalItens: totalItens,
      paginaAtual: paginaAjustada,
      totalPaginas: totalPaginas
    };
    
    console.log("SubProdutosController_obterListaSubProdutosPaginada: Retornando objeto de dados.");
    return objetoDeRetorno;

  } catch (e) {
    console.error("ERRO em SubProdutosController_obterListaSubProdutosPaginada: " + e.toString() + " Stack: " + (e.stack || 'N/A'));
    const paginaErro = (options && options.pagina) ? parseInt(options.pagina, 10) : 1;
    return { error: "Falha ao buscar dados dos subprodutos. Detalhes: " + e.message, cabecalhosParaExibicao: colunasParaExibicao, subProdutosPaginados: [], totalItens: 0, paginaAtual: paginaErro, totalPaginas: 0 };
  } finally {
    console.log("SubProdutosController_obterListaSubProdutosPaginada: Bloco finally executado.");
  }
}

/**
 * Obtém todos os produtos para popular dropdowns.
 * @return {Array<Object>} Array de objetos de produtos, cada um com {ID, Produto}.
 */
function SubProdutosController_obterTodosProdutosParaDropdown() {
  try {
    console.log("SubProdutosController_obterTodosProdutosParaDropdown: Iniciando busca de produtos.");
    const produtos = SubProdutosCRUD_obterTodosProdutosParaDropdown(); // Chama a função do CRUD
    console.log(`SubProdutosController_obterTodosProdutosParaDropdown: ${produtos.length} produtos encontrados.`);
    return produtos;
  } catch (e) {
    console.error("ERRO em SubProdutosController_obterTodosProdutosParaDropdown: " + e.toString());
    // Lançar o erro para que o .withFailureHandler no cliente possa pegá-lo
    throw new Error("Falha ao buscar lista de produtos: " + e.message);
  }
}

/**
 * Obtém todos os fornecedores para popular dropdowns.
 * @return {Array<Object>} Array de objetos de fornecedores, cada um com {ID, Fornecedor}.
 */
function SubProdutosController_obterTodosFornecedoresParaDropdown() {
  try {
    console.log("SubProdutosController_obterTodosFornecedoresParaDropdown: Iniciando busca de fornecedores.");
    const fornecedores = SubProdutosCRUD_obterTodosFornecedoresParaDropdown(); // Chama a função do CRUD
    console.log(`SubProdutosController_obterTodosFornecedoresParaDropdown: ${fornecedores.length} fornecedores encontrados.`);
    return fornecedores;
  } catch (e) {
    console.error("ERRO em SubProdutosController_obterTodosFornecedoresParaDropdown: " + e.toString());
    // Lançar o erro para que o .withFailureHandler no cliente possa pegá-lo
    throw new Error("Falha ao buscar lista de fornecedores: " + e.message);
  }
}
