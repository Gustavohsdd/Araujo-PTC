// @ts-nocheck
// Arquivo: ProdutosController.gs

/**
 * @OnlyCurrentDoc
 */

const ProdutosController_NOMES_CABECALHOS_PARA_EXIBIR_NA_TABELA = [
  "Produto", 
  "Tamanho", 
  "UN", 
  "Estoque Minimo", 
  "Status"
];

/**
 * Normaliza o texto para comparação: remove acentos, converte para minúsculas e remove espaços extras.
 * @param {string} texto O texto a ser normalizado.
 * @return {string} O texto normalizado.
 */
function ProdutosController_normalizarTextoComparacao(texto) {
  if (!texto || typeof texto !== 'string') return "";
  return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
}

/**
 * Obtém os dados dos produtos de forma paginada e com filtro de busca.
 * @param {object} options Objeto com { pagina: number, itensPorPagina: number, termoBusca?: string }.
 * @return {object} { cabecalhosParaExibicao: Array<string>, produtosPaginados: Array<Object>, totalItens: number, paginaAtual: number, totalPaginas: number, error?: string }
 */
function ProdutosController_obterListaProdutosPaginada(options) {
  console.log("ProdutosController_obterListaProdutosPaginada: Iniciando com options:", JSON.stringify(options));

  const colunasParaExibicao = ProdutosController_NOMES_CABECALHOS_PARA_EXIBIR_NA_TABELA;

  try {
    // 1. Validar e obter opções
    const pagina = (options && typeof options.pagina === 'number' && options.pagina > 0) ? parseInt(options.pagina, 10) : 1;
    const itensPorPagina = (options && typeof options.itensPorPagina === 'number' && options.itensPorPagina > 0) ? parseInt(options.itensPorPagina, 10) : 10;
    const termoBusca = (options && options.termoBusca && typeof options.termoBusca === 'string') ? 
                       ProdutosController_normalizarTextoComparacao(options.termoBusca) : null;
    
    console.log(`ProdutosController_obterListaProdutosPaginada: Termo de busca normalizado: '${termoBusca}'`);

    // 2. Verificar constantes globais
    if (typeof NOME_PLANILHA === 'undefined' || typeof ABA_PRODUTOS === 'undefined' || typeof CABECALHOS_PRODUTOS === 'undefined' || CABECALHOS_PRODUTOS.length === 0) {
      console.error("ProdutosController_obterListaProdutosPaginada: Constantes NOME_PLANILHA, ABA_PRODUTOS ou CABECALHOS_PRODUTOS não definidas ou vazias.");
      return { error: "Erro de configuração interna: Constantes essenciais não definidas.", cabecalhosParaExibicao: colunasParaExibicao, produtosPaginados: [], totalItens: 0, paginaAtual: 1, totalPaginas: 0 };
    }

    // 3. Acessar planilha e aba
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Planilha ativa não acessível.");
    const abaProdutos = ss.getSheetByName(ABA_PRODUTOS);
    if (!abaProdutos) throw new Error(`Aba "${ABA_PRODUTOS}" não encontrada.`);

    // 4. Obter dados
    const range = abaProdutos.getDataRange();
    const todasAsLinhas = range.getValues();

    if (todasAsLinhas.length <= 1) {
      console.log("ProdutosController_obterListaProdutosPaginada: Nenhum produto na planilha.");
      return { cabecalhosParaExibicao: colunasParaExibicao, produtosPaginados: [], totalItens: 0, paginaAtual: 1, totalPaginas: 1 };
    }

    // 5. Mapear para objetos
    const cabecalhosCompletosDaPlanilha = CABECALHOS_PRODUTOS;
    const dadosBrutosDaPlanilha = todasAsLinhas.slice(1);
    let todosProdutosObj = dadosBrutosDaPlanilha.map(linha => {
      const produtoObj = {};
      cabecalhosCompletosDaPlanilha.forEach((nomeChave, index) => {
        let valorCelula = linha[index];
        if (nomeChave === CABECALHOS_PRODUTOS[0] && valorCelula instanceof Date) { // "Data de Cadastro"
          valorCelula = Utilities.formatDate(valorCelula, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        } else if (valorCelula instanceof Date) {
          valorCelula = Utilities.formatDate(valorCelula, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
        produtoObj[nomeChave] = valorCelula !== null && valorCelula !== undefined ? String(valorCelula) : "";
      });
      return produtoObj;
    });

    // 6. Aplicar filtro de busca (NOVA LÓGICA)
    let produtosFiltrados = todosProdutosObj;
    if (termoBusca && termoBusca.length > 0) {
      console.log(`ProdutosController_obterListaProdutosPaginada: Aplicando filtro com termo: '${termoBusca}'`);
      produtosFiltrados = todosProdutosObj.filter(produto => {
        // Busca em todas as colunas definidas em CABECALHOS_PRODUTOS
        return CABECALHOS_PRODUTOS.some(nomeCabecalho => {
          const valorDoCampo = produto[nomeCabecalho]; // produtoObj já tem as chaves corretas
          if (valorDoCampo) { // Checa se o campo existe no objeto
            return ProdutosController_normalizarTextoComparacao(String(valorDoCampo)).includes(termoBusca);
          }
          return false;
        });
      });
      console.log(`ProdutosController_obterListaProdutosPaginada: ${produtosFiltrados.length} produtos após filtro.`);
    }


    // 7. Calcular totais e offset para paginação (sobre os dados filtrados)
    const totalItens = produtosFiltrados.length;
    const totalPaginas = Math.ceil(totalItens / itensPorPagina) || 1;
    const paginaAjustada = Math.min(Math.max(1, pagina), totalPaginas);
    const offset = (paginaAjustada - 1) * itensPorPagina;

    // 8. Fatiar os objetos para a página atual
    const produtosPaginadosObj = produtosFiltrados.slice(offset, offset + itensPorPagina);

    console.log(`ProdutosController_obterListaProdutosPaginada: Página Solicitada ${pagina}, Ajustada ${paginaAjustada}. Itens na página: ${produtosPaginadosObj.length}. Total Itens Filtrados: ${totalItens}. Total Páginas: ${totalPaginas}`);

    // 9. Construir o objeto de retorno
    const objetoDeRetorno = {
      cabecalhosParaExibicao: colunasParaExibicao,
      produtosPaginados: produtosPaginadosObj,
      totalItens: totalItens,
      paginaAtual: paginaAjustada,
      totalPaginas: totalPaginas
    };

    try {
      const jsonString = JSON.stringify(objetoDeRetorno);
      console.log("ProdutosController_obterListaProdutosPaginada: Objeto de retorno serializado. Tamanho: " + jsonString.length + ". Início: " + jsonString.substring(0, 250));
    } catch (jsonError) {
      console.error("ProdutosController_obterListaProdutosPaginada: ERRO AO SERIALIZAR OBJETO DE RETORNO: " + jsonError.toString());
    }
    
    console.log("ProdutosController_obterListaProdutosPaginada: Retornando objeto de dados.");
    return objetoDeRetorno;

  } catch (e) {
    console.error("ERRO em ProdutosController_obterListaProdutosPaginada: " + e.toString() + " Stack: " + e.stack);
    const paginaErro = (options && options.pagina) ? parseInt(options.pagina, 10) : 1;
    return { error: "Falha ao buscar dados dos produtos. Detalhes: " + e.message, cabecalhosParaExibicao: colunasParaExibicao, produtosPaginados: [], totalItens: 0, paginaAtual: paginaErro, totalPaginas: 0 };
  } finally {
    console.log("ProdutosController_obterListaProdutosPaginada: Bloco finally executado.");
  }
}