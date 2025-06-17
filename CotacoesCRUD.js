// @ts-nocheck

/**
 * @file CotacoesCRUD.gs
 * @description Funções CRUD para a aba "Cotacoes", com foco em retornar resumos de cotações e criar novas cotações.
 */

const CotacoesCRUD_STATUS_NOVA_COTACAO = "Nova Cotação";

/**
 * Lê todas as linhas da aba "Cotacoes" e processa para retornar um resumo para cada "ID da Cotação" único.
 * Cada resumo incluirá o ID da Cotação, Data de Abertura (formatada e como string ISO), Status e uma lista de Categorias únicas.
 *
 * @return {Array<object>|null} Um array de objetos de resumo de cotação, ou null em caso de erro.
 */
function CotacoesCRUD_obterResumosDeCotacoes() {
  console.log("CotacoesCRUD_obterResumosDeCotacoes: Iniciando execução.");
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaCotacoes = planilha.getSheetByName(ABA_COTACOES);

    if (!abaCotacoes) {
      console.error(`CotacoesCRUD: Aba "${ABA_COTACOES}" não encontrada.`);
      return null;
    }
    
    const ultimaLinha = abaCotacoes.getLastRow();
    if (ultimaLinha <= 1) {
      console.log(`CotacoesCRUD: Aba "${ABA_COTACOES}" vazia ou só cabeçalho.`);
      return [];
    }
    
    const ultimaColuna = abaCotacoes.getLastColumn();
    const rangeDados = abaCotacoes.getRange(2, 1, ultimaLinha - 1, ultimaColuna);
    const todasAsLinhasDeProdutos = rangeDados.getValues();
    
    const cabecalhos = Utilities_obterCabecalhos(ABA_COTACOES); 
    if (!cabecalhos || cabecalhos.length === 0) {
        console.error("CotacoesCRUD: Cabeçalhos da aba Cotações não puderam ser lidos.");
         return null; 
    }


    const indiceIdCotacao = cabecalhos.indexOf("ID da Cotação");
    const indiceDataAbertura = cabecalhos.indexOf("Data Abertura");
    const indiceCategoria = cabecalhos.indexOf("Categoria");
    const indiceStatus = cabecalhos.indexOf("Status da Cotação");

    if ([indiceIdCotacao, indiceDataAbertura, indiceCategoria, indiceStatus].includes(-1)) {
        console.error("CotacoesCRUD: Um ou mais cabeçalhos essenciais (ID da Cotação, Data Abertura, Categoria, Status da Cotação) não foram encontrados nos cabeçalhos lidos para a aba Cotações.");
        return null;
    }

    const cotacoesUnicas = {};

    todasAsLinhasDeProdutos.forEach((linhaProduto, i) => {
      let idCotacao = linhaProduto[indiceIdCotacao]; 
      if (idCotacao === null || idCotacao === undefined || idCotacao === "") return; 
      idCotacao = String(idCotacao); 

      const dataAberturaValor = linhaProduto[indiceDataAbertura]; 
      const categoriaProduto = linhaProduto[indiceCategoria];
      const statusCotacao = linhaProduto[indiceStatus];


      if (!cotacoesUnicas[idCotacao]) {
        cotacoesUnicas[idCotacao] = {
          ID_da_Cotacao: linhaProduto[indiceIdCotacao], 
          Data_Abertura_Original_ISO: null, 
          Data_Abertura_Formatada: "N/A",
          Status_da_Cotacao: statusCotacao || "Status Desconhecido", 
          Categorias: new Set()
        };

        if (dataAberturaValor) {
            try {
                const dataObj = new Date(dataAberturaValor);
                if (!isNaN(dataObj.getTime())) {
                    cotacoesUnicas[idCotacao].Data_Abertura_Original_ISO = dataObj.toISOString();
                    cotacoesUnicas[idCotacao].Data_Abertura_Formatada = Utilities_formatarDataParaDDMMYYYY(dataObj);
                }
            } catch(e) { /* Ignora data inválida */ }
        }
      }
      
      if (statusCotacao && (!cotacoesUnicas[idCotacao].Status_da_Cotacao || cotacoesUnicas[idCotacao].Status_da_Cotacao === "Status Desconhecido")) {
        cotacoesUnicas[idCotacao].Status_da_Cotacao = statusCotacao;
      }


      if (categoriaProduto) cotacoesUnicas[idCotacao].Categorias.add(categoriaProduto);
    });

    const arrayDeResumos = Object.values(cotacoesUnicas).map(cotacao => ({
      ID_da_Cotacao: cotacao.ID_da_Cotacao, 
      Data_Abertura_Original: cotacao.Data_Abertura_Original_ISO,
      Data_Abertura_Formatada: cotacao.Data_Abertura_Formatada,
      Status_da_Cotacao: cotacao.Status_da_Cotacao,
      Categorias_Unicas_String: Array.from(cotacao.Categorias).join(', ')
    }));
    
    return arrayDeResumos;

  } catch (error) {
    console.error("!!!!!!!! ERRO CAPTURADO em CotacoesCRUD_obterResumosDeCotacoes !!!!!!!!");
    console.error("Mensagem do Erro: " + error.toString(), error.stack);
    return null;
  }
}

/**
 * Gera o próximo ID de cotação disponível como um número inteiro.
 * @return {number|null} O próximo ID numérico ou null em caso de erro.
 */
function CotacoesCRUD_gerarProximoIdCotacao() {
  console.log("CotacoesCRUD_gerarProximoIdCotacao: Iniciando.");
  try {
    const abaCotacoes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_COTACOES);
    if (!abaCotacoes) {
      console.error("CotacoesCRUD_gerarProximoIdCotacao: Aba Cotações não encontrada.");
      return null;
    }
    const cabecalhos = Utilities_obterCabecalhos(ABA_COTACOES);
    const indiceIdCotacao = cabecalhos.indexOf("ID da Cotação");
    if (indiceIdCotacao === -1) {
        console.error("CotacoesCRUD_gerarProximoIdCotacao: Coluna 'ID da Cotação' não encontrada.");
        return null;
    }

    const ultimaLinha = abaCotacoes.getLastRow();
    let maxIdNum = 0;
    if (ultimaLinha > 1) {
      const idsExistentes = abaCotacoes.getRange(2, indiceIdCotacao + 1, ultimaLinha - 1, 1).getValues();
      idsExistentes.forEach(row => {
        const idCellValue = row[0];
        if (idCellValue !== null && idCellValue !== undefined && idCellValue !== "") {
          const numPart = parseInt(idCellValue, 10); 
          if (!isNaN(numPart) && numPart > maxIdNum) {
            maxIdNum = numPart;
          }
        }
      });
    }
    const proximoNum = maxIdNum + 1;
    console.log("CotacoesCRUD_gerarProximoIdCotacao: Próximo ID numérico gerado:", proximoNum);
    return proximoNum; 
  } catch (e) {
    console.error("Erro em CotacoesCRUD_gerarProximoIdCotacao: " + e.toString(), e.stack);
    return null;
  }
}


/**
 * Função auxiliar para ler dados de uma aba e converter para array de objetos.
 * @param {string} nomeAba O nome da aba.
 * @param {Array<string>} cabecalhosEsperados Array com os nomes dos cabeçalhos.
 * @return {Array<Object>|null} Array de objetos ou null se erro.
 */
function CotacoesCRUD_obterDadosCompletosDaAba(nomeAba, cabecalhosEsperados) {
    console.log(`CotacoesCRUD_obterDadosCompletosDaAba: Lendo aba "${nomeAba}"`);
    try {
        const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
        if (!aba) {
            console.error(`Aba "${nomeAba}" não encontrada.`);
            return null;
        }
        const ultimaLinha = aba.getLastRow();
        if (ultimaLinha <= 1) return []; 

        const numColunasParaLer = Math.min(aba.getLastColumn(), cabecalhosEsperados.length);
        if (numColunasParaLer === 0) return [];


        const dadosRange = aba.getRange(1, 1, ultimaLinha, numColunasParaLer).getValues();
        const dadosObjetos = [];

        for (let i = 1; i < dadosRange.length; i++) {
            const linha = dadosRange[i];
            const objLinha = {};
            for (let j = 0; j < numColunasParaLer; j++) {
                if (j < cabecalhosEsperados.length) { 
                    objLinha[cabecalhosEsperados[j]] = linha[j];
                }
            }
            for (let j = numColunasParaLer; j < cabecalhosEsperados.length; j++) {
                objLinha[cabecalhosEsperados[j]] = undefined; 
            }
            dadosObjetos.push(objLinha);
        }
        console.log(`CotacoesCRUD_obterDadosCompletosDaAba: ${dadosObjetos.length} registros lidos da aba "${nomeAba}".`);
        return dadosObjetos;
    } catch (e) {
        console.error(`Erro ao ler dados da aba "${nomeAba}": ${e.toString()}`, e.stack);
        return null;
    }
}


/**
 * Cria uma nova cotação na aba "Cotacoes".
 * @param {object} opcoesCriacao { tipo: string, selecoes: Array<string> }.
 * @return {object} { success: boolean, idCotacao: string|null, numItens: number|null, message: string|null }.
 */
function CotacoesCRUD_criarNovaCotacao(opcoesCriacao) {
  console.log("CotacoesCRUD_criarNovaCotacao: Iniciando com opções:", JSON.stringify(opcoesCriacao));
  try {
    const novoIdCotacaoNumerico = CotacoesCRUD_gerarProximoIdCotacao(); 
    if (novoIdCotacaoNumerico === null) { 
      return { success: false, message: "Falha ao gerar novo ID de cotação." };
    }
    const dataAbertura = new Date();

    const todosSubProdutos = CotacoesCRUD_obterDadosCompletosDaAba(ABA_SUBPRODUTOS, CABECALHOS_SUBPRODUTOS);
    const todosProdutos = CotacoesCRUD_obterDadosCompletosDaAba(ABA_PRODUTOS, CABECALHOS_PRODUTOS);

    if (!todosSubProdutos || !todosProdutos) {
      return { success: false, message: "Falha ao carregar dados de Produtos ou SubProdutos." };
    }
    
    const produtosMap = todosProdutos.reduce((map, prod) => {
        map[prod["Produto"]] = prod; // Chave é o nome do Produto (conforme confirmado pelo usuário)
        return map;
    }, {});


    let subProdutosFiltrados = [];
    const tipo = opcoesCriacao.tipo;
    // As seleções são os NOMES das categorias, fornecedores ou produtos (conforme enviado pelo modal)
    const selecoesLowerCase = opcoesCriacao.selecoes.map(s => String(s).toLowerCase()); 

    console.log(`CotacoesCRUD_criarNovaCotacao: Tipo "${tipo}", Seleções em LowerCase: ${JSON.stringify(selecoesLowerCase)}`);

    if (tipo === 'categoria') {
      const nomesProdutosDaCategoria = todosProdutos
                                        .filter(p => p["Categoria"] && selecoesLowerCase.includes(String(p["Categoria"]).toLowerCase()))
                                        .map(p => String(p["Produto"]).toLowerCase()); 
      subProdutosFiltrados = todosSubProdutos.filter(sp => {
          const produtoVinculado = sp["Produto Vinculado"] ? String(sp["Produto Vinculado"]).toLowerCase() : null;
          return produtoVinculado && nomesProdutosDaCategoria.includes(produtoVinculado);
      });
    } else if (tipo === 'fornecedor') {
      // Filtra subprodutos cujo campo "Fornecedor" (na aba SubProdutos) corresponde a um dos nomes de fornecedores selecionados.
      subProdutosFiltrados = todosSubProdutos.filter(sp => {
          const fornecedorSubProduto = sp["Fornecedor"] ? String(sp["Fornecedor"]).toLowerCase() : null;
          return fornecedorSubProduto && selecoesLowerCase.includes(fornecedorSubProduto);
      });
    } else if (tipo === 'curvaABC') {
      const nomesProdutosDaCurva = todosProdutos
                                    .filter(p => p["ABC"] && selecoesLowerCase.includes(String(p["ABC"]).toLowerCase()))
                                    .map(p => String(p["Produto"]).toLowerCase()); 
      subProdutosFiltrados = todosSubProdutos.filter(sp => {
          const produtoVinculado = sp["Produto Vinculado"] ? String(sp["Produto Vinculado"]).toLowerCase() : null;
          return produtoVinculado && nomesProdutosDaCurva.includes(produtoVinculado);
      });
    } else if (tipo === 'produtoEspecifico') {
      // Filtra subprodutos cujo campo "Produto Vinculado" (na aba SubProdutos) corresponde a um dos nomes de produtos principais selecionados.
      subProdutosFiltrados = todosSubProdutos.filter(sp => {
          const produtoVinculado = sp["Produto Vinculado"] ? String(sp["Produto Vinculado"]).toLowerCase() : null;
          return produtoVinculado && selecoesLowerCase.includes(produtoVinculado);
      });
    } else {
      return { success: false, message: "Tipo de criação desconhecido: " + tipo };
    }
    
    console.log(`CotacoesCRUD_criarNovaCotacao: ${subProdutosFiltrados.length} subprodutos filtrados para o tipo "${tipo}".`);

    if (subProdutosFiltrados.length === 0) {
      return { success: true, idCotacao: novoIdCotacaoNumerico, numItens: 0, message: "Nenhum subproduto encontrado para os critérios selecionados. Cotação criada vazia." };
    }

    const linhasParaAdicionar = subProdutosFiltrados.map(subProd => {
      const produtoPrincipal = produtosMap[subProd["Produto Vinculado"]]; // Busca pelo nome
      const estoqueMinimo = produtoPrincipal ? produtoPrincipal["Estoque Minimo"] : "";
      const nomeProdutoPrincipalParaCotacao = subProd["Produto Vinculado"]; // Já é o nome correto


      let linha = []; 
      CABECALHOS_COTACOES.forEach(header => {
        switch(header) {
          case "ID da Cotação": linha.push(novoIdCotacaoNumerico); break; 
          case "Data Abertura": linha.push(dataAbertura); break;
          case "Produto": linha.push(nomeProdutoPrincipalParaCotacao); break; 
          case "SubProduto": linha.push(subProd["SubProduto"]); break;
          case "Categoria": linha.push(produtoPrincipal ? produtoPrincipal["Categoria"] : subProd["Categoria"]); break; // Categoria do Produto Principal, se existir, senão do SubProduto
          case "Fornecedor": linha.push(subProd["Fornecedor"]); break;
          case "Tamanho": linha.push(subProd["Tamanho"]); break;
          case "UN": linha.push(subProd["UN"]); break;
          case "Fator": linha.push(subProd["Fator"]); break;
          case "Estoque Mínimo": linha.push(estoqueMinimo); break;
          case "NCM": linha.push(subProd["NCM"]); break;
          case "CST": linha.push(subProd["CST"]); break;
          case "CFOP": linha.push(subProd["CFOP"]); break;
          case "Status da Cotação": linha.push(CotacoesCRUD_STATUS_NOVA_COTACAO); break;
          case "Estoque Atual":
          case "Preço":
          case "Preço por Fator":
          case "Comprar":
          case "Valor Total":
          case "Economia em Cotação":
          case "Empresa Faturada":
          case "Condição de Pagamento":
            linha.push(""); 
            break;
          default:
            linha.push(""); 
        }
      });
      return linha;
    });

    const abaCotacoes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_COTACOES);
    abaCotacoes.getRange(abaCotacoes.getLastRow() + 1, 1, linhasParaAdicionar.length, CABECALHOS_COTACOES.length)
               .setValues(linhasParaAdicionar);
    
    console.log(`CotacoesCRUD_criarNovaCotacao: ${linhasParaAdicionar.length} itens adicionados à cotação ${novoIdCotacaoNumerico}.`);
    return { 
      success: true, 
      idCotacao: novoIdCotacaoNumerico, 
      numItens: linhasParaAdicionar.length,
      message: "Nova cotação criada com sucesso."
    };

  } catch (e) {
    console.error("Erro em CotacoesCRUD_criarNovaCotacao: " + e.toString(), e.stack);
    return { success: false, message: "Erro no servidor ao criar nova cotação: " + e.message };
  }
}


/**
 * Obtém uma lista única de categorias da aba Produtos.
 * @return {Array<string>|null} Array de nomes de categorias ou null em caso de erro.
 */
function CotacoesCRUD_obterListaCategoriasProdutos() {
  console.log("CotacoesCRUD_obterListaCategoriasProdutos: Iniciando.");
  try {
    const produtos = CotacoesCRUD_obterDadosCompletosDaAba(ABA_PRODUTOS, CABECALHOS_PRODUTOS);
    if (!produtos) return null;
    const categorias = [...new Set(produtos.map(p => p["Categoria"]).filter(cat => cat))]; 
    console.log(`CotacoesCRUD: ${categorias.length} categorias encontradas.`);
    return categorias.sort();
  } catch (e) {
    console.error("Erro em CotacoesCRUD_obterListaCategoriasProdutos: " + e.toString(), e.stack);
    return null;
  }
}

/**
 * Obtém uma lista de fornecedores (ID e Nome) da aba Fornecedores.
 * No modal, o 'nome' será usado como valor para a seleção.
 * @return {Array<{id: string, nome: string}>|null} Array de objetos ou null em caso de erro.
 */
function CotacoesCRUD_obterListaFornecedores() {
  console.log("CotacoesCRUD_obterListaFornecedores: Iniciando.");
  try {
    const fornecedoresDados = CotacoesCRUD_obterDadosCompletosDaAba(ABA_FORNECEDORES, CABECALHOS_FORNECEDORES);
    if (!fornecedoresDados) return null;
    const fornecedores = fornecedoresDados
        .map(f => ({ id: f["ID"], nome: f["Fornecedor"] })) 
        .filter(f => f.id && f.nome); 
    console.log(`CotacoesCRUD: ${fornecedores.length} fornecedores encontrados.`);
    return fornecedores.sort((a,b) => a.nome.localeCompare(b.nome));
  } catch (e) {
    console.error("Erro em CotacoesCRUD_obterListaFornecedores: " + e.toString(), e.stack);
    return null;
  }
}

/**
 * Obtém uma lista de produtos (ID e Nome) da aba Produtos.
 * No modal, o 'nome' será usado como valor para a seleção.
 * @return {Array<{id: string, nome: string}>|null} Array de objetos ou null em caso de erro.
 */
function CotacoesCRUD_obterListaProdutos() {
  console.log("CotacoesCRUD_obterListaProdutos: Iniciando.");
  try {
    const produtosDados = CotacoesCRUD_obterDadosCompletosDaAba(ABA_PRODUTOS, CABECALHOS_PRODUTOS);
    if (!produtosDados) return null;
    const produtos = produtosDados
        .map(p => ({ id: p["ID"], nome: p["Produto"] })) 
        .filter(p => p.id && p.nome); 
    console.log(`CotacoesCRUD: ${produtos.length} produtos encontrados.`);
    return produtos.sort((a,b) => a.nome.localeCompare(b.nome));
  } catch (e) {
    console.error("Erro em CotacoesCRUD_obterListaProdutos: " + e.toString(), e.stack);
    return null;
  }
}

// --- Funções Utilitárias ---
function Utilities_padLeft(str, len, char = '0') {
    str = String(str);
    return char.repeat(Math.max(0, len - str.length)) + str;
}

function Utilities_formatarDataParaDDMMYYYY(dateObj) {
    if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) {
        return "Data inválida";
    }
    const dia = Utilities_padLeft(dateObj.getDate().toString(), 2, '0');
    const mes = Utilities_padLeft((dateObj.getMonth() + 1).toString(), 2, '0'); 
    const ano = dateObj.getFullYear();
    return `${dia}/${mes}/${ano}`;
}

function Utilities_obterCabecalhos(nomeAba) {
    try {
        const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
        if (!aba) {
            console.warn(`Utilities_obterCabecalhos: Aba "${nomeAba}" não encontrada.`);
            switch(nomeAba) {
                case ABA_COTACOES: return CABECALHOS_COTACOES;
                case ABA_PRODUTOS: return CABECALHOS_PRODUTOS;
                case ABA_SUBPRODUTOS: return CABECALHOS_SUBPRODUTOS;
                case ABA_FORNECEDORES: return CABECALHOS_FORNECEDORES;
                default:
                    console.error(`Utilities_obterCabecalhos: Constante de cabeçalho não definida para "${nomeAba}".`);
                    return null;
            }
        }
        if (aba.getLastRow() === 0) return []; 
        const ultimaColuna = aba.getLastColumn();
        if (ultimaColuna === 0) return [];

        return aba.getRange(1, 1, 1, ultimaColuna).getValues()[0].map(String); 
    } catch (e) {
        console.error(`Erro em Utilities_obterCabecalhos para aba "${nomeAba}": ` + e.toString(), e.stack);
        return null;
    }
}
