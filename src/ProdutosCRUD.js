// @ts-nocheck
// Arquivo: ProdutosCRUD.gs

/**
 * @OnlyCurrentDoc
 */

// Constantes da Planilha 
const ProdutosCRUD_ABA_PRODUTOS = ABA_PRODUTOS; 
const ProdutosCRUD_CABECALHOS_PRODUTOS = CABECALHOS_PRODUTOS; 
const ProdutosCRUD_ABA_SUBPRODUTOS = ABA_SUBPRODUTOS;
const ProdutosCRUD_CABECALHOS_SUBPRODUTOS = CABECALHOS_SUBPRODUTOS;
const ProdutosCRUD_ABA_FORNECEDORES = ABA_FORNECEDORES;
const ProdutosCRUD_CABECALHOS_FORNECEDORES = CABECALHOS_FORNECEDORES;

// Índices das colunas Produtos
const ProdutosCRUD_IDX_PRODUTO_DATA_CADASTRO = CABECALHOS_PRODUTOS.indexOf("Data de Cadastro");
const ProdutosCRUD_IDX_PRODUTO_ID = CABECALHOS_PRODUTOS.indexOf("ID");
const ProdutosCRUD_IDX_PRODUTO_NOME = CABECALHOS_PRODUTOS.indexOf("Produto");

// Índices das colunas SubProdutos
const ProdutosCRUD_IDX_SUBPRODUTO_DATA_CADASTRO = CABECALHOS_SUBPRODUTOS.indexOf("Data de Cadastro");
const ProdutosCRUD_IDX_SUBPRODUTO_ID = CABECALHOS_SUBPRODUTOS.indexOf("ID");
const ProdutosCRUD_IDX_SUBPRODUTO_NOME = CABECALHOS_SUBPRODUTOS.indexOf("SubProduto");
const ProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO = CABECALHOS_SUBPRODUTOS.indexOf("Produto Vinculado");
const ProdutosCRUD_IDX_SUBPRODUTO_FORNECEDOR = CABECALHOS_SUBPRODUTOS.indexOf("Fornecedor");
const ProdutosCRUD_IDX_SUBPRODUTO_CATEGORIA = CABECALHOS_SUBPRODUTOS.indexOf("Categoria");
const ProdutosCRUD_IDX_SUBPRODUTO_TAMANHO = CABECALHOS_SUBPRODUTOS.indexOf("Tamanho");
const ProdutosCRUD_IDX_SUBPRODUTO_UN = CABECALHOS_SUBPRODUTOS.indexOf("UN");
const ProdutosCRUD_IDX_SUBPRODUTO_FATOR = CABECALHOS_SUBPRODUTOS.indexOf("Fator");
const ProdutosCRUD_IDX_SUBPRODUTO_NCM = CABECALHOS_SUBPRODUTOS.indexOf("NCM");
const ProdutosCRUD_IDX_SUBPRODUTO_CST = CABECALHOS_SUBPRODUTOS.indexOf("CST");
const ProdutosCRUD_IDX_SUBPRODUTO_CFOP = CABECALHOS_SUBPRODUTOS.indexOf("CFOP");
const ProdutosCRUD_IDX_SUBPRODUTO_STATUS = CABECALHOS_SUBPRODUTOS.indexOf("Status");

// Índices das colunas Fornecedores
const ProdutosCRUD_IDX_FORNECEDOR_ID = CABECALHOS_FORNECEDORES.indexOf("ID");
const ProdutosCRUD_IDX_FORNECEDOR_NOME = CABECALHOS_FORNECEDORES.indexOf("Fornecedor");

/**
 * Normaliza o texto para comparação: remove acentos, converte para minúsculas e remove espaços extras.
 * @param {string} texto O texto a ser normalizado.
 * @return {string} O texto normalizado.
 */
function ProdutosCRUD_normalizarTextoComparacao(texto) {
  if (!texto || typeof texto !== 'string') return "";
  return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
}

/**
 * Cria um novo produto na planilha.
 * @param {object} dadosNovoProduto Objeto contendo os dados do novo produto.
 * @return {object} { success: boolean, message: string, novoId?: string }
 */
function ProdutosCRUD_criarNovoProduto(dadosNovoProduto) {
  try {
    console.log("ProdutosCRUD_criarNovoProduto: Iniciando com dados:", JSON.stringify(dadosNovoProduto));

    const nomeDoCampoProduto = CABECALHOS_PRODUTOS[ProdutosCRUD_IDX_PRODUTO_NOME];
    if (!dadosNovoProduto || !dadosNovoProduto[nomeDoCampoProduto]) {
      throw new Error(`O campo '${nomeDoCampoProduto}' é obrigatório.`);
    }
    const nomeNovoProdutoNormalizado = ProdutosCRUD_normalizarTextoComparacao(dadosNovoProduto[nomeDoCampoProduto]);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaProdutos = ss.getSheetByName(ProdutosCRUD_ABA_PRODUTOS);
    if (!abaProdutos) throw new Error(`Aba '${ProdutosCRUD_ABA_PRODUTOS}' não encontrada.`);

    const range = abaProdutos.getDataRange();
    const todasAsLinhas = range.getValues();

    if (ProdutosCRUD_IDX_PRODUTO_NOME === -1) {
        throw new Error(`Coluna '${nomeDoCampoProduto}' não encontrada nos cabeçalhos. Verifique Constantes.gs.`);
    }
    for (let i = 1; i < todasAsLinhas.length; i++) {
      const nomeLinhaAtual = todasAsLinhas[i][ProdutosCRUD_IDX_PRODUTO_NOME];
      if (ProdutosCRUD_normalizarTextoComparacao(String(nomeLinhaAtual)) === nomeNovoProdutoNormalizado) {
        throw new Error(`O produto '${dadosNovoProduto[nomeDoCampoProduto]}' já está cadastrado.`);
      }
    }

    let proximoId = 1;
    if (ProdutosCRUD_IDX_PRODUTO_ID === -1) {
        throw new Error("Coluna 'ID' não encontrada nos cabeçalhos. Não é possível gerar novo ID.");
    }
    if (todasAsLinhas.length > 1) {
      const idsExistentes = todasAsLinhas.slice(1)
                                        .map(linha => parseInt(linha[ProdutosCRUD_IDX_PRODUTO_ID]))
                                        .filter(id => !isNaN(id));
      if (idsExistentes.length > 0) {
        proximoId = Math.max(...idsExistentes) + 1;
      }
    }
    const novoIdGerado = String(proximoId);

    const novaLinhaArray = [];
    ProdutosCRUD_CABECALHOS_PRODUTOS.forEach(nomeCabecalho => {
      if (nomeCabecalho === CABECALHOS_PRODUTOS[ProdutosCRUD_IDX_PRODUTO_DATA_CADASTRO]) {
        novaLinhaArray.push(new Date());
      } else if (nomeCabecalho === CABECALHOS_PRODUTOS[ProdutosCRUD_IDX_PRODUTO_ID]) {
        novaLinhaArray.push(novoIdGerado);
      } else {
        novaLinhaArray.push(dadosNovoProduto[nomeCabecalho] !== undefined ? dadosNovoProduto[nomeCabecalho] : "");
      }
    });
    
    abaProdutos.appendRow(novaLinhaArray);
    SpreadsheetApp.flush();
    return { success: true, message: "Produto criado com sucesso!", novoId: novoIdGerado };
  } catch (e) {
    console.error("ERRO em ProdutosCRUD_criarNovoProduto: " + e.toString() + " Stack: " + (e.stack || 'N/A'));
    return { success: false, message: e.message };
  }
}

/**
 * Atualiza um produto existente na planilha.
 * @param {object} dadosProdutoAtualizar Objeto contendo os dados do produto, incluindo o "ID".
 * @return {object} { success: boolean, message: string }
 */
function ProdutosCRUD_atualizarProduto(dadosProdutoAtualizar) {
  try {
    const idParaAtualizar = dadosProdutoAtualizar["ID"];
    if (!idParaAtualizar) throw new Error("ID do produto é obrigatório para atualização.");

    const nomeDoCampoProduto = CABECALHOS_PRODUTOS[ProdutosCRUD_IDX_PRODUTO_NOME];
    const nomeProdutoAtualizado = dadosProdutoAtualizar[nomeDoCampoProduto];
    if (!nomeProdutoAtualizado) throw new Error(`O campo '${nomeDoCampoProduto}' é obrigatório.`);
    
    const nomeProdutoAtualizadoNormalizado = ProdutosCRUD_normalizarTextoComparacao(nomeProdutoAtualizado);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaProdutos = ss.getSheetByName(ProdutosCRUD_ABA_PRODUTOS);
    if (!abaProdutos) throw new Error(`Aba '${ProdutosCRUD_ABA_PRODUTOS}' não encontrada.`);
    
    const range = abaProdutos.getDataRange();
    const todasAsLinhas = range.getValues();

    if (ProdutosCRUD_IDX_PRODUTO_ID === -1 || ProdutosCRUD_IDX_PRODUTO_NOME === -1) {
        throw new Error("Colunas 'ID' ou 'Produto' não encontradas na definição de cabeçalhos.");
    }
    
    let linhaParaAtualizarIndexNaPlanilha = -1;
    for (let i = 1; i < todasAsLinhas.length; i++) {
      const idLinhaAtual = String(todasAsLinhas[i][ProdutosCRUD_IDX_PRODUTO_ID]);
      if (idLinhaAtual === String(idParaAtualizar)) {
        linhaParaAtualizarIndexNaPlanilha = i;
      } else {
        const nomeExistenteNormalizado = ProdutosCRUD_normalizarTextoComparacao(String(todasAsLinhas[i][ProdutosCRUD_IDX_PRODUTO_NOME]));
        if (nomeExistenteNormalizado === nomeProdutoAtualizadoNormalizado) {
          throw new Error(`O nome de produto '${nomeProdutoAtualizado}' já está cadastrado para outro ID (${idLinhaAtual}).`);
        }
      }
    }

    if (linhaParaAtualizarIndexNaPlanilha === -1) {
      throw new Error(`Produto com ID '${idParaAtualizar}' não encontrado para atualização.`);
    }

    const linhaOriginalValores = todasAsLinhas[linhaParaAtualizarIndexNaPlanilha];
    const linhaAtualizadaValores = [];
    let alteracoesReais = 0;

    ProdutosCRUD_CABECALHOS_PRODUTOS.forEach((nomeCabecalho, k) => {
      if (nomeCabecalho === CABECALHOS_PRODUTOS[ProdutosCRUD_IDX_PRODUTO_DATA_CADASTRO] || nomeCabecalho === CABECALHOS_PRODUTOS[ProdutosCRUD_IDX_PRODUTO_ID]) {
        linhaAtualizadaValores.push(linhaOriginalValores[k]);
      } else {
        const valorNovo = dadosProdutoAtualizar[nomeCabecalho];
        const valorAntigo = linhaOriginalValores[k];
        const valorParaSalvar = valorNovo !== undefined ? valorNovo : valorAntigo;
        linhaAtualizadaValores.push(valorParaSalvar);
        let comparavelAntigo = typeof valorAntigo === 'string' ? ProdutosCRUD_normalizarTextoComparacao(String(valorAntigo)) : valorAntigo;
        let comparavelNovo = typeof valorParaSalvar === 'string' ? ProdutosCRUD_normalizarTextoComparacao(String(valorParaSalvar)) : valorParaSalvar;
        comparavelAntigo = (comparavelAntigo === null || comparavelAntigo === undefined) ? "" : String(comparavelAntigo);
        comparavelNovo = (comparavelNovo === null || comparavelNovo === undefined) ? "" : String(comparavelNovo);
        if (comparavelAntigo !== comparavelNovo) {
          alteracoesReais++;
        }
      }
    });

    if (alteracoesReais > 0) {
      abaProdutos.getRange(linhaParaAtualizarIndexNaPlanilha + 1, 1, 1, linhaAtualizadaValores.length).setValues([linhaAtualizadaValores]);
      SpreadsheetApp.flush();
      return { success: true, message: "Produto atualizado com sucesso!" };
    } else {
      return { success: true, message: "Nenhum dado foi modificado." };
    }
  } catch (e) {
    console.error("ERRO em ProdutosCRUD_atualizarProduto: " + e.toString() + " Stack: " + (e.stack || 'N/A'));
    return { success: false, message: e.message };
  }
}

/**
 * Obtém os subprodutos vinculados a um determinado nome de produto, incluindo todos os campos de CABECALHOS_SUBPRODUTOS.
 * @param {string} nomeProduto O nome do produto principal.
 * @return {Array<Object>} Array de objetos de subprodutos.
 */
function ProdutosCRUD_obterSubProdutosPorProduto(nomeProduto) {
  try {
    if (!nomeProduto) return [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSubProdutos = ss.getSheetByName(ProdutosCRUD_ABA_SUBPRODUTOS);
    if (!abaSubProdutos) {
      console.warn(`Aba de subprodutos '${ProdutosCRUD_ABA_SUBPRODUTOS}' não encontrada.`);
      return [];
    }
    const range = abaSubProdutos.getDataRange();
    const values = range.getValues();
    if (values.length <= 1) return []; // Aba vazia ou só com cabeçalhos

    const headersSubProdutosPlanilha = values[0].map(String); 
    const idxProdutoVinculadoPlanilha = headersSubProdutosPlanilha.indexOf(CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO]);
    const idxIdSubProdutoPlanilha = headersSubProdutosPlanilha.indexOf(CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_ID]);
    
    if (idxProdutoVinculadoPlanilha === -1) {
      throw new Error(`Coluna '${CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO]}' não encontrada na aba SubProdutos.`);
    }
    if (idxIdSubProdutoPlanilha === -1) {
        console.warn("Coluna 'ID' de Subproduto não encontrada na planilha. IDs podem não ser retornados corretamente.");
    }

    const subProdutosVinculados = [];
    const nomeProdutoNormalizado = ProdutosCRUD_normalizarTextoComparacao(nomeProduto);

    for (let i = 1; i < values.length; i++) {
      const produtoVinculadoAtual = values[i][idxProdutoVinculadoPlanilha];
      if (produtoVinculadoAtual && ProdutosCRUD_normalizarTextoComparacao(String(produtoVinculadoAtual)) === nomeProdutoNormalizado) {
        const subProdutoObj = {};
        // Mapeia usando os cabeçalhos definidos em Constantes.gs para consistência
        ProdutosCRUD_CABECALHOS_SUBPRODUTOS.forEach(headerConstante => {
            const idxNaPlanilha = headersSubProdutosPlanilha.indexOf(headerConstante);
            if (idxNaPlanilha !== -1) {
                 subProdutoObj[headerConstante] = values[i][idxNaPlanilha] !== null && values[i][idxNaPlanilha] !== undefined ? String(values[i][idxNaPlanilha]) : "";
            } else {
                subProdutoObj[headerConstante] = ""; // Garante que o objeto tenha todas as chaves esperadas
            }
        });
        subProdutosVinculados.push(subProdutoObj);
      }
    }
    return subProdutosVinculados;
  } catch (e) {
    console.error("Erro em ProdutosCRUD_obterSubProdutosPorProduto: " + e.toString());
    throw e;
  }
}

/**
 * Obtém a lista de outros produtos (exceto o que está sendo excluído).
 * @param {string} idProdutoExcluido O ID do produto que não deve ser incluído na lista.
 * @return {Array<Object>} Array de objetos de produtos, cada um com {id, nome}.
 */
function ProdutosCRUD_obterListaOutrosProdutos(idProdutoExcluido) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaProdutos = ss.getSheetByName(ProdutosCRUD_ABA_PRODUTOS);
    if (!abaProdutos) throw new Error(`Aba '${ProdutosCRUD_ABA_PRODUTOS}' não encontrada.`);
    const range = abaProdutos.getDataRange();
    const values = range.getValues();
    if (ProdutosCRUD_IDX_PRODUTO_ID === -1 || ProdutosCRUD_IDX_PRODUTO_NOME === -1) {
        throw new Error("Colunas 'ID' ou 'Produto' não encontradas na aba Produtos.");
    }
    const outrosProdutos = [];
    if (values.length > 1) {
      for (let i = 1; i < values.length; i++) {
        const idAtual = String(values[i][ProdutosCRUD_IDX_PRODUTO_ID]);
        if (idAtual !== String(idProdutoExcluido)) {
          outrosProdutos.push({ id: idAtual, nome: values[i][ProdutosCRUD_IDX_PRODUTO_NOME] });
        }
      }
    }
    return outrosProdutos;
  } catch (e) { console.error("Erro em ProdutosCRUD_obterListaOutrosProdutos: " + e.toString()); throw e; }
}

/**
 * Processa a exclusão de um produto, incluindo o tratamento de subprodutos vinculados.
 * @param {string} idProduto O ID do produto a ser excluído.
 * @param {string} nomeProdutoOriginal O nome original do produto.
 * @param {boolean} deletarSubprodutosVinculados Se true, os subprodutos são excluídos.
 * @param {Array<Object>|null} realocacoesSubprodutos Array de { subProdutoId: string, novoProdutoVinculadoNome: string }.
 * @return {object} { success: boolean, message: string }
 */
function ProdutosCRUD_processarExclusaoProduto(idProduto, nomeProdutoOriginal, deletarSubprodutosVinculados, realocacoesSubprodutos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaProdutos = ss.getSheetByName(ProdutosCRUD_ABA_PRODUTOS);
    const abaSubProdutos = ss.getSheetByName(ProdutosCRUD_ABA_SUBPRODUTOS);
    if (!abaProdutos) return { success: false, message: `Aba '${ProdutosCRUD_ABA_PRODUTOS}' não encontrada.` };
    
    const dataProdutos = abaProdutos.getDataRange().getValues();
    let produtoExcluido = false;
    for (let i = dataProdutos.length - 1; i >= 1; i--) {
      if (String(dataProdutos[i][ProdutosCRUD_IDX_PRODUTO_ID]) === String(idProduto)) {
        abaProdutos.deleteRow(i + 1);
        produtoExcluido = true;
        break;
      }
    }
    if (!produtoExcluido) return { success: false, message: `Produto com ID '${idProduto}' não encontrado.` };

    let mensagemFinal = `Produto '${nomeProdutoOriginal}' excluído.`;
    let subProdutosAfetadosCount = 0;

    if (abaSubProdutos) {
      if (ProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO === -1 || ProdutosCRUD_IDX_SUBPRODUTO_ID === -1) {
        console.warn("Colunas 'Produto Vinculado' ou 'ID' não encontradas em SubProdutos. Subprodutos não serão processados.");
      } else {
        const dataSubProdutos = abaSubProdutos.getDataRange().getValues();
        const nomeProdutoOriginalNormalizado = ProdutosCRUD_normalizarTextoComparacao(nomeProdutoOriginal);

        if (deletarSubprodutosVinculados) {
          for (let i = dataSubProdutos.length - 1; i >= 1; i--) {
            const produtoVinculadoAtual = dataSubProdutos[i][ProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO];
            if (produtoVinculadoAtual && ProdutosCRUD_normalizarTextoComparacao(String(produtoVinculadoAtual)) === nomeProdutoOriginalNormalizado) {
              abaSubProdutos.deleteRow(i + 1);
              subProdutosAfetadosCount++;
            }
          }
          if (subProdutosAfetadosCount > 0) mensagemFinal += ` ${subProdutosAfetadosCount} subprodutos vinculados foram excluídos.`;
        } else if (realocacoesSubprodutos && realocacoesSubprodutos.length > 0) {
          const subProdutoRowMap = {};
          for (let i = 1; i < dataSubProdutos.length; i++) {
            const subProdutoIdAtual = dataSubProdutos[i][ProdutosCRUD_IDX_SUBPRODUTO_ID];
            if(subProdutoIdAtual) subProdutoRowMap[String(subProdutoIdAtual)] = i + 1;
          }
          realocacoesSubprodutos.forEach(realocacao => {
            const linhaParaAtualizar = subProdutoRowMap[String(realocacao.subProdutoId)];
            if (linhaParaAtualizar && dataSubProdutos[linhaParaAtualizar-1][ProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO] &&
                ProdutosCRUD_normalizarTextoComparacao(String(dataSubProdutos[linhaParaAtualizar-1][ProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO])) === nomeProdutoOriginalNormalizado) {
              abaSubProdutos.getRange(linhaParaAtualizar, ProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO + 1).setValue(realocacao.novoProdutoVinculadoNome);
              subProdutosAfetadosCount++;
            }
          });
          if (subProdutosAfetadosCount > 0) mensagemFinal += ` ${subProdutosAfetadosCount} subprodutos realocados.`;
        }
      }
    } else if (deletarSubprodutosVinculados || (realocacoesSubprodutos && realocacoesSubprodutos.length > 0)) {
      mensagemFinal += ` (Aba SubProdutos '${ProdutosCRUD_ABA_SUBPRODUTOS}' não encontrada).`;
    }
    SpreadsheetApp.flush();
    return { success: true, message: mensagemFinal };
  } catch (e) {
    console.error("ERRO em ProdutosCRUD_processarExclusaoProduto: " + e.toString());
    return { success: false, message: "Falha ao excluir produto: " + e.message };
  }
}

/**
 * Obtém a lista completa de fornecedores para popular dropdowns.
 * @return {Array<Object>} Array de objetos de fornecedores, cada um com {id, nome}.
 */
function ProdutosCRUD_obterListaTodosFornecedores() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaFornecedores = ss.getSheetByName(ProdutosCRUD_ABA_FORNECEDORES);
    if (!abaFornecedores) {
      console.warn(`Aba de fornecedores '${ProdutosCRUD_ABA_FORNECEDORES}' não encontrada.`);
      return [];
    }
    const range = abaFornecedores.getDataRange();
    const values = range.getValues();
    if (ProdutosCRUD_IDX_FORNECEDOR_ID === -1 || ProdutosCRUD_IDX_FORNECEDOR_NOME === -1) {
        throw new Error("Colunas 'ID' ou 'Fornecedor' não encontradas na aba Fornecedores. Verifique Constantes.gs e os índices.");
    }
    const fornecedores = [];
    if (values.length > 1) {
      for (let i = 1; i < values.length; i++) {
        fornecedores.push({
          id: String(values[i][ProdutosCRUD_IDX_FORNECEDOR_ID]),
          nome: String(values[i][ProdutosCRUD_IDX_FORNECEDOR_NOME])
        });
      }
    }
    return fornecedores;
  } catch (e) {
    console.error("Erro em ProdutosCRUD_obterListaTodosFornecedores: " + e.toString());
    throw e;
  }
}

/**
 * Adiciona um novo subproduto vinculado a um produto principal.
 * @param {object} dadosSubProduto Objeto com os dados do novo subproduto.
 * @return {object} { success: boolean, message: string, novoId?: string }
 */
function ProdutosCRUD_adicionarNovoSubProdutoVinculado(dadosSubProduto) {
  try {
    const nomeCampoSubProduto = CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_NOME];
    const nomeCampoProdutoVinculado = CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO];

    if (!dadosSubProduto || !dadosSubProduto[nomeCampoSubProduto]) throw new Error(`O campo '${nomeCampoSubProduto}' é obrigatório.`);
    if (!dadosSubProduto[nomeCampoProdutoVinculado]) throw new Error(`O campo '${nomeCampoProdutoVinculado}' é obrigatório.`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSubProdutos = ss.getSheetByName(ProdutosCRUD_ABA_SUBPRODUTOS);
    if (!abaSubProdutos) throw new Error(`Aba '${ProdutosCRUD_ABA_SUBPRODUTOS}' não encontrada.`);

    const todasAsLinhasSubProdutos = abaSubProdutos.getDataRange().getValues();
    let proximoIdSubProduto = 1;
    if (ProdutosCRUD_IDX_SUBPRODUTO_ID === -1) throw new Error("Coluna 'ID' de SubProdutos não encontrada.");
    if (todasAsLinhasSubProdutos.length > 1) {
      const idsExistentes = todasAsLinhasSubProdutos.slice(1).map(linha => parseInt(linha[ProdutosCRUD_IDX_SUBPRODUTO_ID])).filter(id => !isNaN(id));
      if (idsExistentes.length > 0) proximoIdSubProduto = Math.max(...idsExistentes) + 1;
    }
    const novoIdGeradoSubProduto = String(proximoIdSubProduto);

    const novaLinhaSubProdutoArray = [];
    ProdutosCRUD_CABECALHOS_SUBPRODUTOS.forEach(nomeCabecalho => {
      if (nomeCabecalho === CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_DATA_CADASTRO]) {
        novaLinhaSubProdutoArray.push(new Date());
      } else if (nomeCabecalho === CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_ID]) {
        novaLinhaSubProdutoArray.push(novoIdGeradoSubProduto);
      } else {
        novaLinhaSubProdutoArray.push(dadosSubProduto[nomeCabecalho] !== undefined ? dadosSubProduto[nomeCabecalho] : "");
      }
    });

    abaSubProdutos.appendRow(novaLinhaSubProdutoArray);
    SpreadsheetApp.flush();
    return { success: true, message: "Subproduto vinculado adicionado!", novoId: novoIdGeradoSubProduto };
  } catch (e) {
    console.error("ERRO em ProdutosCRUD_adicionarNovoSubProdutoVinculado: " + e.toString());
    return { success: false, message: e.message };
  }
}

/**
 * Atualiza um subproduto existente na planilha.
 * @param {object} dadosSubProdutoAtualizar Objeto contendo os dados do subproduto, incluindo 'ID_SubProduto_Edicao'.
 * @return {object} { success: boolean, message: string }
 */
function ProdutosCRUD_atualizarSubProdutoVinculado(dadosSubProdutoAtualizar) {
  try {
    const idSubProdutoParaAtualizar = dadosSubProdutoAtualizar["ID_SubProduto_Edicao"];
    if (!idSubProdutoParaAtualizar) throw new Error("ID do Subproduto é obrigatório para atualização.");

    const nomeCampoSubProduto = CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_NOME];
    if (!dadosSubProdutoAtualizar[nomeCampoSubProduto]) {
        throw new Error(`O nome do Subproduto ('${nomeCampoSubProduto}') é obrigatório.`);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSubProdutos = ss.getSheetByName(ProdutosCRUD_ABA_SUBPRODUTOS);
    if (!abaSubProdutos) throw new Error(`Aba '${ProdutosCRUD_ABA_SUBPRODUTOS}' não encontrada.`);

    const range = abaSubProdutos.getDataRange();
    const todasAsLinhas = range.getValues();
    const cabecalhosDaPlanilha = todasAsLinhas[0].map(String);
    const idxIdNaPlanilha = cabecalhosDaPlanilha.indexOf(CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_ID]);
    if (idxIdNaPlanilha === -1) throw new Error("Coluna 'ID' não encontrada na aba SubProdutos.");

    let linhaParaAtualizarIndex = -1;
    for (let i = 1; i < todasAsLinhas.length; i++) {
      if (String(todasAsLinhas[i][idxIdNaPlanilha]) === String(idSubProdutoParaAtualizar)) {
        linhaParaAtualizarIndex = i;
        break;
      }
    }
    if (linhaParaAtualizarIndex === -1) {
      throw new Error(`Subproduto com ID '${idSubProdutoParaAtualizar}' não encontrado.`);
    }

    const linhaOriginalValores = todasAsLinhas[linhaParaAtualizarIndex];
    const linhaAtualizadaValores = [];
    let alteracoesReais = 0;

    // Mapeia os CABECALHOS_SUBPRODUTOS para os índices reais da planilha
    const mapaCabecalhosConstantesParaPlanilha = {};
    ProdutosCRUD_CABECALHOS_SUBPRODUTOS.forEach(headerConst => {
        mapaCabecalhosConstantesParaPlanilha[headerConst] = cabecalhosDaPlanilha.indexOf(headerConst);
    });

    ProdutosCRUD_CABECALHOS_SUBPRODUTOS.forEach(nomeCabecalhoConstante => {
      const indiceNaPlanilha = mapaCabecalhosConstantesParaPlanilha[nomeCabecalhoConstante];
      if (indiceNaPlanilha === -1) { // Se o cabeçalho da constante não existe na planilha, pula
          // Isso pode acontecer se CABECALHOS_SUBPRODUTOS tiver mais colunas que a planilha atual
          // Para colunas que DEVEM existir (como ID), a validação anterior já pegaria
          console.warn(`Cabeçalho '${nomeCabecalhoConstante}' de Constantes.gs não encontrado na planilha SubProdutos. Será ignorado na atualização.`);
          return; // continua para o próximo cabeçalho da constante
      }

      // Manter Data de Cadastro, ID e Produto Vinculado originais.
      if (nomeCabecalhoConstante === CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_DATA_CADASTRO] ||
          nomeCabecalhoConstante === CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_ID] ||
          nomeCabecalhoConstante === CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO]) {
        linhaAtualizadaValores[indiceNaPlanilha] = linhaOriginalValores[indiceNaPlanilha];
      } else if (dadosSubProdutoAtualizar.hasOwnProperty(nomeCabecalhoConstante)) {
        const valorNovo = dadosSubProdutoAtualizar[nomeCabecalhoConstante];
        const valorAntigo = linhaOriginalValores[indiceNaPlanilha];
        linhaAtualizadaValores[indiceNaPlanilha] = valorNovo;
        if (String(valorAntigo) !== String(valorNovo)) {
          alteracoesReais++;
        }
      } else {
        linhaAtualizadaValores[indiceNaPlanilha] = linhaOriginalValores[indiceNaPlanilha];
      }
    });
    
    // Preenche quaisquer lacunas se a linhaAtualizadaValores for menor que o número de colunas da planilha (improvável se iteramos por cabecalhosDaPlanilha)
    // Mais seguro é construir a linha baseada na ordem da planilha e substituir os valores.
    // Refazendo a lógica de construção da linha para garantir a ordem correta:
    const linhaFinalParaSalvar = [...linhaOriginalValores]; // Começa com uma cópia da linha original
    alteracoesReais = 0; // Reseta contador

    ProdutosCRUD_CABECALHOS_SUBPRODUTOS.forEach(nomeCabecalhoConstante => {
        const indiceNaPlanilha = cabecalhosDaPlanilha.indexOf(nomeCabecalhoConstante);
        if (indiceNaPlanilha === -1) return; // Pula se não existe na planilha

        if (nomeCabecalhoConstante === CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_DATA_CADASTRO] ||
            nomeCabecalhoConstante === CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_ID] ||
            nomeCabecalhoConstante === CABECALHOS_SUBPRODUTOS[ProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO]) {
            // Mantém os valores originais para estes campos
            linhaFinalParaSalvar[indiceNaPlanilha] = linhaOriginalValores[indiceNaPlanilha];
        } else if (dadosSubProdutoAtualizar.hasOwnProperty(nomeCabecalhoConstante)) {
            const valorNovo = dadosSubProdutoAtualizar[nomeCabecalhoConstante];
            const valorAntigo = linhaOriginalValores[indiceNaPlanilha];
            if (String(valorAntigo) !== String(valorNovo)) {
                alteracoesReais++;
            }
            linhaFinalParaSalvar[indiceNaPlanilha] = valorNovo;
        } else {
            // Mantém o valor original se não foi fornecido na atualização
            linhaFinalParaSalvar[indiceNaPlanilha] = linhaOriginalValores[indiceNaPlanilha];
        }
    });


    if (alteracoesReais > 0) {
      abaSubProdutos.getRange(linhaParaAtualizarIndex + 1, 1, 1, linhaFinalParaSalvar.length).setValues([linhaFinalParaSalvar]);
      SpreadsheetApp.flush();
      return { success: true, message: "Subproduto atualizado com sucesso!" };
    } else {
      return { success: true, message: "Nenhum dado do subproduto foi modificado." };
    }
  } catch (e) {
    console.error("ERRO em ProdutosCRUD_atualizarSubProdutoVinculado: " + e.toString());
    return { success: false, message: e.message };
  }
}