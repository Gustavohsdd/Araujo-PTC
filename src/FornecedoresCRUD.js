// Assume-se que ABA_FORNECEDORES, ABA_SUBPRODUTOS, CABECALHOS_FORNECEDORES, CABECALHOS_SUBPRODUTOS
// são definidos globalmente (ex: em um arquivo Constantes.gs)

const FornecedoresCRUD_ABA_FORNECEDORES_NOME = ABA_FORNECEDORES;
const FornecedoresCRUD_ABA_SUBPRODUTOS_NOME = ABA_SUBPRODUTOS;

// Índices para a aba Fornecedores
const FornecedoresCRUD_IDX_FORN_ID = CABECALHOS_FORNECEDORES.indexOf("ID");
const FornecedoresCRUD_IDX_FORN_NOME = CABECALHOS_FORNECEDORES.indexOf("Fornecedor");
// Adicione outros índices de Fornecedores se necessário para outras funções

// Índices para a aba SubProdutos (Itens) - Usaremos um conjunto completo para as novas funções CRUD de itens
const ITEM_IDX_DATA_CADASTRO = CABECALHOS_SUBPRODUTOS.indexOf("Data de Cadastro");
const ITEM_IDX_ID = CABECALHOS_SUBPRODUTOS.indexOf("ID"); // ID do Item/SubProduto
const ITEM_IDX_NOME = CABECALHOS_SUBPRODUTOS.indexOf("SubProduto"); // Nome do Item
const ITEM_IDX_PRODUTO_VINCULADO = CABECALHOS_SUBPRODUTOS.indexOf("Produto Vinculado"); // Produto principal ao qual o item pode estar associado
const ITEM_IDX_FORNECEDOR_VINCULADO = CABECALHOS_SUBPRODUTOS.indexOf("Fornecedor"); // Fornecedor do Item
const ITEM_IDX_CATEGORIA = CABECALHOS_SUBPRODUTOS.indexOf("Categoria");
const ITEM_IDX_TAMANHO = CABECALHOS_SUBPRODUTOS.indexOf("Tamanho");
const ITEM_IDX_UN = CABECALHOS_SUBPRODUTOS.indexOf("UN");
const ITEM_IDX_FATOR = CABECALHOS_SUBPRODUTOS.indexOf("Fator");
const ITEM_IDX_NCM = CABECALHOS_SUBPRODUTOS.indexOf("NCM");
const ITEM_IDX_CST = CABECALHOS_SUBPRODUTOS.indexOf("CST");
const ITEM_IDX_CFOP = CABECALHOS_SUBPRODUTOS.indexOf("CFOP");
const ITEM_IDX_STATUS = CABECALHOS_SUBPRODUTOS.indexOf("Status");
// Adicione outros índices de SubProdutos conforme os cabeçalhos em CABECALHOS_SUBPRODUTOS

/**
 * Converte uma linha da planilha de subprodutos (array) para um objeto.
 * @param {Array} rowArray Array de valores da linha.
 * @param {Array} headers Array dos cabeçalhos da planilha de subprodutos.
 * @return {Object} Objeto representando o subproduto/item.
 */
function FornecedoresCRUD_mapSubProdutoRowToObject(rowArray, headers) {
  const obj = {};
  headers.forEach((header, index) => {
    obj[header] = rowArray[index];
  });
  // Garante que o ID do subproduto seja tratado corretamente, pode ser nomeado como 'ID' ou 'ID_SubProduto' na sua planilha.
  // O frontend espera 'ID_SubProduto' para edição e listagem individual.
  // E 'SubProduto' para o nome.
  obj['ID_SubProduto'] = rowArray[ITEM_IDX_ID];
  obj['SubProduto'] = rowArray[ITEM_IDX_NOME];
  return obj;
}


// --- Funções CRUD para Itens de Fornecedores ---

/**
 * Cria um novo item (subproduto) e o vincula a um fornecedor e, opcionalmente, a um produto principal.
 * Chamado por: FornecedoresScript_handleSalvarItemFornecedor (modo criação)
 * dadosItem: Objeto contendo os dados do formulário do item.
 */
function SubProdutosCRUD_criarNovoSubProduto(dadosItem) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. VALIDAÇÃO DO FORNECEDOR (Já está correta)
    const nomeFornecedor = dadosItem["Fornecedor"];
    if (!nomeFornecedor || String(nomeFornecedor).trim() === "") {
      return { success: false, message: "O nome do fornecedor é obrigatório. Ocorreu um erro ao obter o fornecedor selecionado." };
    }
    const abaFornecedores = ss.getSheetByName(ABA_FORNECEDORES);
    if (!abaFornecedores) {
      return { success: false, message: `Aba de validação '${ABA_FORNECEDORES}' não encontrada.` };
    }
    const valoresFornecedores = abaFornecedores.getDataRange().getValues();
    const idxColunaNomeFornecedor = CABECALHOS_FORNECEDORES.indexOf("Fornecedor");
    let fornecedorEncontrado = false;
    for (let i = 1; i < valoresFornecedores.length; i++) {
      if (valoresFornecedores[i][idxColunaNomeFornecedor] === nomeFornecedor) {
        fornecedorEncontrado = true;
        break;
      }
    }
    if (!fornecedorEncontrado) {
      return { success: false, message: `O fornecedor '${nomeFornecedor}' não foi encontrado no cadastro. Verifique a lista de fornecedores.` };
    }

    // 2. VALIDAÇÃO DO PRODUTO VINCULADO (Já está correta)
    const idProdutoVinculado = dadosItem["Produto Vinculado"];
    let nomeProdutoParaSalvar = ""; 
    if (idProdutoVinculado && String(idProdutoVinculado).trim() !== "") {
      const abaProdutos = ss.getSheetByName(ABA_PRODUTOS);
      if (!abaProdutos) {
        return { success: false, message: `Aba de validação '${ABA_PRODUTOS}' não encontrada.` };
      }
      const valoresProdutos = abaProdutos.getDataRange().getValues();
      const idxColunaIdProduto = CABECALHOS_PRODUTOS.indexOf("ID");
      const idxColunaNomeProduto = CABECALHOS_PRODUTOS.indexOf("Produto");
      let produtoEncontrado = false;
      for (let i = 1; i < valoresProdutos.length; i++) {
        if (String(valoresProdutos[i][idxColunaIdProduto]) === String(idProdutoVinculado)) {
          produtoEncontrado = true;
          nomeProdutoParaSalvar = valoresProdutos[i][idxColunaNomeProduto];
          break;
        }
      }
      if (!produtoEncontrado) {
        return { success: false, message: `Produto Vinculado com ID '${idProdutoVinculado}' não encontrado.` };
      }
    }

    // 3. CRIAÇÃO E INSERÇÃO DO SUBPRODUTO
    const abaSubProdutos = ss.getSheetByName(FornecedoresCRUD_ABA_SUBPRODUTOS_NOME);
    if (!abaSubProdutos) {
      return { success: false, message: `Aba '${FornecedoresCRUD_ABA_SUBPRODUTOS_NOME}' não encontrada.` };
    }

    const novoId = Utils_gerarProximoId(abaSubProdutos, ITEM_IDX_ID);
    const dataCadastro = new Date();

    // LÓGICA DE MONTAGEM DA LINHA CORRIGIDA
    // Esta abordagem é mais segura pois monta a linha explicitamente na ordem dos cabeçalhos.
    const novaLinha = CABECALHOS_SUBPRODUTOS.map(header => {
      switch (header) {
        case "ID":
          return novoId;
        case "Data de Cadastro":
          return dataCadastro;
        case "Fornecedor":
          return nomeFornecedor; // AQUI ESTÁ A CORREÇÃO PRINCIPAL: Garante que o fornecedor correto seja salvo
        case "Produto Vinculado":
          return nomeProdutoParaSalvar; // Salva o nome do produto, não o ID
        default:
          // Pega todos os outros valores que vieram do formulário (SubProduto, Categoria, NCM, etc.)
          return dadosItem[header] || ""; 
      }
    });

    abaSubProdutos.appendRow(novaLinha);
    SpreadsheetApp.flush();
    return { success: true, message: "Item adicionado com sucesso!", newItemId: novoId };

  } catch (e) {
    console.error("Erro em SubProdutosCRUD_criarNovoSubProduto: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: "Falha ao criar item: " + e.message };
  }
}

/**
 * (Substitui/Complementa SubProdutosCRUD_atualizarSubProduto)
 * Atualiza um item (subproduto) existente.
 * Chamado por: FornecedoresScript_handleSalvarItemFornecedor (modo edição)
 * dadosItem: Objeto contendo os dados do formulário, incluindo 'ID_SubProduto_Edicao'.
 */
function SubProdutosCRUD_atualizarSubProduto(dadosItem) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSubProdutos = ss.getSheetByName(FornecedoresCRUD_ABA_SUBPRODUTOS_NOME);
    if (!abaSubProdutos) return { success: false, message: `Aba '${FornecedoresCRUD_ABA_SUBPRODUTOS_NOME}' não encontrada.` };

    const itemIdParaEditar = dadosItem.ID_SubProduto_Edicao;
    if (!itemIdParaEditar) return { success: false, message: "ID do item para edição não fornecido." };

    const range = abaSubProdutos.getDataRange();
    const values = range.getValues();
    const headers = values[0];
    let linhaEncontrada = -1;

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][ITEM_IDX_ID]) === String(itemIdParaEditar)) {
        linhaEncontrada = i + 1; // Linha da planilha (1-indexed)
        break;
      }
    }

    if (linhaEncontrada === -1) {
      return { success: false, message: `Item com ID '${itemIdParaEditar}' não encontrado para atualização.` };
    }

    // VALIDAÇÃO DO PRODUTO VINCULADO (se foi alterado)
    const idProdutoVinculado = dadosItem["Produto Vinculado"];
    let nomeProdutoParaSalvar = ""; 
    if (idProdutoVinculado && String(idProdutoVinculado).trim() !== "") {
      const abaProdutos = ss.getSheetByName(ABA_PRODUTOS);
      if (!abaProdutos) return { success: false, message: `Aba de validação '${ABA_PRODUTOS}' não encontrada.` };
      
      const valoresProdutos = abaProdutos.getDataRange().getValues();
      const idxColunaIdProduto = CABECALHOS_PRODUTOS.indexOf("ID");
      const idxColunaNomeProduto = CABECALHOS_PRODUTOS.indexOf("Produto");
      let produtoEncontrado = false;
      for (let i = 1; i < valoresProdutos.length; i++) {
        if (String(valoresProdutos[i][idxColunaIdProduto]) === String(idProdutoVinculado)) {
          produtoEncontrado = true;
          nomeProdutoParaSalvar = valoresProdutos[i][idxColunaNomeProduto];
          break;
        }
      }
      if (!produtoEncontrado) return { success: false, message: `Produto Vinculado com ID '${idProdutoVinculado}' não encontrado.` };
    }
    
    // Atualiza os valores na linha encontrada
    const linhaOriginal = values[linhaEncontrada - 1];
    const linhaAtualizada = linhaOriginal.map((celula, index) => {
      const header = headers[index];
      
      if (header === "ID" || header === "Data de Cadastro") {
        return celula; // Não atualiza ID nem Data de Cadastro
      }
      
      if (header === "Fornecedor") {
        // O fornecedor é definido pelo contexto do modal, não deve ser alterado aqui.
        // Usa o valor que já está na planilha.
        return celula; 
      }
      
      if (header === "Produto Vinculado") {
        return nomeProdutoParaSalvar; // Salva o novo nome do produto ou uma string vazia se foi desmarcado
      }

      // Para todos os outros campos, verifica se um novo valor foi enviado pelo formulário
      if (dadosItem.hasOwnProperty(header)) {
        return dadosItem[header];
      }
      
      // Se não veio do formulário, mantém o valor original
      return celula;
    });

    abaSubProdutos.getRange(linhaEncontrada, 1, 1, linhaAtualizada.length).setValues([linhaAtualizada]);
    SpreadsheetApp.flush();
    return { success: true, message: "Item atualizado com sucesso." };
  } catch (e) {
    console.error("Erro em SubProdutosCRUD_atualizarSubProduto: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: "Falha ao atualizar item: " + e.message };
  }
}

/**
 * Obtém todos os itens (subprodutos) vinculados a um nome de fornecedor específico, com detalhes completos.
 * Chamado por: FornecedoresScript_carregarEListarItensFornecedor
 */
function SubProdutosCRUD_obterSubProdutosPorPai(nomePai, tipoPai) {
  try {
    if (tipoPai !== 'FORNECEDOR' && tipoPai !== 'PRODUTO') {
      throw new Error("Tipo de pai inválido. Deve ser 'FORNECEDOR' ou 'PRODUTO'.");
    }
    if (!nomePai) {
        return []; // Retorna um array vazio se nenhum nome de pai for fornecido
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSubProdutos = ss.getSheetByName(FornecedoresCRUD_ABA_SUBPRODUTOS_NOME);
    if (!abaSubProdutos) throw new Error(`Aba '${FornecedoresCRUD_ABA_SUBPRODUTOS_NOME}' não encontrada.`);

    const range = abaSubProdutos.getDataRange();
    const values = range.getValues();
    const headers = values[0];

    if (ITEM_IDX_ID === -1 || ITEM_IDX_NOME === -1 || ITEM_IDX_FORNECEDOR_VINCULADO === -1 || ITEM_IDX_PRODUTO_VINCULADO === -1) {
      throw new Error("Colunas essenciais (ID, SubProduto, Fornecedor, Produto Vinculado) não encontradas na aba SubProdutos.");
    }

    const itensVinculados = [];
    const colunaFiltroIdx = tipoPai === 'FORNECEDOR' ? ITEM_IDX_FORNECEDOR_VINCULADO : ITEM_IDX_PRODUTO_VINCULADO;
    
    // CORREÇÃO: Normaliza os nomes antes de comparar para evitar problemas com espaços
    const nomePaiNormalizado = String(nomePai).trim();

    for (let i = 1; i < values.length; i++) {
      const valorCelulaNormalizado = String(values[i][colunaFiltroIdx]).trim();
      
      if (valorCelulaNormalizado === nomePaiNormalizado) {
        itensVinculados.push(FornecedoresCRUD_mapSubProdutoRowToObject(values[i], headers));
      }
    }
    return itensVinculados;
  } catch (e) {
    console.error(`Erro em SubProdutosCRUD_obterSubProdutosPorPai (tipo: ${tipoPai}): ` + e.toString() + " Stack: " + e.stack);
    throw e; 
  }
}

/**
 * (Substitui/Complementa SubProdutosCRUD_obterDetalhesSubProdutoPorId)
 * Obtém os detalhes completos de um item (subproduto) específico pelo seu ID.
 * Chamado por: FornecedoresScript_prepararFormParaEditarItemFornecedor
 */
function SubProdutosCRUD_obterDetalhesSubProdutoPorId(itemId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSubProdutos = ss.getSheetByName(FornecedoresCRUD_ABA_SUBPRODUTOS_NOME);
    if (!abaSubProdutos) throw new Error(`Aba '${FornecedoresCRUD_ABA_SUBPRODUTOS_NOME}' não encontrada.`);

    const range = abaSubProdutos.getDataRange();
    const values = range.getValues();
    const headers = values[0];

    if (ITEM_IDX_ID === -1) {
      throw new Error("Coluna 'ID' não encontrada na aba SubProdutos.");
    }

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][ITEM_IDX_ID]) === String(itemId)) {
        return FornecedoresCRUD_mapSubProdutoRowToObject(values[i], headers);
      }
    }
    return null; // Item não encontrado
  } catch (e) {
    console.error("Erro em SubProdutosCRUD_obterDetalhesSubProdutoPorId: " + e.toString());
    throw e;
  }
}

/**
 * (Substitui/Complementa SubProdutosCRUD_excluirSubProduto)
 * Exclui um item (subproduto) específico.
 * Chamado por: FornecedoresScript_confirmarExcluirItemFornecedor
 */
function SubProdutosCRUD_excluirSubProduto(itemId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSubProdutos = ss.getSheetByName(FornecedoresCRUD_ABA_SUBPRODUTOS_NOME);
    if (!abaSubProdutos) return { success: false, message: `Aba '${FornecedoresCRUD_ABA_SUBPRODUTOS_NOME}' não encontrada.` };

    const range = abaSubProdutos.getDataRange();
    const values = range.getValues();
    let itemExcluido = false;

    for (let i = values.length - 1; i >= 1; i--) { // Itera de baixo para cima para evitar problemas com índices ao deletar
      if (String(values[i][ITEM_IDX_ID]) === String(itemId)) {
        abaSubProdutos.deleteRow(i + 1);
        itemExcluido = true;
        break;
      }
    }

    if (itemExcluido) {
      SpreadsheetApp.flush();
      return { success: true, message: "Item excluído com sucesso." };
    } else {
      return { success: false, message: `Item com ID '${itemId}' não encontrado para exclusão.` };
    }
  } catch (e) {
    console.error("Erro em SubProdutosCRUD_excluirSubProduto: " + e.toString());
    return { success: false, message: "Falha ao excluir item: " + e.message };
  }
}

/**
 * Função para buscar nomes e IDs de produtos principais.
 * Idealmente, esta função estaria em ProdutosCRUD.gs, mas é chamada pelo FornecedoresScript.
 * Chamado por: FornecedoresScript_carregarProdutosParaSelectNoModalItem
 */
function ProdutosCRUD_obterNomesEIdsProdutos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaProdutos = ss.getSheetByName(ABA_PRODUTOS);
    if (!abaProdutos) throw new Error(`Aba de Produtos ('${ABA_PRODUTOS}') não encontrada.`);

    const range = abaProdutos.getDataRange();
    const values = range.getValues();
    
    const idxIdProduto = CABECALHOS_PRODUTOS.indexOf("ID");
    const idxNomeProduto = CABECALHOS_PRODUTOS.indexOf("Produto");
    const idxStatusProduto = CABECALHOS_PRODUTOS.indexOf("Status");

    if (idxIdProduto === -1 || idxNomeProduto === -1) {
      throw new Error("Colunas 'ID' ou 'Produto' não encontradas na aba Produtos.");
    }

    const produtos = [];
    for (let i = 1; i < values.length; i++) {
      if (values[i][idxIdProduto] && values[i][idxNomeProduto]) {
        let isAtivo = (idxStatusProduto === -1) || (values[i][idxStatusProduto].toString().toLowerCase() === 'ativo');
        if (isAtivo) {
          produtos.push({
              id: values[i][idxIdProduto],
              nome: values[i][idxNomeProduto]
          });
        }
      }
    }

    // CORREÇÃO: Ordena a lista de produtos em ordem alfabética pelo nome
    produtos.sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR', { sensitivity: 'base' }));
    
    return produtos;

  } catch (e) {
    console.error("Erro em ProdutosCRUD_obterNomesEIdsProdutos: " + e.toString());
    return []; 
  }
}


// --- Funções existentes em FornecedoresCRUD.gs (revisadas para clareza) ---

/**
 * Obtém uma lista simplificada de subprodutos (id e nome) vinculados a um fornecedor.
 * Usado principalmente para a verificação durante a exclusão do fornecedor.
 */
function FornecedoresCRUD_obterSubProdutosPorNomeFornecedor(nomeFornecedor) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSubProdutos = ss.getSheetByName(FornecedoresCRUD_ABA_SUBPRODUTOS_NOME);
    if (!abaSubProdutos) throw new Error(`Aba '${FornecedoresCRUD_ABA_SUBPRODUTOS_NOME}' não encontrada.`);

    const range = abaSubProdutos.getDataRange();
    const values = range.getValues();

    if (ITEM_IDX_ID === -1 || ITEM_IDX_NOME === -1 || ITEM_IDX_FORNECEDOR_VINCULADO === -1) {
      throw new Error("Colunas essenciais ('ID', 'SubProduto', 'Fornecedor') não encontradas na aba SubProdutos.");
    }

    const subProdutosVinculados = [];
    for (let i = 1; i < values.length; i++) { // Começa de 1 para pular cabeçalho
      if (values[i][ITEM_IDX_FORNECEDOR_VINCULADO] === nomeFornecedor) {
        subProdutosVinculados.push({
          id: values[i][ITEM_IDX_ID],       // ID do SubProduto/Item
          nome: values[i][ITEM_IDX_NOME]   // Nome do SubProduto/Item
        });
      }
    }
    return subProdutosVinculados;
  } catch (e) {
    console.error("Erro em FornecedoresCRUD_obterSubProdutosPorNomeFornecedor: " + e.toString());
    throw e;
  }
}

function FornecedoresCRUD_obterListaOutrosFornecedores(idFornecedorExcluido) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaFornecedores = ss.getSheetByName(FornecedoresCRUD_ABA_FORNECEDORES_NOME);
    if (!abaFornecedores) throw new Error(`Aba '${FornecedoresCRUD_ABA_FORNECEDORES_NOME}' não encontrada.`);

    const range = abaFornecedores.getDataRange();
    const values = range.getValues();
    
    if (FornecedoresCRUD_IDX_FORN_ID === -1 || FornecedoresCRUD_IDX_FORN_NOME === -1) {
        throw new Error("Colunas 'ID' ou 'Fornecedor' não encontradas na aba Fornecedores.");
    }

    const outrosFornecedores = [];
    for (let i = 1; i < values.length; i++) { // Começa de 1 para pular cabeçalho
      const idAtual = String(values[i][FornecedoresCRUD_IDX_FORN_ID]);
      if (idAtual !== String(idFornecedorExcluido)) {
        outrosFornecedores.push({
          id: idAtual,
          nome: values[i][FornecedoresCRUD_IDX_FORN_NOME]
        });
      }
    }
    return outrosFornecedores;
  } catch (e) {
    console.error("Erro em FornecedoresCRUD_obterListaOutrosFornecedores: " + e.toString());
    throw e;
  }
}

function FornecedoresCRUD_processarExclusaoFornecedor(idFornecedor, nomeFornecedorOriginal, deletarSubprodutosVinculados, realocacoesSubprodutos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaFornecedores = ss.getSheetByName(FornecedoresCRUD_ABA_FORNECEDORES_NOME);
    const abaSubProdutos = ss.getSheetByName(FornecedoresCRUD_ABA_SUBPRODUTOS_NOME);

    if (!abaFornecedores) return { success: false, message: `Aba '${FornecedoresCRUD_ABA_FORNECEDORES_NOME}' não encontrada.` };
    if (!abaSubProdutos && (deletarSubprodutosVinculados || (realocacoesSubprodutos && realocacoesSubprodutos.length > 0))) {
      return { success: false, message: `Aba '${FornecedoresCRUD_ABA_SUBPRODUTOS_NOME}' não encontrada.` };
    }
    
    const dataFornecedores = abaFornecedores.getDataRange().getValues();
    let fornecedorExcluido = false;
    for (let i = dataFornecedores.length - 1; i >= 1; i--) { // Itera de baixo para cima
      if (String(dataFornecedores[i][FornecedoresCRUD_IDX_FORN_ID]) === String(idFornecedor)) {
        abaFornecedores.deleteRow(i + 1);
        fornecedorExcluido = true;
        break;
      }
    }
    if (!fornecedorExcluido) {
      return { success: false, message: `Fornecedor com ID '${idFornecedor}' não encontrado para exclusão.` };
    }

    let mensagemFinal = `Fornecedor '${nomeFornecedorOriginal}' excluído com sucesso.`;

    if (abaSubProdutos) {
        if (deletarSubprodutosVinculados) {
          const dataSubProdutos = abaSubProdutos.getDataRange().getValues();
          let subProdutosDeletadosCount = 0;
          for (let i = dataSubProdutos.length - 1; i >= 1; i--) { // Itera de baixo para cima
            if (dataSubProdutos[i][ITEM_IDX_FORNECEDOR_VINCULADO] === nomeFornecedorOriginal) {
              abaSubProdutos.deleteRow(i + 1);
              subProdutosDeletadosCount++;
            }
          }
          if (subProdutosDeletadosCount > 0) mensagemFinal += ` ${subProdutosDeletadosCount} itens vinculados foram excluídos.`;
        } else if (realocacoesSubprodutos && realocacoesSubprodutos.length > 0) {
          const dataSubProdutos = abaSubProdutos.getDataRange().getValues();
          let subProdutosAtualizadosCount = 0;

          const subProdutoRowMap = {}; // Mapeia ID do subproduto para seu índice de linha
          for (let i = 1; i < dataSubProdutos.length; i++) {
            const subProdutoIdAtual = dataSubProdutos[i][ITEM_IDX_ID];
            if(subProdutoIdAtual) {
                subProdutoRowMap[String(subProdutoIdAtual)] = i + 1; // linha 1-indexed
            }
          }
          
          realocacoesSubprodutos.forEach(realocacao => {
            const linhaParaAtualizar = subProdutoRowMap[String(realocacao.subProdutoId)];
            // Verifica se o subproduto realmente pertencia ao fornecedor original antes de realocar
            if (linhaParaAtualizar && dataSubProdutos[linhaParaAtualizar-1][ITEM_IDX_FORNECEDOR_VINCULADO] === nomeFornecedorOriginal) {
              abaSubProdutos.getRange(linhaParaAtualizar, ITEM_IDX_FORNECEDOR_VINCULADO + 1).setValue(realocacao.novoFornecedorNome);
              subProdutosAtualizadosCount++;
            }
          });
          if (subProdutosAtualizadosCount > 0) mensagemFinal += ` ${subProdutosAtualizadosCount} itens realocados.`;
        }
    }

    SpreadsheetApp.flush();
    return { success: true, message: mensagemFinal };

  } catch (e) {
    console.error("ERRO em FornecedoresCRUD_processarExclusaoFornecedor: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: "Falha ao processar exclusão do fornecedor. Detalhes: " + e.message };
  }
}

/**
 * Função utilitária para gerar o próximo ID sequencial para uma coluna específica.
 * @param {Sheet} sheet A planilha onde o ID será gerado.
 * @param {number} idColumnIndex O índice da coluna do ID (0-indexed).
 * @return {number} O próximo ID.
 */
function Utils_gerarProximoId(sheet, idColumnIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 1; // Se a planilha estiver vazia (sem cabeçalhos sequer)
  
  const range = sheet.getRange(1, idColumnIndex + 1, lastRow, 1);
  const values = range.getValues();
  let maxId = 0;
  
  // Começa do final para pegar IDs mais recentes, assumindo que podem não estar ordenados
  // ou que a primeira linha é cabeçalho
  for (let i = values.length - 1; i >= (sheet.getFrozenRows() > 0 ? sheet.getFrozenRows() : 1) ; i--) {
    const currentId = parseInt(values[i][0], 10);
    if (!isNaN(currentId) && currentId > maxId) {
      maxId = currentId;
    }
  }
  // Se não encontrou nenhum ID numérico válido (ex: só cabeçalho ou planilha vazia de dados)
  if (maxId === 0 && lastRow > 0 && values.length > 0) {
      // Tenta pegar o último valor da coluna de ID, mesmo que seja o cabeçalho, se for o único
      const potentialLastId = parseInt(sheet.getRange(lastRow, idColumnIndex + 1).getValue(), 10);
      if (!isNaN(potentialLastId)) {
          maxId = Math.max(maxId, potentialLastId);
      }
  }
  return maxId + 1;
}

// ===================================================================================
// =================== NOVAS FUNÇÕES CONSTRUÍDAS DO ZERO ============================
// ===================================================================================

/**
 * [NOVA FUNÇÃO] Mapeia uma linha de dados da planilha para um objeto JS limpo.
 * Lida explicitamente com datas para evitar erros de serialização.
 * @param {Array} rowArray A linha de dados da planilha.
 * @param {Array} headers O array de cabeçalhos da planilha.
 * @return {Object} Um objeto representando o SubProduto.
 */
function FornecedoresCRUD_mapSubProdutoRowToObject_NOVO(rowArray, headers) {
  const obj = {};
  headers.forEach((header, index) => {
    let cellValue = rowArray[index];
    // Tratamento explícito de datas para garantir que sejam enviadas como strings formatadas
    if (cellValue instanceof Date) {
      if (cellValue.getHours() > 0 || cellValue.getMinutes() > 0 || cellValue.getSeconds() > 0) {
        obj[header] = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      } else {
        obj[header] = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
    } else {
      obj[header] = cellValue;
    }
  });

  // Adiciona as propriedades que o lado do cliente (frontend) espera
  obj['ID_SubProduto'] = rowArray[ITEM_IDX_ID];
  obj['SubProduto'] = rowArray[ITEM_IDX_NOME];
  
  return obj;
}


/**
 * [NOVA FUNÇÃO] Cria um novo SubProduto na planilha.
 * Esta função constrói a linha de forma manual e explícita para máxima confiabilidade.
 * @param {Object} dadosItem Dados vindos do formulário do frontend.
 * @return {Object} Resultado da operação {success: boolean, message: string}.
 */
function SubProdutosCRUD_criarNovoSubProduto_NOVO(dadosItem) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSubProdutos = ss.getSheetByName(ABA_SUBPRODUTOS);
    if (!abaSubProdutos) {
      return { success: false, message: `A aba de destino '${ABA_SUBPRODUTOS}' não foi encontrada.` };
    }

    // --- 1. Validar e preparar os dados essenciais ---
    const nomeFornecedor = dadosItem["Fornecedor"];
    if (!nomeFornecedor) {
      return { success: false, message: "O nome do Fornecedor é obrigatório e não foi recebido." };
    }
    
    const nomeSubProduto = dadosItem["SubProduto"];
    if (!nomeSubProduto) {
      return { success: false, message: "O nome do SubProduto é obrigatório." };
    }

    // --- CORREÇÃO: LÓGICA PARA BUSCAR O NOME DO PRODUTO VINCULADO ---
    const idProdutoVinculado = dadosItem["Produto Vinculado"];
    let nomeProdutoParaSalvar = ""; // Inicia como vazio
    
    if (idProdutoVinculado && String(idProdutoVinculado).trim() !== "") {
      const abaProdutos = ss.getSheetByName(ABA_PRODUTOS);
      if (!abaProdutos) {
        return { success: false, message: `Aba de validação '${ABA_PRODUTOS}' não encontrada.` };
      }
      const valoresProdutos = abaProdutos.getDataRange().getValues();
      const idxColunaIdProduto = CABECALHOS_PRODUTOS.indexOf("ID");
      const idxColunaNomeProduto = CABECALHOS_PRODUTOS.indexOf("Produto");
      
      let produtoEncontrado = false;
      for (let i = 1; i < valoresProdutos.length; i++) {
        if (String(valoresProdutos[i][idxColunaIdProduto]) === String(idProdutoVinculado)) {
          nomeProdutoParaSalvar = valoresProdutos[i][idxColunaNomeProduto]; // Armazena o NOME
          produtoEncontrado = true;
          break;
        }
      }
      if (!produtoEncontrado) {
        return { success: false, message: `O Produto Vinculado com ID '${idProdutoVinculado}' não foi encontrado.` };
      }
    }

    // --- 2. Construir a nova linha na ordem EXATA dos cabeçalhos ---
    const novaLinha = [];
    const novoId = Utils_gerarProximoId(abaSubProdutos, ITEM_IDX_ID);
    const dataCadastro = new Date();

    for (const header of CABECALHOS_SUBPRODUTOS) {
      let valorParaColuna = "";

      switch (header) {
        case "Data de Cadastro":
          valorParaColuna = dataCadastro;
          break;
        case "ID":
          valorParaColuna = novoId;
          break;
        case "Fornecedor":
          valorParaColuna = nomeFornecedor;
          break;
        case "Produto Vinculado":
          valorParaColuna = nomeProdutoParaSalvar; // Ponto CRÍTICO: Usa a variável com o nome do produto
          break;
        default:
          if (dadosItem.hasOwnProperty(header)) {
            valorParaColuna = dadosItem[header];
          }
          break;
      }
      novaLinha.push(valorParaColuna);
    }
    
    // --- 3. Inserir a linha na planilha ---
    abaSubProdutos.appendRow(novaLinha);
    SpreadsheetApp.flush();
    
    return { success: true, message: "SubProduto criado com sucesso!" };

  } catch (e) {
    console.error("Erro em SubProdutosCRUD_criarNovoSubProduto_NOVO: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: "Falha crítica ao criar o SubProduto: " + e.message };
  }
}

/**
 * [NOVA FUNÇÃO] Obtém todos os SubProdutos de um fornecedor.
 * Esta função normaliza os nomes para comparação e usa a nova função de mapeamento.
 * @param {string} nomePai O nome do Fornecedor a ser buscado.
 * @param {string} tipoPai Deve ser 'FORNECEDOR'.
 * @return {Array<Object>} Uma lista de objetos de subprodutos.
 */
function SubProdutosCRUD_obterSubProdutosPorPai_NOVO(nomePai, tipoPai) {
  try {
    if (tipoPai !== 'FORNECEDOR') {
      return []; // Apenas responde a buscas por fornecedor
    }
    if (!nomePai) {
      return []; // Retorna vazio se nenhum nome for fornecido
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSubProdutos = ss.getSheetByName(ABA_SUBPRODUTOS);
    if (!abaSubProdutos) {
      throw new Error(`Aba '${ABA_SUBPRODUTOS}' não foi encontrada.`);
    }

    const range = abaSubProdutos.getDataRange();
    const allValues = range.getValues();
    const headers = allValues[0];
    const dataRows = allValues.slice(1);
    
    const itensVinculados = [];
    const nomePaiNormalizado = String(nomePai).trim(); // Normaliza o nome do fornecedor buscado

    // Pega o índice da coluna "Fornecedor" a partir das constantes
    const idxFornecedor = ITEM_IDX_FORNECEDOR_VINCULADO;
    if (idxFornecedor === -1) {
      throw new Error("A coluna 'Fornecedor' não foi encontrada nos cabeçalhos da aba SubProdutos.");
    }

    for (const row of dataRows) {
      // Normaliza o nome do fornecedor da planilha antes de comparar
      const fornecedorDaLinha = String(row[idxFornecedor]).trim();
      
      if (fornecedorDaLinha === nomePaiNormalizado) {
        // Usa a nova função de mapeamento para criar o objeto
        itensVinculados.push(FornecedoresCRUD_mapSubProdutoRowToObject_NOVO(row, headers));
      }
    }

    return itensVinculados;
    
  } catch (e) {
    console.error("Erro em SubProdutosCRUD_obterSubProdutosPorPai_NOVO: " + e.toString() + " Stack: " + e.stack);
    // Retorna o erro para o frontend para que a falha seja visível
    return { error: true, message: e.message };
  }
}