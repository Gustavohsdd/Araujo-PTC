// @ts-nocheck
// Arquivo: SubProdutosCRUD.gs

/**
 * @OnlyCurrentDoc
 */

// Constantes da Planilha (definidas em Constantes.gs e acessíveis globalmente)
const SubProdutosCRUD_ABA_SUBPRODUTOS = ABA_SUBPRODUTOS;
const SubProdutosCRUD_CABECALHOS_SUBPRODUTOS = CABECALHOS_SUBPRODUTOS;

const SubProdutosCRUD_ABA_PRODUTOS = ABA_PRODUTOS; // Para buscar nomes de produtos
const SubProdutosCRUD_CABECALHOS_PRODUTOS = CABECALHOS_PRODUTOS;

const SubProdutosCRUD_ABA_FORNECEDORES = ABA_FORNECEDORES; // Para buscar nomes de fornecedores
const SubProdutosCRUD_CABECALHOS_FORNECEDORES = CABECALHOS_FORNECEDORES;


// Índices das colunas SubProdutos (baseado em Constantes.gs)
const SubProdutosCRUD_IDX_SUBPRODUTO_DATA_CADASTRO = CABECALHOS_SUBPRODUTOS.indexOf("Data de Cadastro");
const SubProdutosCRUD_IDX_SUBPRODUTO_ID = CABECALHOS_SUBPRODUTOS.indexOf("ID");
const SubProdutosCRUD_IDX_SUBPRODUTO_NOME = CABECALHOS_SUBPRODUTOS.indexOf("SubProduto");
const SubProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO = CABECALHOS_SUBPRODUTOS.indexOf("Produto Vinculado"); // Armazena NOME do produto
const SubProdutosCRUD_IDX_SUBPRODUTO_FORNECEDOR = CABECALHOS_SUBPRODUTOS.indexOf("Fornecedor"); // Armazena NOME do fornecedor
// Adicione outros índices de SubProdutosCRUD_CABECALHOS_SUBPRODUTOS conforme necessário

// Índices das colunas Produtos (para buscar nomes a partir de IDs)
const SubProdutosCRUD_IDX_PRODUTO_ID_REF = CABECALHOS_PRODUTOS.indexOf("ID");
const SubProdutosCRUD_IDX_PRODUTO_NOME_REF = CABECALHOS_PRODUTOS.indexOf("Produto");

// Índices das colunas Fornecedores (para buscar nomes a partir de IDs)
const SubProdutosCRUD_IDX_FORNECEDOR_ID_REF = CABECALHOS_FORNECEDORES.indexOf("ID");
const SubProdutosCRUD_IDX_FORNECEDOR_NOME_REF = CABECALHOS_FORNECEDORES.indexOf("Fornecedor");


/**
 * Normaliza o texto para comparação: remove acentos, converte para minúsculas e remove espaços extras.
 * @param {string} texto O texto a ser normalizado.
 * @return {string} O texto normalizado.
 */
function SubProdutosCRUD_normalizarTextoComparacao(texto) {
  if (!texto || typeof texto !== 'string') return "";
  return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
}

/**
 * Obtém o nome de um produto pelo seu ID.
 * @param {string} produtoId O ID do produto.
 * @param {SpreadsheetApp.Spreadsheet} ss A planilha ativa.
 * @return {string|null} O nome do produto ou null se não encontrado.
 */
function SubProdutosCRUD_obterNomeProdutoPorId(produtoId, ss) {
  if (!produtoId) return null;
  const abaProdutos = ss.getSheetByName(SubProdutosCRUD_ABA_PRODUTOS);
  if (!abaProdutos) {
    console.warn(`Aba Produtos '${SubProdutosCRUD_ABA_PRODUTOS}' não encontrada ao tentar obter nome por ID.`);
    return null;
  }
  const range = abaProdutos.getDataRange().getValues();
  if (SubProdutosCRUD_IDX_PRODUTO_ID_REF === -1 || SubProdutosCRUD_IDX_PRODUTO_NOME_REF === -1) {
      console.error("Índices de ID ou Nome de Produto não encontrados em Constantes.gs para Produtos.");
      return null;
  }
  for (let i = 1; i < range.length; i++) {
    if (String(range[i][SubProdutosCRUD_IDX_PRODUTO_ID_REF]) === String(produtoId)) {
      return String(range[i][SubProdutosCRUD_IDX_PRODUTO_NOME_REF]);
    }
  }
  console.warn(`Produto com ID '${produtoId}' não encontrado na aba Produtos.`);
  return null;
}

/**
 * Obtém o nome de um fornecedor pelo seu ID.
 * @param {string} fornecedorId O ID do fornecedor.
 * @param {SpreadsheetApp.Spreadsheet} ss A planilha ativa.
 * @return {string|null} O nome do fornecedor ou null se não encontrado.
 */
function SubProdutosCRUD_obterNomeFornecedorPorId(fornecedorId, ss) {
  if (!fornecedorId) return null; // Fornecedor é opcional
  const abaFornecedores = ss.getSheetByName(SubProdutosCRUD_ABA_FORNECEDORES);
  if (!abaFornecedores) {
    console.warn(`Aba Fornecedores '${SubProdutosCRUD_ABA_FORNECEDORES}' não encontrada ao tentar obter nome por ID.`);
    return null;
  }
  const range = abaFornecedores.getDataRange().getValues();
   if (SubProdutosCRUD_IDX_FORNECEDOR_ID_REF === -1 || SubProdutosCRUD_IDX_FORNECEDOR_NOME_REF === -1) {
      console.error("Índices de ID ou Nome de Fornecedor não encontrados em Constantes.gs para Fornecedores.");
      return null;
  }
  for (let i = 1; i < range.length; i++) {
    if (String(range[i][SubProdutosCRUD_IDX_FORNECEDOR_ID_REF]) === String(fornecedorId)) {
      return String(range[i][SubProdutosCRUD_IDX_FORNECEDOR_NOME_REF]);
    }
  }
  console.warn(`Fornecedor com ID '${fornecedorId}' não encontrado na aba Fornecedores.`);
  return null;
}


/**
 * Cria um novo subproduto na planilha.
 * @param {object} dadosNovoSubProduto Objeto contendo os dados do novo subproduto. Espera ID para Produto Vinculado e Fornecedor.
 * @return {object} { success: boolean, message: string, novoId?: string }
 */
function SubProdutosCRUD_criarNovoSubProduto(dadosNovoSubProduto) {
  try {
    console.log("SubProdutosCRUD_criarNovoSubProduto: Iniciando com dados:", JSON.stringify(dadosNovoSubProduto));

    const nomeDoCampoSubProduto = SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_NOME];
    const nomeDoCampoProdutoVinculadoForm = "Produto Vinculado"; // Nome do campo como vem do form (ID do produto)
    const nomeDoCampoFornecedorForm = "Fornecedor"; // Nome do campo como vem do form (ID do fornecedor)
    const nomeDoCampoUN = SubProdutosCRUD_CABECALHOS_SUBPRODUTOS.find(h => h.toUpperCase() === "UN");


    if (!dadosNovoSubProduto || !dadosNovoSubProduto[nomeDoCampoSubProduto]) {
      throw new Error(`O campo '${nomeDoCampoSubProduto}' é obrigatório.`);
    }
    if (!dadosNovoSubProduto[nomeDoCampoProdutoVinculadoForm]) {
      throw new Error(`O campo '${nomeDoCampoProdutoVinculadoForm}' (ID do Produto) é obrigatório.`);
    }
     if (!nomeDoCampoUN || !dadosNovoSubProduto[nomeDoCampoUN]) {
      throw new Error(`O campo '${nomeDoCampoUN}' é obrigatório.`);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSubProdutos = ss.getSheetByName(SubProdutosCRUD_ABA_SUBPRODUTOS);
    if (!abaSubProdutos) throw new Error(`Aba '${SubProdutosCRUD_ABA_SUBPRODUTOS}' não encontrada.`);

    // Obter NOME do Produto Vinculado a partir do ID recebido
    const idProdutoVinculadoRecebido = dadosNovoSubProduto[nomeDoCampoProdutoVinculadoForm];
    const nomeProdutoVinculadoParaSalvar = SubProdutosCRUD_obterNomeProdutoPorId(idProdutoVinculadoRecebido, ss);
    if (!nomeProdutoVinculadoParaSalvar) {
        throw new Error(`Produto Vinculado com ID '${idProdutoVinculadoRecebido}' não encontrado. Não é possível criar o subproduto.`);
    }

    // Obter NOME do Fornecedor a partir do ID recebido (se houver)
    let nomeFornecedorParaSalvar = "";
    const idFornecedorRecebido = dadosNovoSubProduto[nomeDoCampoFornecedorForm];
    if (idFornecedorRecebido) {
        nomeFornecedorParaSalvar = SubProdutosCRUD_obterNomeFornecedorPorId(idFornecedorRecebido, ss);
        if (!nomeFornecedorParaSalvar) {
            console.warn(`Fornecedor com ID '${idFornecedorRecebido}' não encontrado, mas o subproduto será criado sem fornecedor ou com ID.`);
            // Decida se quer lançar erro ou salvar com ID/vazio. Por ora, salva vazio se não encontrar nome.
            nomeFornecedorParaSalvar = ""; // Ou poderia ser idFornecedorRecebido se a planilha aceitar IDs.
        }
    }

    const nomeNovoSubProdutoNormalizado = SubProdutosCRUD_normalizarTextoComparacao(dadosNovoSubProduto[nomeDoCampoSubProduto]);
    const nomeProdutoVinculadoNormalizado = SubProdutosCRUD_normalizarTextoComparacao(nomeProdutoVinculadoParaSalvar);

    const todasAsLinhas = abaSubProdutos.getDataRange().getValues();
    if (SubProdutosCRUD_IDX_SUBPRODUTO_NOME === -1 || SubProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO === -1) {
      throw new Error(`Colunas '${nomeDoCampoSubProduto}' ou '${SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO]}' não encontradas nos cabeçalhos de SubProdutos.`);
    }

    for (let i = 1; i < todasAsLinhas.length; i++) {
      const nomeSubProdutoExistente = todasAsLinhas[i][SubProdutosCRUD_IDX_SUBPRODUTO_NOME];
      const nomeProdutoVinculadoExistente = todasAsLinhas[i][SubProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO];
      if (SubProdutosCRUD_normalizarTextoComparacao(String(nomeSubProdutoExistente)) === nomeNovoSubProdutoNormalizado &&
          SubProdutosCRUD_normalizarTextoComparacao(String(nomeProdutoVinculadoExistente)) === nomeProdutoVinculadoNormalizado) {
        throw new Error(`O subproduto '${dadosNovoSubProduto[nomeDoCampoSubProduto]}' já está cadastrado para o produto '${nomeProdutoVinculadoParaSalvar}'.`);
      }
    }

    let proximoId = 1;
    if (SubProdutosCRUD_IDX_SUBPRODUTO_ID === -1) {
      throw new Error("Coluna 'ID' de SubProdutos não encontrada. Não é possível gerar novo ID.");
    }
    if (todasAsLinhas.length > 1) {
      const idsExistentes = todasAsLinhas.slice(1)
        .map(linha => parseInt(linha[SubProdutosCRUD_IDX_SUBPRODUTO_ID]))
        .filter(id => !isNaN(id));
      if (idsExistentes.length > 0) {
        proximoId = Math.max(...idsExistentes) + 1;
      }
    }
    const novoIdGerado = String(proximoId);

    const novaLinhaArray = [];
    SubProdutosCRUD_CABECALHOS_SUBPRODUTOS.forEach(nomeCabecalho => {
      if (nomeCabecalho === SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_DATA_CADASTRO]) {
        novaLinhaArray.push(new Date());
      } else if (nomeCabecalho === SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_ID]) {
        novaLinhaArray.push(novoIdGerado);
      } else if (nomeCabecalho === SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO]) {
        novaLinhaArray.push(nomeProdutoVinculadoParaSalvar); // Salva o NOME do produto
      } else if (nomeCabecalho === SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_FORNECEDOR]) {
        novaLinhaArray.push(nomeFornecedorParaSalvar); // Salva o NOME do fornecedor
      } else {
        // Para outros campos, pega o valor do objeto de dados se existir
        novaLinhaArray.push(dadosNovoSubProduto[nomeCabecalho] !== undefined ? dadosNovoSubProduto[nomeCabecalho] : "");
      }
    });

    abaSubProdutos.appendRow(novaLinhaArray);
    SpreadsheetApp.flush();
    return { success: true, message: "Subproduto criado com sucesso!", novoId: novoIdGerado };
  } catch (e) {
    console.error("ERRO em SubProdutosCRUD_criarNovoSubProduto: " + e.toString() + " Stack: " + (e.stack || 'N/A'));
    return { success: false, message: e.message };
  }
}

/**
 * Atualiza um subproduto existente na planilha.
 * @param {object} dadosSubProdutoAtualizar Objeto contendo os dados do subproduto, incluindo o "ID". Espera ID para Produto Vinculado e Fornecedor.
 * @return {object} { success: boolean, message: string }
 */
function SubProdutosCRUD_atualizarSubProduto(dadosSubProdutoAtualizar) {
  try {
    console.log("SubProdutosCRUD_atualizarSubProduto: Iniciando com dados:", JSON.stringify(dadosSubProdutoAtualizar));
    const idParaAtualizar = dadosSubProdutoAtualizar["ID"];
    if (!idParaAtualizar) throw new Error("ID do subproduto é obrigatório para atualização.");

    const nomeDoCampoSubProduto = SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_NOME];
    const nomeSubProdutoAtualizado = dadosSubProdutoAtualizar[nomeDoCampoSubProduto];
    if (!nomeSubProdutoAtualizado) throw new Error(`O campo '${nomeDoCampoSubProduto}' é obrigatório.`);

    const nomeDoCampoProdutoVinculadoForm = "Produto Vinculado"; // Nome do campo como vem do form (ID do produto)
    const idProdutoVinculadoRecebido = dadosSubProdutoAtualizar[nomeDoCampoProdutoVinculadoForm];
     if (!idProdutoVinculadoRecebido) {
      throw new Error(`O campo '${nomeDoCampoProdutoVinculadoForm}' (ID do Produto) é obrigatório.`);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSubProdutos = ss.getSheetByName(SubProdutosCRUD_ABA_SUBPRODUTOS);
    if (!abaSubProdutos) throw new Error(`Aba '${SubProdutosCRUD_ABA_SUBPRODUTOS}' não encontrada.`);

    // Obter NOME do Produto Vinculado a partir do ID recebido
    const nomeProdutoVinculadoParaSalvar = SubProdutosCRUD_obterNomeProdutoPorId(idProdutoVinculadoRecebido, ss);
    if (!nomeProdutoVinculadoParaSalvar) {
        throw new Error(`Produto Vinculado com ID '${idProdutoVinculadoRecebido}' não encontrado. Não é possível atualizar o subproduto.`);
    }

    // Obter NOME do Fornecedor a partir do ID recebido (se houver)
    let nomeFornecedorParaSalvar = "";
    const nomeDoCampoFornecedorForm = "Fornecedor"; // Nome do campo como vem do form (ID do fornecedor)
    const idFornecedorRecebido = dadosSubProdutoAtualizar[nomeDoCampoFornecedorForm];
    if (idFornecedorRecebido) {
        nomeFornecedorParaSalvar = SubProdutosCRUD_obterNomeFornecedorPorId(idFornecedorRecebido, ss);
         if (!nomeFornecedorParaSalvar) {
            console.warn(`Fornecedor com ID '${idFornecedorRecebido}' não encontrado. Será salvo vazio ou com ID.`);
            nomeFornecedorParaSalvar = "";
        }
    }

    const nomeSubProdutoAtualizadoNormalizado = SubProdutosCRUD_normalizarTextoComparacao(nomeSubProdutoAtualizado);
    const nomeProdutoVinculadoNormalizado = SubProdutosCRUD_normalizarTextoComparacao(nomeProdutoVinculadoParaSalvar);

    const range = abaSubProdutos.getDataRange();
    const todasAsLinhas = range.getValues();

    if (SubProdutosCRUD_IDX_SUBPRODUTO_ID === -1 || SubProdutosCRUD_IDX_SUBPRODUTO_NOME === -1 || SubProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO === -1) {
      throw new Error("Colunas 'ID', 'SubProduto' ou 'Produto Vinculado' não encontradas na definição de cabeçalhos de SubProdutos.");
    }

    let linhaParaAtualizarIndexNaPlanilha = -1;
    for (let i = 1; i < todasAsLinhas.length; i++) {
      const idLinhaAtual = String(todasAsLinhas[i][SubProdutosCRUD_IDX_SUBPRODUTO_ID]);
      if (idLinhaAtual === String(idParaAtualizar)) {
        linhaParaAtualizarIndexNaPlanilha = i;
      } else {
        const nomeSubProdutoExistente = todasAsLinhas[i][SubProdutosCRUD_IDX_SUBPRODUTO_NOME];
        const nomeProdutoVinculadoExistente = todasAsLinhas[i][SubProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO];
        if (SubProdutosCRUD_normalizarTextoComparacao(String(nomeSubProdutoExistente)) === nomeSubProdutoAtualizadoNormalizado &&
            SubProdutosCRUD_normalizarTextoComparacao(String(nomeProdutoVinculadoExistente)) === nomeProdutoVinculadoNormalizado) {
          throw new Error(`O subproduto '${nomeSubProdutoAtualizado}' já está cadastrado para o produto '${nomeProdutoVinculadoParaSalvar}' (em outro ID: ${idLinhaAtual}).`);
        }
      }
    }

    if (linhaParaAtualizarIndexNaPlanilha === -1) {
      throw new Error(`Subproduto com ID '${idParaAtualizar}' não encontrado para atualização.`);
    }

    const linhaOriginalValores = todasAsLinhas[linhaParaAtualizarIndexNaPlanilha];
    const linhaAtualizadaValores = [];
    let alteracoesReais = 0;

    SubProdutosCRUD_CABECALHOS_SUBPRODUTOS.forEach((nomeCabecalho, k_idx_constante) => {
      // Encontrar o índice real deste nomeCabecalho na planilha (pode ser diferente de k_idx_constante)
      const k_idx_planilha = todasAsLinhas[0].indexOf(nomeCabecalho);
      if (k_idx_planilha === -1) {
          console.warn(`Cabeçalho '${nomeCabecalho}' não encontrado na planilha SubProdutos. Será ignorado na atualização.`);
          return; // Pula este cabeçalho se não existir na planilha
      }

      if (nomeCabecalho === SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_DATA_CADASTRO] ||
          nomeCabecalho === SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_ID]) {
        linhaAtualizadaValores[k_idx_planilha] = linhaOriginalValores[k_idx_planilha];
      } else if (nomeCabecalho === SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO]) {
        const valorAntigo = linhaOriginalValores[k_idx_planilha];
        linhaAtualizadaValores[k_idx_planilha] = nomeProdutoVinculadoParaSalvar; // Salva o NOME do produto
        if (String(valorAntigo) !== String(nomeProdutoVinculadoParaSalvar)) alteracoesReais++;
      } else if (nomeCabecalho === SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_FORNECEDOR]) {
        const valorAntigo = linhaOriginalValores[k_idx_planilha];
        linhaAtualizadaValores[k_idx_planilha] = nomeFornecedorParaSalvar; // Salva o NOME do fornecedor
        if (String(valorAntigo) !== String(nomeFornecedorParaSalvar)) alteracoesReais++;
      } else {
        const valorNovo = dadosSubProdutoAtualizar[nomeCabecalho];
        const valorAntigo = linhaOriginalValores[k_idx_planilha];
        const valorParaSalvar = valorNovo !== undefined ? valorNovo : valorAntigo; // Mantém o antigo se não veio novo
        linhaAtualizadaValores[k_idx_planilha] = valorParaSalvar;

        let comparavelAntigo = (valorAntigo === null || valorAntigo === undefined) ? "" : String(valorAntigo);
        let comparavelNovo = (valorParaSalvar === null || valorParaSalvar === undefined) ? "" : String(valorParaSalvar);
        if (comparavelAntigo !== comparavelNovo) {
          alteracoesReais++;
        }
      }
    });
    
    // Garante que a linhaAtualizadaValores tenha o mesmo número de colunas que a planilha
    const linhaFinalParaSalvar = [];
    for(let i = 0; i < todasAsLinhas[0].length; i++) {
        linhaFinalParaSalvar[i] = linhaAtualizadaValores[i] !== undefined ? linhaAtualizadaValores[i] : linhaOriginalValores[i];
    }


    if (alteracoesReais > 0) {
      abaSubProdutos.getRange(linhaParaAtualizarIndexNaPlanilha + 1, 1, 1, linhaFinalParaSalvar.length).setValues([linhaFinalParaSalvar]);
      SpreadsheetApp.flush();
      return { success: true, message: "Subproduto atualizado com sucesso!" };
    } else {
      return { success: true, message: "Nenhum dado foi modificado." };
    }
  } catch (e) {
    console.error("ERRO em SubProdutosCRUD_atualizarSubProduto: " + e.toString() + " Stack: " + (e.stack || 'N/A'));
    return { success: false, message: e.message };
  }
}

/**
 * Exclui um subproduto da planilha.
 * @param {string} subProdutoId O ID do subproduto a ser excluído.
 * @return {object} { success: boolean, message: string }
 */
function SubProdutosCRUD_excluirSubProduto(subProdutoId) {
  try {
    console.log("SubProdutosCRUD_excluirSubProduto: Iniciando exclusão do ID:", subProdutoId);
    if (!subProdutoId) throw new Error("ID do subproduto é obrigatório para exclusão.");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSubProdutos = ss.getSheetByName(SubProdutosCRUD_ABA_SUBPRODUTOS);
    if (!abaSubProdutos) throw new Error(`Aba '${SubProdutosCRUD_ABA_SUBPRODUTOS}' não encontrada.`);

    const range = abaSubProdutos.getDataRange();
    const todasAsLinhas = range.getValues();

    if (SubProdutosCRUD_IDX_SUBPRODUTO_ID === -1) {
      throw new Error("Coluna 'ID' de SubProdutos não encontrada. Não é possível excluir.");
    }

    let linhaExcluida = false;
    // Iterar de baixo para cima para evitar problemas com a alteração dos índices das linhas ao deletar
    for (let i = todasAsLinhas.length - 1; i >= 1; i--) { // i >= 1 para pular o cabeçalho
      if (String(todasAsLinhas[i][SubProdutosCRUD_IDX_SUBPRODUTO_ID]) === String(subProdutoId)) {
        abaSubProdutos.deleteRow(i + 1); // i + 1 porque os índices da planilha são base 1
        linhaExcluida = true;
        break; // Encontrou e excluiu, pode sair do loop
      }
    }

    if (linhaExcluida) {
      SpreadsheetApp.flush();
      return { success: true, message: "Subproduto excluído com sucesso!" };
    } else {
      return { success: false, message: `Subproduto com ID '${subProdutoId}' não encontrado.` };
    }
  } catch (e) {
    console.error("ERRO em SubProdutosCRUD_excluirSubProduto: " + e.toString() + " Stack: " + (e.stack || 'N/A'));
    return { success: false, message: e.message };
  }
}

/**
 * Cadastra múltiplos subprodutos de uma vez.
 * O Fornecedor é global para o lote. O Produto Vinculado é individual por subproduto.
 * @param {object} dadosLote Objeto contendo { fornecedorGlobal?: string (ID), subProdutos: Array<Object> }.
 * Cada objeto em subProdutos deve ter 'ProdutoVinculadoID' e outros campos do subproduto.
 * @return {object} { success: boolean, message: string, detalhes?: Array<{nome: string, status: string, erro?: string}> }
 */
function SubProdutosCRUD_cadastrarMultiplosSubProdutos(dadosLote) {
  try {
    console.log("SubProdutosCRUD_cadastrarMultiplosSubProdutos (ajustado): Iniciando com dados:", JSON.stringify(dadosLote));
    // Ajuste: 'produtoVinculadoGlobal' não é mais esperado no nível raiz de dadosLote.
    // Cada subProduto em dadosLote.subProdutos deve ter sua própria propriedade para o ID do produto vinculado.
    if (!dadosLote || !dadosLote.subProdutos || dadosLote.subProdutos.length === 0) {
      throw new Error("Dados insuficientes para cadastro em lote. A lista 'subProdutos' é obrigatória e deve conter itens.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSubProdutos = ss.getSheetByName(SubProdutosCRUD_ABA_SUBPRODUTOS);
    if (!abaSubProdutos) throw new Error(`Aba '${SubProdutosCRUD_ABA_SUBPRODUTOS}' não encontrada.`);

    // Fornecedor Global (opcional)
    let nomeFornecedorGlobalParaSalvar = "";
    if (dadosLote.fornecedorGlobal) {
        nomeFornecedorGlobalParaSalvar = SubProdutosCRUD_obterNomeFornecedorPorId(dadosLote.fornecedorGlobal, ss);
        if (!nomeFornecedorGlobalParaSalvar) {
            console.warn(`Fornecedor Global com ID '${dadosLote.fornecedorGlobal}' não encontrado. Subprodutos serão cadastrados sem este fornecedor ou com o ID se a planilha permitir.`);
            nomeFornecedorGlobalParaSalvar = ""; // Ou dadosLote.fornecedorGlobal
        }
    }

    const todasAsLinhasExistentes = abaSubProdutos.getDataRange().getValues();
    let proximoId = 1;
    if (SubProdutosCRUD_IDX_SUBPRODUTO_ID === -1) {
      throw new Error("Coluna 'ID' de SubProdutos não encontrada.");
    }
    if (todasAsLinhasExistentes.length > 1) {
      const idsExistentes = todasAsLinhasExistentes.slice(1)
        .map(linha => parseInt(linha[SubProdutosCRUD_IDX_SUBPRODUTO_ID]))
        .filter(id => !isNaN(id));
      if (idsExistentes.length > 0) {
        proximoId = Math.max(...idsExistentes) + 1;
      }
    }

    const resultadosDetalhados = [];
    let subProdutosAdicionadosComSucesso = 0;
    const novasLinhasParaAdicionar = [];

    const nomeDoCampoSubProdutoConst = SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_NOME];
    const nomeDoCampoUNConst = SubProdutosCRUD_CABECALHOS_SUBPRODUTOS.find(h => h.toUpperCase() === "UN");
    const nomeDoCampoProdutoVinculadoIndividual = "ProdutoVinculadoID"; // Espera-se que cada subProduto no array tenha esta chave com o ID do produto.

    for (const subProdutoIndividual of dadosLote.subProdutos) {
      const nomeSubProdutoAtual = subProdutoIndividual[nomeDoCampoSubProdutoConst];
      const unSubProdutoAtual = subProdutoIndividual[nomeDoCampoUNConst];
      const idProdutoVinculadoIndividual = subProdutoIndividual[nomeDoCampoProdutoVinculadoIndividual];

      if (!nomeSubProdutoAtual || !unSubProdutoAtual || !idProdutoVinculadoIndividual) {
        resultadosDetalhados.push({
          nome: nomeSubProdutoAtual || "Nome não fornecido",
          status: "Falha",
          erro: `Campos '${nomeDoCampoSubProdutoConst}', '${nomeDoCampoUNConst}' e '${nomeDoCampoProdutoVinculadoIndividual}' são obrigatórios para cada subproduto.`
        });
        continue;
      }

      const nomeProdutoVinculadoIndividualParaSalvar = SubProdutosCRUD_obterNomeProdutoPorId(idProdutoVinculadoIndividual, ss);
      if (!nomeProdutoVinculadoIndividualParaSalvar) {
        resultadosDetalhados.push({ nome: nomeSubProdutoAtual, status: "Falha", erro: `Produto Vinculado com ID '${idProdutoVinculadoIndividual}' não encontrado.` });
        continue;
      }

      const nomeSubProdutoNormalizado = SubProdutosCRUD_normalizarTextoComparacao(nomeSubProdutoAtual);
      const nomeProdutoVinculadoIndividualNormalizado = SubProdutosCRUD_normalizarTextoComparacao(nomeProdutoVinculadoIndividualParaSalvar);

      // Verificar duplicidade contra dados já existentes na planilha
      let duplicadoExistente = false;
      for (let i = 1; i < todasAsLinhasExistentes.length; i++) {
        const nomeSubProdutoExistentePlanilha = todasAsLinhasExistentes[i][SubProdutosCRUD_IDX_SUBPRODUTO_NOME];
        const nomeProdutoVinculadoExistentePlanilha = todasAsLinhasExistentes[i][SubProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO];
        if (SubProdutosCRUD_normalizarTextoComparacao(String(nomeSubProdutoExistentePlanilha)) === nomeSubProdutoNormalizado &&
            SubProdutosCRUD_normalizarTextoComparacao(String(nomeProdutoVinculadoExistentePlanilha)) === nomeProdutoVinculadoIndividualNormalizado) {
          duplicadoExistente = true;
          break;
        }
      }
      if (duplicadoExistente) {
        resultadosDetalhados.push({ nome: nomeSubProdutoAtual, status: "Falha", erro: `Já cadastrado para o produto vinculado '${nomeProdutoVinculadoIndividualParaSalvar}'.` });
        continue;
      }

      // Verificar duplicidade contra itens já processados neste lote (para o mesmo produto vinculado)
      let duplicadoNoLote = novasLinhasParaAdicionar.some(linhaAdicionada =>
          SubProdutosCRUD_normalizarTextoComparacao(String(linhaAdicionada[SubProdutosCRUD_IDX_SUBPRODUTO_NOME])) === nomeSubProdutoNormalizado &&
          SubProdutosCRUD_normalizarTextoComparacao(String(linhaAdicionada[SubProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO])) === nomeProdutoVinculadoIndividualNormalizado
      );
       if (duplicadoNoLote) {
        resultadosDetalhados.push({ nome: nomeSubProdutoAtual, status: "Falha", erro: `Duplicado neste lote para o produto vinculado '${nomeProdutoVinculadoIndividualParaSalvar}'.` });
        continue;
      }

      const novaLinhaArray = [];
      const idAtualGerado = String(proximoId++);
      SubProdutosCRUD_CABECALHOS_SUBPRODUTOS.forEach(nomeCabecalho => {
        if (nomeCabecalho === SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_DATA_CADASTRO]) {
          novaLinhaArray.push(new Date());
        } else if (nomeCabecalho === SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_ID]) {
          novaLinhaArray.push(idAtualGerado);
        } else if (nomeCabecalho === SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_PRODUTO_VINCULADO]) {
          novaLinhaArray.push(nomeProdutoVinculadoIndividualParaSalvar); // Salva o nome do produto vinculado individual
        } else if (nomeCabecalho === SubProdutosCRUD_CABECALHOS_SUBPRODUTOS[SubProdutosCRUD_IDX_SUBPRODUTO_FORNECEDOR]) {
          novaLinhaArray.push(nomeFornecedorGlobalParaSalvar); // Salva o nome do fornecedor global
        } else {
          novaLinhaArray.push(subProdutoIndividual[nomeCabecalho] !== undefined ? subProdutoIndividual[nomeCabecalho] : "");
        }
      });
      novasLinhasParaAdicionar.push(novaLinhaArray);
      resultadosDetalhados.push({ nome: nomeSubProdutoAtual, status: "Sucesso", id: idAtualGerado, produtoVinculado: nomeProdutoVinculadoIndividualParaSalvar });
      subProdutosAdicionadosComSucesso++;
    }

    if (novasLinhasParaAdicionar.length > 0) {
      abaSubProdutos.getRange(abaSubProdutos.getLastRow() + 1, 1, novasLinhasParaAdicionar.length, novasLinhasParaAdicionar[0].length).setValues(novasLinhasParaAdicionar);
      SpreadsheetApp.flush();
    }

    let mensagemFinal = `${subProdutosAdicionadosComSucesso} de ${dadosLote.subProdutos.length} subprodutos foram processados.`;
    if (subProdutosAdicionadosComSucesso === dadosLote.subProdutos.length && dadosLote.subProdutos.length > 0) {
        mensagemFinal = "Todos os subprodutos foram cadastrados com sucesso!";
    } else if (subProdutosAdicionadosComSucesso === 0 && dadosLote.subProdutos.length > 0) {
        mensagemFinal = "Nenhum subproduto pôde ser cadastrado. Verifique os erros e se os produtos vinculados existem.";
    } else if (dadosLote.subProdutos.length === 0) {
        // Isso não deveria acontecer devido à validação inicial, mas é uma segurança.
        mensagemFinal = "Nenhum subproduto foi fornecido para cadastro.";
    }


    return { success: true, message: mensagemFinal, detalhes: resultadosDetalhados };

  } catch (e) {
    console.error("ERRO em SubProdutosCRUD_cadastrarMultiplosSubProdutos (ajustado): " + e.toString() + " Stack: " + (e.stack || 'N/A'));
    return { success: false, message: e.message, detalhes: [] };
  }
}


/**
 * Obtém a lista completa de produtos para popular dropdowns.
 * Retorna objetos com ID e Nome do Produto.
 * @return {Array<Object>} Array de objetos de produtos, cada um com {ID, Produto}.
 */
function SubProdutosCRUD_obterTodosProdutosParaDropdown() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaProdutos = ss.getSheetByName(SubProdutosCRUD_ABA_PRODUTOS);
    if (!abaProdutos) {
      console.warn(`Aba de produtos '${SubProdutosCRUD_ABA_PRODUTOS}' não encontrada.`);
      return [];
    }
    const range = abaProdutos.getDataRange().getValues();
    if (SubProdutosCRUD_IDX_PRODUTO_ID_REF === -1 || SubProdutosCRUD_IDX_PRODUTO_NOME_REF === -1) {
      throw new Error("Colunas 'ID' ou 'Produto' não encontradas na aba Produtos. Verifique Constantes.gs e os índices (SubProdutosCRUD_IDX_PRODUTO_ID_REF, SubProdutosCRUD_IDX_PRODUTO_NOME_REF).");
    }
    const produtos = [];
    if (range.length > 1) { // Pula cabeçalho
      for (let i = 1; i < range.length; i++) {
        if (range[i][SubProdutosCRUD_IDX_PRODUTO_ID_REF] && range[i][SubProdutosCRUD_IDX_PRODUTO_NOME_REF]) { // Garante que ID e Nome existam
            produtos.push({
                ID: String(range[i][SubProdutosCRUD_IDX_PRODUTO_ID_REF]),
                Produto: String(range[i][SubProdutosCRUD_IDX_PRODUTO_NOME_REF])
            });
        }
      }
    }
    // Ordenar por nome do produto
    produtos.sort((a, b) => a.Produto.localeCompare(b.Produto));
    return produtos;
  } catch (e) {
    console.error("Erro em SubProdutosCRUD_obterTodosProdutosParaDropdown: " + e.toString());
    throw e; // Re-lança para ser pego pelo controller
  }
}

/**
 * Obtém a lista completa de fornecedores para popular dropdowns.
 * Retorna objetos com ID e Nome do Fornecedor.
 * @return {Array<Object>} Array de objetos de fornecedores, cada um com {ID, Fornecedor}.
 */
function SubProdutosCRUD_obterTodosFornecedoresParaDropdown() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaFornecedores = ss.getSheetByName(SubProdutosCRUD_ABA_FORNECEDORES);
    if (!abaFornecedores) {
      console.warn(`Aba de fornecedores '${SubProdutosCRUD_ABA_FORNECEDORES}' não encontrada.`);
      return [];
    }
    const range = abaFornecedores.getDataRange().getValues();
    if (SubProdutosCRUD_IDX_FORNECEDOR_ID_REF === -1 || SubProdutosCRUD_IDX_FORNECEDOR_NOME_REF === -1) {
      throw new Error("Colunas 'ID' ou 'Fornecedor' não encontradas na aba Fornecedores. Verifique Constantes.gs e os índices (SubProdutosCRUD_IDX_FORNECEDOR_ID_REF, SubProdutosCRUD_IDX_FORNECEDOR_NOME_REF).");
    }
    const fornecedores = [];
    if (range.length > 1) { // Pula cabeçalho
      for (let i = 1; i < range.length; i++) {
         if (range[i][SubProdutosCRUD_IDX_FORNECEDOR_ID_REF] && range[i][SubProdutosCRUD_IDX_FORNECEDOR_NOME_REF]) { // Garante que ID e Nome existam
            fornecedores.push({
                ID: String(range[i][SubProdutosCRUD_IDX_FORNECEDOR_ID_REF]),
                Fornecedor: String(range[i][SubProdutosCRUD_IDX_FORNECEDOR_NOME_REF])
            });
        }
      }
    }
    // Ordenar por nome do fornecedor
    fornecedores.sort((a, b) => a.Fornecedor.localeCompare(b.Fornecedor));
    return fornecedores;
  } catch (e) {
    console.error("Erro em SubProdutosCRUD_obterTodosFornecedoresParaDropdown: " + e.toString());
    throw e; // Re-lança para ser pego pelo controller
  }
}
