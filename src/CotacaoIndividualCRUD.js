// @ts-nocheck

//####################################################################################################
// MÓDULO: COTACAO INDIVIDUAL (SERVER-SIDE CRUD PRINCIPAL)
// Funções CRUD para os detalhes de uma cotação individual e operações relacionadas,
// excluindo lógicas de módulos específicos como Etapas e Funções do Portal.
//####################################################################################################

/**
 * CotacaoIndividualCRUD_parseNumeroPtBr
 * Normaliza números vindos do Sheets (pt-BR e en-US).
 * Aceita: number, "1.234,56", "1234.56", "2,5", "10,0000", etc.
 * Retorna NaN para valores vazios/Date/texto não numérico.
 */
function CotacaoIndividualCRUD_parseNumeroPtBr(valor) {
  if (valor === null || valor === undefined) return NaN;
  if (typeof valor === 'number') return Number(valor);
  if (valor instanceof Date) return NaN;

  const s = String(valor).trim();
  if (!s) return NaN;

  // remove espaços; remove separador de milhar "."; troca "," por "."
  const normalizado = s
    .replace(/\s+/g, '')
    .replace(/\.(?=\d{3}(?:\D|$))/g, '') // remove "." de milhar (ex.: 1.234,56)
    .replace(',', '.');

  const n = Number(normalizado);
  return Number.isFinite(n) ? n : NaN;
}

/**
 * Cria um mapa de Produto -> Média das 3 últimas compras a partir da aba "Cotacoes".
 * Itera a planilha de baixo para cima para otimizar a busca pelos itens mais recentes.
 * @returns {object} Um mapa onde a chave é o nome do produto e o valor é a demanda média.
 */
function CotacaoIndividualCRUD_criarMapaDemandaMediaProdutos() {
  console.log("CotacaoIndividualCRUD_criarMapaDemandaMediaProdutos: Iniciando criação do mapa de demanda média.");
  const mapaDemandas = {};
  const valoresComprasPorProduto = {};

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaCotacoes = planilha.getSheetByName(ABA_COTACOES); // Constante global

    if (!abaCotacoes) {
      console.error(`CotacaoIndividualCRUD_criarMapaDemandaMediaProdutos: Aba "${ABA_COTACOES}" não encontrada.`);
      return mapaDemandas;
    }

    const ultimaLinha = abaCotacoes.getLastRow();
    if (ultimaLinha <= 1) {
      console.log(`CotacaoIndividualCRUD_criarMapaDemandaMediaProdutos: Aba "${ABA_COTACOES}" vazia.`);
      return mapaDemandas;
    }

    const todosOsValores = abaCotacoes.getRange(1, 1, ultimaLinha, abaCotacoes.getLastColumn()).getValues();
    const cabecalhos = todosOsValores[0];
    
    const indiceProduto = cabecalhos.indexOf("Produto");
    const indiceComprar = cabecalhos.indexOf("Comprar");

    if (indiceProduto === -1 || indiceComprar === -1) {
      console.error(`CotacaoIndividualCRUD_criarMapaDemandaMediaProdutos: Colunas "Produto" ou "Comprar" não encontradas na aba "${ABA_COTACOES}".`);
      return mapaDemandas;
    }

    // Itera de baixo para cima para pegar as compras mais recentes primeiro
    for (let i = ultimaLinha - 1; i >= 1; i--) {
      const linha = todosOsValores[i];
      const nomeProduto = String(linha[indiceProduto]).trim();

      if (nomeProduto) {
        // Se ainda não temos 3 valores para este produto, buscamos mais
        if (!valoresComprasPorProduto[nomeProduto] || valoresComprasPorProduto[nomeProduto].length < 3) {
          const valorComprar = linha[indiceComprar];
          const quantidade = CotacaoIndividualCRUD_parseNumeroPtBr(valorComprar);

          // Considera apenas valores de compra positivos
          if (Number.isFinite(quantidade) && quantidade > 0) {
            if (!valoresComprasPorProduto[nomeProduto]) {
              valoresComprasPorProduto[nomeProduto] = [];
            }
            valoresComprasPorProduto[nomeProduto].push(quantidade);
          }
        }
      }
    }

    // Calcula a média para cada produto que teve valores encontrados
    for (const produto in valoresComprasPorProduto) {
      const compras = valoresComprasPorProduto[produto];
      if (compras.length > 0) {
        const soma = compras.reduce((acc, val) => acc + val, 0);
        mapaDemandas[produto] = soma / compras.length;
      }
    }

    console.log(`CotacaoIndividualCRUD_criarMapaDemandaMediaProdutos: Mapa de demanda média criado com ${Object.keys(mapaDemandas).length} entradas.`);
  } catch (error) {
    console.error("CotacaoIndividualCRUD_criarMapaDemandaMediaProdutos: Erro ao criar mapa de demanda: " + error.toString() + " Stack: " + error.stack);
  }
  return mapaDemandas;
}

/**
 * Cria um mapa de Produto -> Estoque Mínimo a partir da aba de Produtos.
 * @return {object} Um mapa onde a chave é o nome do produto e o valor é o estoque mínimo.
 */
function CotacaoIndividualCRUD_criarMapaEstoqueMinimoProdutos() {
  // ... (código original desta função permanece aqui, pois é genérico)
  console.log("CotacaoIndividualCRUD_criarMapaEstoqueMinimoProdutos: Iniciando criação do mapa de estoque mínimo.");
  const mapaEstoque = {};
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaProdutos = planilha.getSheetByName(ABA_PRODUTOS); // Constante global

    if (!abaProdutos) {
      console.error(`CotacaoIndividualCRUD_criarMapaEstoqueMinimoProdutos: Aba "${ABA_PRODUTOS}" não encontrada.`);
      return mapaEstoque; 
    }

    const ultimaLinha = abaProdutos.getLastRow();
    if (ultimaLinha <= 1) { 
      console.log(`CotacaoIndividualCRUD_criarMapaEstoqueMinimoProdutos: Aba "${ABA_PRODUTOS}" vazia ou só cabeçalho.`);
      return mapaEstoque;
    }

    const rangeCompleto = abaProdutos.getRange(1, 1, ultimaLinha, abaProdutos.getLastColumn());
    const todosOsValores = rangeCompleto.getValues();
    const cabecalhosPlanilhaProdutos = todosOsValores[0]; 

    const indiceProduto = cabecalhosPlanilhaProdutos.indexOf("Produto");
    const indiceEstoqueMinimo = cabecalhosPlanilhaProdutos.indexOf("Estoque Minimo");

    if (indiceProduto === -1) {
      console.error(`CotacaoIndividualCRUD_criarMapaEstoqueMinimoProdutos: Coluna "Produto" não encontrada na aba "${ABA_PRODUTOS}".`);
      return mapaEstoque; 
    }
    if (indiceEstoqueMinimo === -1) {
        console.warn(`CotacaoIndividualCRUD_criarMapaEstoqueMinimoProdutos: Coluna "Estoque Minimo" não encontrada na aba "${ABA_PRODUTOS}". Estoques mínimos não serão mapeados.`);
    }
    
    for (let i = 1; i < todosOsValores.length; i++) {
      const linha = todosOsValores[i];
      const nomeProduto = String(linha[indiceProduto]).trim();
      let estoqueMinimo = null;

      if (indiceEstoqueMinimo !== -1) { 
        const valorEstoque = linha[indiceEstoqueMinimo];
        if (valorEstoque !== "" && valorEstoque !== null && valorEstoque !== undefined) {
          const num = parseFloat(String(valorEstoque).replace(",", ".")); 
          estoqueMinimo = isNaN(num) ? String(valorEstoque).trim() : num; 
        }
      }
      
      if (nomeProduto) { 
        mapaEstoque[nomeProduto] = estoqueMinimo;
      }
    }
    console.log(`CotacaoIndividualCRUD_criarMapaEstoqueMinimoProdutos: Mapa de estoque mínimo criado com ${Object.keys(mapaEstoque).length} entradas.`);
  } catch (error) {
    console.error("CotacaoIndividualCRUD_criarMapaEstoqueMinimoProdutos: Erro ao criar mapa de estoque mínimo: " + error.toString() + " Stack: " + error.stack);
  }
  return mapaEstoque;
}

/**
 * Busca todos os produtos/linhas de uma cotação específica na aba 'Cotações'.
 * Adiciona o "EstoqueMinimoProdutoPrincipal" e a "DemandaMediaProdutoPrincipal".
 * @param {string} idCotacaoAlvo O ID da cotação a ser buscada.
 * @return {Array<object>|null} Um array de objetos, onde cada objeto representa uma linha da cotação, ou null em caso de erro.
 */
function CotacaoIndividualCRUD_buscarProdutosPorIdCotacao(idCotacaoAlvo) {
  console.log("CotacaoIndividualCRUD_buscarProdutosPorIdCotacao: Buscando produtos para ID '" + idCotacaoAlvo + "'.");

  // CHAMA AS DUAS FUNÇÕES DE MAPEAMENTO
  const mapaEstoqueMinimoProdutos = CotacaoIndividualCRUD_criarMapaEstoqueMinimoProdutos();
  const mapaDemandaMediaProdutos = CotacaoIndividualCRUD_criarMapaDemandaMediaProdutos();

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaCotacoes = planilha.getSheetByName(ABA_COTACOES);
    if (!abaCotacoes) {
      console.error('CotacaoIndividualCRUD: Aba "' + ABA_COTACOES + '" não encontrada.');
      return null;
    }

    const ultimaLinha = abaCotacoes.getLastRow();
    if (ultimaLinha <= 1) {
      console.log('CotacaoIndividualCRUD: Aba "' + ABA_COTACOES + '" vazia ou só cabeçalho.');
      return [];
    }

    const ultimaColuna = abaCotacoes.getLastColumn();
    const range = abaCotacoes.getRange(1, 1, ultimaLinha, ultimaColuna);

    const valores = range.getValues();
    const displays = range.getDisplayValues();

    const cabPlanilha = valores[0];
    const cabConst = CABECALHOS_COTACOES; 

    if (!cabConst || cabConst.length === 0) {
      console.error('CotacaoIndividualCRUD: CABECALHOS_COTACOES não definida ou vazia.');
      return null;
    }

    const idxIdCotacao = cabPlanilha.indexOf('ID da Cotação');
    if (idxIdCotacao === -1) {
      console.error('CotacaoIndividualCRUD: Coluna "ID da Cotação" não encontrada.');
      return null;
    }

    const idxProdutoPrincipal = cabPlanilha.indexOf('Produto');

    const camposNumericosEsperados = [
      'Fator', 'Estoque Mínimo', 'Estoque Atual', 'Preço', 'Preço por Fator',
      'Comprar', 'Valor Total', 'Economia em Cotação'
    ];

    const itens = [];

    for (let i = 1; i < valores.length; i++) {
      const rowRaw = valores[i];
      const rowDisp = displays[i];

      if (String(rowRaw[idxIdCotacao]).trim() !== String(idCotacaoAlvo).trim()) continue;

      const item = {};

      let nomeProdutoPrincipal = null;
      if (idxProdutoPrincipal !== -1) {
        nomeProdutoPrincipal = String(rowRaw[idxProdutoPrincipal] || '').trim();
      }

      // INCLUSÃO DOS NOVOS DADOS NO ITEM
      if (nomeProdutoPrincipal) {
        item['EstoqueMinimoProdutoPrincipal'] = mapaEstoqueMinimoProdutos.hasOwnProperty(nomeProdutoPrincipal) 
            ? mapaEstoqueMinimoProdutos[nomeProdutoPrincipal] : null;

        item['DemandaMediaProdutoPrincipal'] = mapaDemandaMediaProdutos.hasOwnProperty(nomeProdutoPrincipal) 
            ? mapaDemandaMediaProdutos[nomeProdutoPrincipal] : null;
      } else {
        item['EstoqueMinimoProdutoPrincipal'] = null;
        item['DemandaMediaProdutoPrincipal'] = null;
      }

      // mapeamento campo a campo (lógica original mantida)
      cabConst.forEach(function (nomeCol) {
        const idx = cabPlanilha.indexOf(nomeCol);
        if (idx === -1 || idx >= rowRaw.length) {
          item[nomeCol] = null;
          return;
        }

        const valorRaw = rowRaw[idx];
        const valorDisp = rowDisp[idx];

        if (nomeCol === 'Data Abertura') {
          if (valorRaw instanceof Date) {
            item[nomeCol] = valorRaw.toISOString();
          } else if (valorRaw) {
            const d = new Date(valorRaw);
            item[nomeCol] = Number.isFinite(d.getTime()) ? d.toISOString() : String(valorDisp || valorRaw).trim();
          } else {
            item[nomeCol] = null;
          }
          return;
        }

        if (camposNumericosEsperados.indexOf(nomeCol) !== -1) {
          let n = CotacaoIndividualCRUD_parseNumeroPtBr(valorDisp);
          if (!Number.isFinite(n)) n = CotacaoIndividualCRUD_parseNumeroPtBr(valorRaw);
          item[nomeCol] = Number.isFinite(n) ? n : null;
          return;
        }

        if (valorRaw instanceof Date) {
          item[nomeCol] = valorRaw.toISOString();
        } else {
          item[nomeCol] = (valorRaw !== null && valorRaw !== undefined) ? String(valorRaw).trim() : null;
        }
      });

      if (!item._subProdutoOriginalPersistido) {
        item._subProdutoOriginalPersistido = item.SubProduto || null;
      }

      itens.push(item);
    }

    console.log('CotacaoIndividualCRUD: ' + itens.length + " produtos encontrados para ID '" + idCotacaoAlvo + "'.");
    return itens;
  } catch (e) {
    console.error('ERRO em CotacaoIndividualCRUD_buscarProdutosPorIdCotacao para ID "' + idCotacaoAlvo + '": ' + e.toString() + ' Stack: ' + e.stack);
    return null;
  }
}



/**
 * Salva a alteração de uma célula individual feita na página de Cotação Individual.
 * Se a coluna for "Preço", "Comprar" ou "Fator", recalcula e atualiza os campos dependentes.
 * @param {string} idCotacao O ID da cotação.
 * @param {object} identificadoresLinha Contém Produto, SubProdutoChave (nome original do subproduto), Fornecedor.
 * @param {string} colunaAlterada O nome da coluna que foi alterada.
 * @param {string|number|null} novoValor O novo valor para a célula.
 * @return {object} Um objeto { success: boolean, message: string, ..., valoresCalculados?: { valorTotal: number, precoPorFator: number } }.
 */
function CotacaoIndividualCRUD_salvarEdicaoCelulaCotacao(idCotacao, identificadoresLinha, colunaAlterada, novoValor) {
  Logger.log(
    "CotacaoIndividualCRUD_salvarEdicaoCelulaCotacao: ID Cotação '" + idCotacao +
    "', Identificadores: " + JSON.stringify(identificadoresLinha) +
    ", Coluna '" + colunaAlterada + "', Novo Valor '" + novoValor + "'"
  );

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaCotacoes = planilha.getSheetByName(ABA_COTACOES);
  const abaSubProdutos = planilha.getSheetByName(ABA_SUBPRODUTOS);

  const COLUNAS_SINCRONIZAVEIS_COM_SUBPRODUTOS =
    (typeof COLUNAS_PARA_ABA_SUBPRODUTOS !== 'undefined')
      ? COLUNAS_PARA_ABA_SUBPRODUTOS
      : ['SubProduto', 'Tamanho', 'UN', 'Fator'];

  const colunasTriggerCalculo = ['Preço', 'Comprar', 'Fator'];

  let updatedInCotacoes = false;
  let updatedInSubProdutos = false;
  let nomeProdutoPrincipalDaLinhaCotacao = identificadoresLinha.Produto;

  const resultado = {
    success: false,
    message: 'Nenhuma alteração realizada.',
    updatedInCotacoes: false,
    updatedInSubProdutos: false
  };

  if (!abaCotacoes) {
    resultado.message = 'Aba "' + ABA_COTACOES + '" não encontrada.';
    Logger.log('CotacaoIndividualCRUD_salvarEdicaoCelulaCotacao: ' + resultado.message);
    return resultado;
  }

  try {
    const dadosCot = abaCotacoes.getDataRange().getValues();
    const cabecalhosCot = dadosCot[0];
    const indicesCot = cabecalhosCot.reduce(function (acc, c, i) { acc[c] = i; return acc; }, {});

    const idxColunaAlteradaCot = indicesCot[colunaAlterada];
    const idxIdCotacaoCot      = indicesCot['ID da Cotação'];
    const idxProdutoCot        = indicesCot['Produto'];
    const idxSubProdutoCot     = indicesCot['SubProduto'];
    const idxFornecedorCot     = indicesCot['Fornecedor'];

    if (idxColunaAlteradaCot === undefined) throw new Error('Coluna "' + colunaAlterada + '" não encontrada na aba "' + ABA_COTACOES + '".');
    if ([idxIdCotacaoCot, idxProdutoCot, idxSubProdutoCot, idxFornecedorCot].some(function (x) { return x === undefined; })) {
      throw new Error('Colunas chave (ID da Cotação, Produto, SubProduto, Fornecedor) não encontradas na ABA_COTACOES.');
    }

    let linhaEncontradaCot = -1;
    for (let i = 1; i < dadosCot.length; i++) {
      const linhaAtual = dadosCot[i];
      if (
        String(linhaAtual[idxIdCotacaoCot]).trim() === String(idCotacao).trim() &&
        String(linhaAtual[idxProdutoCot]).trim() === String(identificadoresLinha.Produto).trim() &&
        String(linhaAtual[idxSubProdutoCot]).trim() === String(identificadoresLinha.SubProdutoChave).trim() &&
        String(linhaAtual[idxFornecedorCot]).trim() === String(identificadoresLinha.Fornecedor).trim()
      ) {
        abaCotacoes.getRange(i + 1, idxColunaAlteradaCot + 1).setValue(novoValor);
        updatedInCotacoes = true;
        linhaEncontradaCot = i + 1;
        Logger.log('CotacaoIndividualCRUD_salvarEdicaoCelulaCotacao: ABA_COTACOES - Linha ' + linhaEncontradaCot + ', Coluna "' + colunaAlterada + '" atualizada para: ' + novoValor);
        break;
      }
    }

    if (!updatedInCotacoes) {
      resultado.message = 'Linha não encontrada na ABA_COTACOES para os identificadores fornecidos.';
      Logger.log('CotacaoIndividualCRUD_salvarEdicaoCelulaCotacao: ' + resultado.message);
      return resultado;
    }

    // Recalcular usando DISPLAY VALUES (robusto contra formatação/Date)
    if (updatedInCotacoes && colunasTriggerCalculo.indexOf(colunaAlterada) !== -1) {
      const rangeLinha = abaCotacoes.getRange(linhaEncontradaCot, 1, 1, abaCotacoes.getLastColumn());
      const displays = rangeLinha.getDisplayValues()[0];

      const preco   = CotacaoIndividualCRUD_parseNumeroPtBr(displays[indicesCot['Preço']])   || 0;
      const comprar = CotacaoIndividualCRUD_parseNumeroPtBr(displays[indicesCot['Comprar']]) || 0;
      const fator   = CotacaoIndividualCRUD_parseNumeroPtBr(displays[indicesCot['Fator']])   || 0;

      const valorTotalCalculado    = (preco * comprar);
      const precoPorFatorCalculado = (fator !== 0) ? (preco / fator) : 0;

      if (indicesCot['Valor Total'] !== undefined) {
        abaCotacoes.getRange(linhaEncontradaCot, indicesCot['Valor Total'] + 1).setValue(valorTotalCalculado);
      }
      if (indicesCot['Preço por Fator'] !== undefined) {
        abaCotacoes.getRange(linhaEncontradaCot, indicesCot['Preço por Fator'] + 1).setValue(precoPorFatorCalculado);
      }

      resultado.valoresCalculados = {
        valorTotal: valorTotalCalculado,
        precoPorFator: precoPorFatorCalculado
      };
      Logger.log(
        'CotacaoIndividualCRUD_salvarEdicaoCelulaCotacao: Recalculo -> Valor Total: ' +
        valorTotalCalculado + ', Preço por Fator: ' + precoPorFatorCalculado
      );
    }

    if (colunaAlterada === 'SubProduto') {
      resultado.novoSubProdutoNomeSeAlterado = novoValor;
    }

    if (COLUNAS_SINCRONIZAVEIS_COM_SUBPRODUTOS.indexOf(colunaAlterada) !== -1 && abaSubProdutos) {
      const dadosSub = abaSubProdutos.getDataRange().getValues();
      const cabSub = dadosSub[0];
      const idxSub = cabSub.reduce(function (acc, c, i) { acc[c] = i; return acc; }, {});
      const idxProdV   = idxSub['Produto Vinculado'];
      const idxSubProd = idxSub['SubProduto'];
      const idxForn    = idxSub['Fornecedor'];
      const idxColSub  = idxSub[colunaAlterada];

      if (idxProdV !== undefined && idxSubProd !== undefined && idxColSub !== undefined) {
        for (let i = 1; i < dadosSub.length; i++) {
          const linhaSub = dadosSub[i];
          const fornecedorPlanilha = (idxForn !== undefined) ? String(linhaSub[idxForn]).trim() : null;

          const match = (
            String(linhaSub[idxProdV]).trim() === String(nomeProdutoPrincipalDaLinhaCotacao).trim() &&
            String(linhaSub[idxSubProd]).trim() === String(identificadoresLinha.SubProdutoChave).trim() &&
            (fornecedorPlanilha === null || fornecedorPlanilha === String(identificadoresLinha.Fornecedor).trim())
          );

          if (match) {
            abaSubProdutos.getRange(i + 1, idxColSub + 1).setValue(novoValor);
            updatedInSubProdutos = true;
            Logger.log('CotacaoIndividualCRUD_salvarEdicaoCelulaCotacao: ABA_SUBPRODUTOS - Linha ' + (i + 1) + ', Coluna "' + colunaAlterada + '" atualizada.');
            break;
          }
        }
      }
    }

    if (updatedInCotacoes) {
      resultado.success = true;
      resultado.message = '"' + colunaAlterada + '" atualizado com sucesso.';
    }
    resultado.updatedInCotacoes = updatedInCotacoes;
    resultado.updatedInSubProdutos = updatedInSubProdutos;

    return resultado;
  } catch (e) {
    Logger.log('ERRO CRÍTICO em CotacaoIndividualCRUD_salvarEdicaoCelulaCotacao: ' + e.toString() + ' Stack: ' + e.stack);
    resultado.success = false;
    resultado.message = 'Erro ao salvar alteração da célula: ' + e.message;
    return resultado;
  }
}


/**
 * NOVA FUNÇÃO: Salva um conjunto de alterações do modal de detalhes (SubProduto, Tamanho, etc.).
 * Atualiza ABA_COTACOES e, se aplicável, ABA_SUBPRODUTOS para todas as colunas alteradas em uma única operação.
 * @param {string} idCotacao O ID da cotação.
 * @param {object} identificadoresLinha Contém Produto, SubProdutoChave (nome original), Fornecedor.
 * @param {object} alteracoes Um objeto com as colunas e novos valores. Ex: { SubProduto: "Novo Nome", Fator: 1.5 }.
 * @return {object} Um objeto de resultado com { success, message, novoSubProdutoNomeSeAlterado? }.
 */
function CotacaoIndividualCRUD_salvarEdicoesModalDetalhes(idCotacao, identificadoresLinha, alteracoes) {
  Logger.log(`CotacaoIndividualCRUD_salvarEdicoesModalDetalhes: ID Cotação '${idCotacao}', Alterações: ${JSON.stringify(alteracoes)}`);
  
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaCotacoes = planilha.getSheetByName(ABA_COTACOES);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Aguarda até 30 segundos pelo lock

  const resultado = { success: false, message: "Nenhuma alteração realizada." };

  try {
    if (!abaCotacoes) throw new Error(`Aba "${ABA_COTACOES}" não encontrada.`);

    const dadosCot = abaCotacoes.getDataRange().getValues();
    const cabecalhosCot = dadosCot[0];
    const indicesCot = cabecalhosCot.reduce((acc, c, i) => ({ ...acc, [c]: i }), {});

    const colunasChave = ["ID da Cotação", "Produto", "SubProduto", "Fornecedor"];
    if (colunasChave.some(c => indicesCot[c] === undefined)) {
      throw new Error(`Uma ou mais colunas chave não foram encontradas na aba "${ABA_COTACOES}".`);
    }

    let linhaEncontradaIndex = -1;
    for (let i = 1; i < dadosCot.length; i++) {
      const linha = dadosCot[i];
      if (String(linha[indicesCot["ID da Cotação"]]).trim() === String(idCotacao).trim() &&
          String(linha[indicesCot["Produto"]]).trim() === String(identificadoresLinha.Produto).trim() &&
          String(linha[indicesCot["SubProduto"]]).trim() === String(identificadoresLinha.SubProdutoChave).trim() &&
          String(linha[indicesCot["Fornecedor"]]).trim() === String(identificadoresLinha.Fornecedor).trim()) {
        linhaEncontradaIndex = i;
        break;
      }
    }

    if (linhaEncontradaIndex === -1) {
      throw new Error("Linha correspondente não encontrada na cotação para atualização.");
    }
    
    const linhaParaAtualizarRange = abaCotacoes.getRange(linhaEncontradaIndex + 1, 1, 1, cabecalhosCot.length);
    const valoresLinha = linhaParaAtualizarRange.getValues()[0];

    // Atualiza os valores na linha da aba "Cotações"
    for (const coluna in alteracoes) {
      if (indicesCot[coluna] !== undefined) {
        valoresLinha[indicesCot[coluna]] = alteracoes[coluna];
        Logger.log(`Preparando para atualizar [Cotações] Coluna: ${coluna} para o valor: ${alteracoes[coluna]}`);
      }
    }
    linhaParaAtualizarRange.setValues([valoresLinha]);
    resultado.updatedInCotacoes = true;
    
    // Se o SubProduto foi alterado, guarda o novo nome para retornar
    if (alteracoes.SubProduto) {
      resultado.novoSubProdutoNomeSeAlterado = alteracoes.SubProduto;
    }

    // Lógica para sincronizar com a aba "SubProdutos", se necessário
    const COLUNAS_SINCRONIZAVEIS = (typeof COLUNAS_PARA_ABA_SUBPRODUTOS !== 'undefined') ? COLUNAS_PARA_ABA_SUBPRODUTOS : ["SubProduto", "Tamanho", "UN", "Fator"];
    const alteracoesSincronizaveis = Object.keys(alteracoes).some(k => COLUNAS_SINCRONIZAVEIS.includes(k));

    if (alteracoesSincronizaveis) {
      const abaSubProdutos = planilha.getSheetByName(ABA_SUBPRODUTOS);
      if (abaSubProdutos) {
        // ... (A lógica interna e robusta de atualização da aba SubProdutos é executada aqui)
        // Por simplicidade e segurança, vamos reusar a lógica de busca e atualização da função original
        // Esta parte pode ser otimizada depois, mas por agora garante que nada quebre.
        for (const coluna in alteracoes) {
            if (COLUNAS_SINCRONIZAVEIS.includes(coluna)) {
                // Para garantir que a lógica complexa não quebre, chamamos a função antiga para cada alteração sincronizável.
                // Nota: O ideal seria ter uma função única que fizesse tudo, mas isso mantém a segurança pedida.
                CotacaoIndividualCRUD_salvarEdicaoCelulaCotacao(idCotacao, identificadoresLinha, coluna, alteracoes[coluna]);
            }
        }
        resultado.updatedInSubProdutos = true;
      }
    }
    
    resultado.success = true;
    resultado.message = "Detalhes do item atualizados com sucesso!";

  } catch (e) {
    Logger.log(`ERRO em CotacaoIndividualCRUD_salvarEdicoesModalDetalhes: ${e.toString()} ${e.stack}`);
    resultado.message = `Erro no servidor: ${e.message}`;
  } finally {
    lock.releaseLock();
  }
  return resultado;
}

/**
 * Acrescenta novos itens a uma cotação existente na aba "Cotacoes".
 * @param {string} idCotacaoExistente O ID da cotação para adicionar itens.
 * @param {object} opcoesCriacao { tipo: string, selecoes: Array<string> }.
 * @return {object} { success: boolean, idCotacao: string|null, numItens: number|null, message: string|null }.
 */
function CotacaoIndividualCRUD_acrescentarItensCotacao(idCotacaoExistente, opcoesCriacao) {
  Logger.log(`CotacaoIndividualCRUD_acrescentarItensCotacao: Adicionando ao ID '${idCotacaoExistente}' com opções: ${JSON.stringify(opcoesCriacao)}`);
  
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const dataAbertura = new Date(); // Usa a data atual para os novos itens

    // Reutiliza as funções de leitura de dados do CotacoesCRUD
    const todosSubProdutos = CotacoesCRUD_obterDadosCompletosDaAba(ABA_SUBPRODUTOS, CABECALHOS_SUBPRODUTOS);
    const todosProdutos = CotacoesCRUD_obterDadosCompletosDaAba(ABA_PRODUTOS, CABECALHOS_PRODUTOS);

    if (!todosSubProdutos || !todosProdutos) {
      return { success: false, message: "Falha ao carregar dados de Produtos ou SubProdutos para acrescentar itens." };
    }
    
    const produtosMap = todosProdutos.reduce((map, prod) => {
        map[prod["Produto"]] = prod;
        return map;
    }, {});

    let subProdutosFiltrados = [];
    const tipo = opcoesCriacao.tipo;
    const selecoesLowerCase = opcoesCriacao.selecoes.map(s => String(s).toLowerCase());

    // Lógica de filtragem (idêntica à de CotacoesCRUD_criarNovaCotacao)
    if (tipo === 'categoria') {
      const nomesProdutosDaCategoria = todosProdutos
        .filter(p => p["Categoria"] && selecoesLowerCase.includes(String(p["Categoria"]).toLowerCase()))
        .map(p => String(p["Produto"]).toLowerCase()); 
      subProdutosFiltrados = todosSubProdutos.filter(sp => {
        const produtoVinculado = sp["Produto Vinculado"] ? String(sp["Produto Vinculado"]).toLowerCase() : null;
        return produtoVinculado && nomesProdutosDaCategoria.includes(produtoVinculado);
      });
    } else if (tipo === 'fornecedor') {
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
      subProdutosFiltrados = todosSubProdutos.filter(sp => {
        const produtoVinculado = sp["Produto Vinculado"] ? String(sp["Produto Vinculado"]).toLowerCase() : null;
        return produtoVinculado && selecoesLowerCase.includes(produtoVinculado);
      });
    } else {
      return { success: false, message: "Tipo de criação desconhecido: " + tipo };
    }

    Logger.log(`CotacaoIndividualCRUD_acrescentarItensCotacao: ${subProdutosFiltrados.length} subprodutos filtrados para acrescentar.`);
    
    if (subProdutosFiltrados.length === 0) {
      return { success: true, idCotacao: idCotacaoExistente, numItens: 0, message: "Nenhum novo subproduto encontrado para os critérios selecionados." };
    }

    // Mapeamento dos novos itens para o formato da aba Cotações
    const linhasParaAdicionar = subProdutosFiltrados.map(subProd => {
      const produtoPrincipal = produtosMap[subProd["Produto Vinculado"]];
      const estoqueMinimo = produtoPrincipal ? produtoPrincipal["Estoque Minimo"] : "";
      const nomeProdutoPrincipalParaCotacao = subProd["Produto Vinculado"];

      let linha = []; 
      CABECALHOS_COTACOES.forEach(header => {
        switch(header) {
          case "ID da Cotação": linha.push(idCotacaoExistente); break; // USA O ID EXISTENTE
          case "Data Abertura": linha.push(dataAbertura); break;
          case "Produto": linha.push(nomeProdutoPrincipalParaCotacao); break;
          case "SubProduto": linha.push(subProd["SubProduto"]); break;
          case "Categoria": linha.push(produtoPrincipal ? produtoPrincipal["Categoria"] : subProd["Categoria"]); break;
          case "Fornecedor": linha.push(subProd["Fornecedor"]); break;
          case "Tamanho": linha.push(subProd["Tamanho"]); break;
          case "UN": linha.push(subProd["UN"]); break;
          case "Fator": linha.push(subProd["Fator"]); break;
          case "Estoque Mínimo": linha.push(estoqueMinimo); break;
          case "NCM": linha.push(subProd["NCM"]); break;
          case "CST": linha.push(subProd["CST"]); break;
          case "CFOP": linha.push(subProd["CFOP"]); break;
          case "Status da Cotação": linha.push(CotacoesCRUD_STATUS_NOVA_COTACAO); break; // Entra como 'Nova Cotação'
          default:
            linha.push(""); // Deixa campos calculáveis em branco
        }
      });
      return linha;
    });

    // Adiciona as novas linhas na aba de Cotações
    const abaCotacoes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_COTACOES);
    abaCotacoes.getRange(abaCotacoes.getLastRow() + 1, 1, linhasParaAdicionar.length, CABECALHOS_COTACOES.length)
               .setValues(linhasParaAdicionar);
    
    Logger.log(`CotacaoIndividualCRUD_acrescentarItensCotacao: ${linhasParaAdicionar.length} itens adicionados à cotação ${idCotacaoExistente}.`);
    return { 
      success: true, 
      idCotacao: idCotacaoExistente, 
      numItens: linhasParaAdicionar.length,
      message: "Itens acrescentados com sucesso."
    };

  } catch (e) {
    Logger.log(`ERRO CRÍTICO em CotacaoIndividualCRUD_acrescentarItensCotacao: ${e.toString()} Stack: ${e.stack}`);
    return { success: false, message: "Erro no servidor ao acrescentar itens à cotação: " + e.message };
  } finally {
      lock.releaseLock();
  }
}
