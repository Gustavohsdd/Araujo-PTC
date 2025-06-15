// @ts-nocheck
/**
 * @file RelatoriosCRUD.gs
 * @description Funções de acesso a dados para gerar relatórios.
 */

/**
 * Gera os dados para o Relatório de Análise de Compra.
 * @param {string} idCotacaoAlvo O ID da cotação cujos produtos serão analisados.
 * @return {Array<object>|null} Um array de objetos, cada um representando a análise de um produto, ou null em caso de erro.
 */
function RelatoriosCRUD_gerarDadosAnaliseCompra(idCotacaoAlvo) {
  try {
    const abaCotacoes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_COTACOES);
    if (!abaCotacoes) {
      throw new Error(`Aba "${ABA_COTACOES}" não encontrada.`);
    }

    const todosOsDados = abaCotacoes.getDataRange().getValues();
    const cabecalhos = todosOsDados.shift();

    const idxSubProduto = cabecalhos.indexOf("SubProduto");
    const idxIdCotacao = cabecalhos.indexOf("ID da Cotação");
    const idxDataAbertura = cabecalhos.indexOf("Data Abertura");
    const idxPreco = cabecalhos.indexOf("Preço");
    const idxComprar = cabecalhos.indexOf("Comprar");
    const idxPrecoPorFator = cabecalhos.indexOf("Preço por Fator");
    const idxUN = cabecalhos.indexOf("UN");

    if ([idxSubProduto, idxIdCotacao, idxDataAbertura, idxPreco, idxComprar, idxPrecoPorFator, idxUN].includes(-1)) {
      throw new Error("Uma ou mais colunas essenciais (SubProduto, ID da Cotação, Data, Preço, Comprar, Preço por Fator, UN) não foram encontradas na aba Cotações.");
    }

    const produtosDaCotacao = new Set();
    todosOsDados.forEach(linha => {
      if (String(linha[idxIdCotacao]) === String(idCotacaoAlvo)) {
        produtosDaCotacao.add(linha[idxSubProduto]);
      }
    });

    const relatorioFinal = [];
    const seisMesesAtras = new Date();
    seisMesesAtras.setMonth(seisMesesAtras.getMonth() - 6);

    for (const produto of produtosDaCotacao) {
      const historicoProduto = todosOsDados
        .filter(linha => linha[idxSubProduto] === produto && linha[idxComprar] > 0)
        .map(linha => ({
          data: new Date(linha[idxDataAbertura]),
          preco: parseFloat(linha[idxPreco]) || 0,
          quantidade: parseFloat(linha[idxComprar]) || 0,
          precoPorFator: parseFloat(linha[idxPrecoPorFator]) || 0,
          un: linha[idxUN] || 'N/A'
        }))
        .filter(item => !isNaN(item.data.getTime()))
        .sort((a, b) => b.data - a.data);

      if (historicoProduto.length === 0) continue;
      
      const unidadeDoProduto = historicoProduto[0].un;

      const ultimos6Pedidos = historicoProduto.slice(0, 6).map(p => ({
        data: p.data.toLocaleDateString('pt-BR'),
        preco: p.preco,
        precoPorFator: p.precoPorFator
      }));
      
      const historico6Meses = historicoProduto.filter(p => p.data >= seisMesesAtras);
      let precoFatorMin6M = null, precoFatorMax6M = null, precoFatorMedioPonderado6M = null;

      if (historico6Meses.length > 0) {
        const somaQuantidade = historico6Meses.reduce((acc, p) => acc + p.quantidade, 0);
        
        precoFatorMin6M = Math.min(...historico6Meses.map(p => p.precoPorFator));
        precoFatorMax6M = Math.max(...historico6Meses.map(p => p.precoPorFator));
        const somaValorTotalFator = historico6Meses.reduce((acc, p) => acc + (p.precoPorFator * p.quantidade), 0);
        precoFatorMedioPonderado6M = somaQuantidade > 0 ? somaValorTotalFator / somaQuantidade : 0;
      }
      
      const volumesPorPedido = historicoProduto.map(p => ({
          data: p.data.toLocaleDateString('pt-BR'),
          quantidade: p.quantidade,
          un: p.un
      }));
      
      let intervaloMedioDias = null;
      if (historicoProduto.length > 1) {
        let somaDiferencasDias = 0;
        for (let i = 0; i < historicoProduto.length - 1; i++) {
          somaDiferencasDias += (historicoProduto[i].data - historicoProduto[i+1].data) / (1000 * 60 * 60 * 24);
        }
        intervaloMedioDias = somaDiferencasDias / (historicoProduto.length - 1);
      }
      
      relatorioFinal.push({
        produto: produto,
        unidade: unidadeDoProduto,
        ultimos6Pedidos: ultimos6Pedidos,
        precoFatorMin6M: precoFatorMin6M,
        precoFatorMax6M: precoFatorMax6M,
        precoFatorMedioPonderado6M: precoFatorMedioPonderado6M,
        volumesPorPedido: volumesPorPedido,
        intervaloMedioDias: intervaloMedioDias
      });
    }
    
    return relatorioFinal;

  } catch (error) {
    Logger.log(`ERRO em RelatoriosCRUD_gerarDadosAnaliseCompra: ${error.toString()} Stack: ${error.stack}`);
    return null;
  }
}
