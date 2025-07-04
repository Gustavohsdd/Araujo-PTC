/**
 * @file RateioController.gs
 * @description Orquestra a lógica do módulo de rateio.
 */

function RateioController_obterDadosParaPagina() {
  try {
    const notas = RateioCrud_obterNotasParaRatear();
    return { success: true, notas: notas };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function RateioController_analisarNFParaRateio(chaveAcesso) {
  try {
    const itensDaNf = RateioCrud_obterItensDaNF(chaveAcesso);
    const totaisDaNf = RateioCrud_obterTotaisDaNF(chaveAcesso);
    const mapaConciliacao = ConciliacaoNFCrud_obterMapeamentoConciliacao(); 
    const regrasRateio = RateioCrud_obterRegrasRateio();

    const totalProdutos = totaisDaNf.totalValorProdutos;
    
    if (itensDaNf.length > 0) {
      if (totalProdutos > 0.001) {
        const custosAdicionais = totaisDaNf.valorTotalNf - totalProdutos;
        itensDaNf.forEach(item => {
          const proporcaoItem = item.valorTotalBrutoItem / totalProdutos;
          item.custoEfetivo = item.valorTotalBrutoItem + (custosAdicionais * proporcaoItem);
        });
      } else {
        const custoPorItem = totaisDaNf.valorTotalNf / itensDaNf.length;
        itensDaNf.forEach(item => item.custoEfetivo = custoPorItem);
      }
    }

    const itensRateados = [];
    const itensSemRegra = [];
    let totaisPorSetor = {};

    const mapaDeRegras = regrasRateio.reduce((map, regra) => {
        if (!map[regra.itemCotacao]) map[regra.itemCotacao] = [];
        map[regra.itemCotacao].push({ setor: regra.setor, porcentagem: regra.porcentagem });
        return map;
    }, {});
    
    itensDaNf.forEach(item => {
      const mapeamento = mapaConciliacao.find(m => m.descricaoNF === item.descricaoProduto);
      const itemCotacao = mapeamento ? mapeamento.itemCotacao : null;
      const regrasDoItem = itemCotacao ? mapaDeRegras[itemCotacao] : null;
      
      // REQUISITO 1: Adiciona o itemCotacao ao objeto para ser exibido no frontend
      const itemInfo = { 
        descricao: item.descricaoProduto, 
        custoTotal: item.custoEfetivo, 
        numeroItem: item.numeroItem,
        itemCotacao: itemCotacao // Adicionado aqui
      };

      if (regrasDoItem && regrasDoItem.length > 0) {
        const rateiosDoItem = [];
        regrasDoItem.forEach(regra => {
          const valorRateado = item.custoEfetivo * (regra.porcentagem / 100);
          rateiosDoItem.push({ setor: regra.setor, valor: valorRateado });
          totaisPorSetor[regra.setor] = (totaisPorSetor[regra.setor] || 0) + valorRateado;
        });
        itensRateados.push({ ...itemInfo, rateios: rateiosDoItem });
      } else {
        itensSemRegra.push(itemInfo);
      }
    });

    const faturas = RateioCrud_obterFaturasDaNF(chaveAcesso);

    return {
      success: true,
      dados: {
        valorTotalNF: totaisDaNf.valorTotalNf,
        itensRateados: itensRateados,
        itensSemRegra: itensSemRegra,
        totaisPorSetor: totaisPorSetor,
        faturas: faturas
      }
    };

  } catch (e) {
    Logger.log(`ERRO CRÍTICO em RateioController_analisarNFParaRateio: ${e.message}\n${e.stack}`);
    return { success: false, message: e.message };
  }
}

function RateioController_salvarRateioFinal(dadosRateio) {
    try {
        const { chaveAcesso, faturas, totaisPorSetor, novasRegras, numeroNF, nomeEmitente } = dadosRateio;

        // REQUISITO 2: Salva as novas regras de rateio criadas manualmente
        if (novasRegras && novasRegras.length > 0) {
          RateioCrud_salvarNovasRegrasDeRateio(novasRegras);
        }

        let valorTotalNF = Object.values(totaisPorSetor).reduce((s, v) => s + v, 0);
        valorTotalNF = parseFloat(valorTotalNF.toFixed(2));

        const porcentagensPorSetor = {};
        if (valorTotalNF > 0) {
          for (const setor in totaisPorSetor) {
              porcentagensPorSetor[setor] = totaisPorSetor[setor] / valorTotalNF;
          }
        }

        const linhasParaContasAPagar = [];
        if (faturas && faturas.length > 0) {
          faturas.forEach(fatura => {
              for (const setor in porcentagensPorSetor) {
                  linhasParaContasAPagar.push({
                      ChavedeAcesso: chaveAcesso, NúmerodaFatura: fatura.numeroFatura,
                      NúmerodaParcela: fatura.numeroParcela, ResumodosItens: `NF ${numeroNF} - ${nomeEmitente}`,
                      DatadeVencimento: new Date(fatura.dataVencimento), // Converte de volta para Date
                      ValordaParcela: fatura.valorParcela,
                      Setor: setor, ValorporSetor: fatura.valorParcela * porcentagensPorSetor[setor]
                  });
              }
          });
        } else {
           for (const setor in porcentagensPorSetor) {
                  linhasParaContasAPagar.push({
                      ChavedeAcesso: chaveAcesso, NúmerodaFatura: numeroNF,
                      NúmerodaParcela: 1, ResumodosItens: `NF ${numeroNF} - ${nomeEmitente}`,
                      DatadeVencimento: new Date(), ValordaParcela: valorTotalNF,
                      Setor: setor, ValorporSetor: valorTotalNF * porcentagensPorSetor[setor]
                  });
              }
        }

        RateioCrud_salvarContasAPagar(linhasParaContasAPagar);
        RateioCrud_atualizarStatusRateio(chaveAcesso, "Concluído");

        return { success: true, message: "Rateio e novas regras salvos com sucesso!" };

    } catch (e) {
        Logger.log(`ERRO em RateioController_salvarRateioFinal: ${e.message}\n${e.stack}`);
        return { success: false, message: e.message };
    }
}