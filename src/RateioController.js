/**
 * @file RateioController.gs
 * @description Orquestra a lógica do módulo de rateio.
 */

/**
 * Obtém os dados iniciais para a página, agora com uma pré-análise
 * para identificar notas que podem ser rateadas 100% automaticamente.
 */
function RateioController_obterDadosParaPagina() {
  try {
    let notas = RateioCrud_obterNotasParaRatear();
    const mapaConciliacao = ConciliacaoNFCrud_obterMapeamentoConciliacao(); 
    const regrasRateio = RateioCrud_obterRegrasRateio();
    
    // Cria um mapa de regras para checagem rápida
    const mapaDeRegras = regrasRateio.reduce((map, regra) => {
        if (!map[regra.itemCotacao]) map[regra.itemCotacao] = true;
        return map;
    }, {});

    // Adiciona uma flag 'isAutomatico' para cada nota
    notas = notas.map(nota => {
      const itensDaNf = RateioCrud_obterItensDaNF(nota.chaveAcesso);
      // Se a nota não tiver itens, não pode ser automática
      if (itensDaNf.length === 0) {
        nota.isAutomatico = false;
        return nota;
      }
      
      // Verifica se todos os itens da nota possuem uma regra de rateio correspondente
      const todosItensTemRegra = itensDaNf.every(item => {
        const mapeamento = mapaConciliacao.find(m => m.descricaoNF === item.descricaoProduto);
        const itemCotacao = mapeamento ? mapeamento.itemCotacao : null;
        return itemCotacao ? mapaDeRegras[itemCotacao] : false;
      });

      nota.isAutomatico = todosItensTemRegra;
      return nota;
    });

    return { success: true, notas: notas };
  } catch (e) {
    Logger.log(`Erro em RateioController_obterDadosParaPagina: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * Analisa uma NF específica, aplica as regras de rateio e identifica
 * itens que precisam de rateio manual. Retorna se o rateio foi 100% automático.
 */
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
      
      const itemInfo = { 
        descricao: item.descricaoProduto, 
        custoTotal: item.custoEfetivo, 
        numeroItem: item.numeroItem,
        itemCotacao: itemCotacao
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
        faturas: faturas,
        rateioCompleto: itensSemRegra.length === 0 
      }
    };

  } catch (e) {
    Logger.log(`ERRO CRÍTICO em RateioController_analisarNFParaRateio: ${e.message}\n${e.stack}`);
    return { success: false, message: e.message };
  }
}

/**
 * Recebe uma lista de chaves de acesso, analisa cada uma e retorna um
 * array de payloads prontos para serem adicionados ao lote.
 */
function RateioController_prepararLoteAutomatico(chavesDeAcesso) {
  const lotePayloads = [];
  chavesDeAcesso.forEach(chave => {
    try {
      const resultadoAnalise = RateioController_analisarNFParaRateio(chave);
      if (resultadoAnalise.success && resultadoAnalise.dados.rateioCompleto) {
        const dados = resultadoAnalise.dados;
        // Busca info básica da nota para o payload
        const notaInfo = RateioCrud_obterNotasParaRatear().find(n => n.chaveAcesso === chave);
        
        const mapaSetorParaItens = {};
        dados.itensRateados.forEach(item => {
          item.rateios.forEach(rateio => {
            if (item.itemCotacao) {
              if (!mapaSetorParaItens[rateio.setor]) mapaSetorParaItens[rateio.setor] = new Set();
              mapaSetorParaItens[rateio.setor].add(item.itemCotacao);
            }
          });
        });
         for(const setor in mapaSetorParaItens){
          mapaSetorParaItens[setor] = Array.from(mapaSetorParaItens[setor]);
        }

        lotePayloads.push({
          chaveAcesso: chave, 
          faturas: dados.faturas, 
          totaisPorSetor: dados.totaisPorSetor, 
          novasRegras: [], // Rateios automáticos não geram novas regras
          mapaSetorParaItens: mapaSetorParaItens,
          numeroNF: notaInfo.numeroNF, 
          nomeEmitente: notaInfo.nomeEmitente
        });
      }
    } catch (e) {
      Logger.log(`Erro ao preparar rateio automático para a chave ${chave}: ${e.message}`);
    }
  });
  return { success: true, payloads: lotePayloads };
}

/**
 * Salva um lote de rateios (manuais e automáticos) de uma só vez.
 */
function RateioController_salvarRateioEmLote(loteDeRateios) {
    try {
        if (!loteDeRateios || loteDeRateios.length === 0) {
          throw new Error("Nenhum dado de rateio recebido.");
        }

        const todasAsLinhasContasAPagar = [];
        const todasAsNovasRegras = [];
        const todasAsChavesParaAtualizar = new Set();

        loteDeRateios.forEach(dadosRateio => {
            const { chaveAcesso, faturas, totaisPorSetor, novasRegras, numeroNF, nomeEmitente, mapaSetorParaItens } = dadosRateio;
            todasAsChavesParaAtualizar.add(chaveAcesso);

            if (novasRegras && novasRegras.length > 0) {
              todasAsNovasRegras.push(...novasRegras);
            }

            let valorTotalNF = Object.values(totaisPorSetor).reduce((s, v) => s + v, 0);
            valorTotalNF = parseFloat(valorTotalNF.toFixed(2));

            const porcentagensPorSetor = {};
            if (valorTotalNF > 0) {
              for (const setor in totaisPorSetor) {
                  porcentagensPorSetor[setor] = totaisPorSetor[setor] / valorTotalNF;
              }
            }
            
            const numFaturasOriginais = faturas.length > 0 ? faturas.length : 1;
            const numSetores = Object.keys(porcentagensPorSetor).length;
            const totalNovosTitulosNota = numFaturasOriginais * numSetores;
            let contadorParcelaNota = 1;

            if (faturas && faturas.length > 0) {
              faturas.forEach(fatura => {
                  for (const setor in porcentagensPorSetor) {
                      const resumoItens = mapaSetorParaItens[setor] ? mapaSetorParaItens[setor].join(', ') : `NF ${numeroNF}`;
                      todasAsLinhasContasAPagar.push({
                          ChavedeAcesso: chaveAcesso, NúmerodaFatura: fatura.numeroFatura,
                          NúmerodaParcela: `${contadorParcelaNota++}/${totalNovosTitulosNota}(${numFaturasOriginais})`, 
                          ResumodosItens: resumoItens,
                          DatadeVencimento: new Date(fatura.dataVencimento),
                          ValordaParcela: fatura.valorParcela,
                          Setor: setor, ValorporSetor: fatura.valorParcela * porcentagensPorSetor[setor]
                      });
                  }
              });
            } else {
               for (const setor in porcentagensPorSetor) {
                      const resumoItens = mapaSetorParaItens[setor] ? mapaSetorParaItens[setor].join(', ') : `NF ${numeroNF}`;
                      todasAsLinhasContasAPagar.push({
                          ChavedeAcesso: chaveAcesso, NúmerodaFatura: numeroNF,
                          NúmerodaParcela: `${contadorParcelaNota++}/${totalNovosTitulosNota}(1)`, 
                          ResumodosItens: resumoItens,
                          DatadeVencimento: new Date(), ValordaParcela: valorTotalNF,
                          Setor: setor, ValorporSetor: valorTotalNF * porcentagensPorSetor[setor]
                      });
                  }
            }
        });

        RateioCrud_salvarNovasRegrasDeRateio(todasAsNovasRegras);
        RateioCrud_salvarContasAPagar(todasAsLinhasContasAPagar);
        Array.from(todasAsChavesParaAtualizar).forEach(chave => {
          RateioCrud_atualizarStatusRateio(chave, "Concluído");
        });

        return { success: true, message: `${loteDeRateios.length} rateio(s) salvos com sucesso!` };

    } catch (e) {
        Logger.log(`ERRO em RateioController_salvarRateioEmLote: ${e.message}\n${e.stack}`);
        return { success: false, message: e.message };
    }
}