/**
 * @file RateioCrud.gs
 * @description Funções CRUD para o módulo de Rateio Financeiro.
 */

/**
 * Função "à prova de balas" para converter valores da planilha em números.
 * Lida com "R$", separador de milhar "." e decimal ",".
 * @param {*} valor O valor da célula.
 * @returns {number} O valor convertido para número.
 */
function _RateioCrud_parsearValorNumerico(valor) {
  if (typeof valor === 'number') return valor;
  if (valor === null || valor === undefined || String(valor).trim() === '') return 0;

  let strValor = String(valor).trim();
  strValor = strValor.replace(/R\$\s*/, '');
  strValor = strValor.replace(/\.(?=\d{3})/g, '');
  strValor = strValor.replace(',', '.');
  
  const numero = parseFloat(strValor);
  return isNaN(numero) ? 0 : numero;
}

/**
 * Busca na planilha de Notas Fiscais todas as NFs que já foram conciliadas
 * mas que ainda não possuem um "Status do Rateio".
 */
function RateioCrud_obterNotasParaRatear() {
  try {
    const planilhaNF = SpreadsheetApp.openById(ID_PLANILHA_NF);
    const aba = planilhaNF.getSheetByName(ABA_NF_NOTAS_FISCAIS);
    const dados = aba.getDataRange().getValues();

    const cabecalhos = dados.shift();
    const colunas = {
      statusConciliacao: cabecalhos.indexOf("Status da Conciliação"),
      statusRateio: cabecalhos.indexOf("Status do Rateio"),
      chaveAcesso: cabecalhos.indexOf("Chave de Acesso"),
      nomeEmitente: cabecalhos.indexOf("Nome Emitente"),
      numeroNF: cabecalhos.indexOf("Número NF"),
      dataEmissao: cabecalhos.indexOf("Data e Hora Emissão")
    };

    const notasParaRatear = [];
    dados.forEach(linha => {
      if (linha[colunas.statusConciliacao] && !linha[colunas.statusRateio]) {
        notasParaRatear.push({
          chaveAcesso: linha[colunas.chaveAcesso],
          nomeEmitente: linha[colunas.nomeEmitente],
          numeroNF: linha[colunas.numeroNF],
          dataEmissao: new Date(linha[colunas.dataEmissao]).toLocaleDateString('pt-BR')
        });
      }
    });
    return notasParaRatear;
  } catch (e) {
    Logger.log(`Erro em RateioCrud_obterNotasParaRatear: ${e.message}`);
    return [];
  }
}

/**
 * Obtém os itens de uma NF específica pela chave de acesso.
 */
function RateioCrud_obterItensDaNF(chaveAcesso) {
  const planilhaNF = SpreadsheetApp.openById(ID_PLANILHA_NF);
  const aba = planilhaNF.getSheetByName(ABA_NF_ITENS);
  const dados = aba.getDataRange().getValues();
  const cabecalhos = dados.shift();
  const colunas = {
      chaveAcesso: cabecalhos.indexOf("Chave de Acesso"),
      descricaoProduto: cabecalhos.indexOf("Descrição Produto (NF)"),
      valorTotalBrutoItem: cabecalhos.indexOf("Valor Total Bruto Item"),
      numeroItem: cabecalhos.indexOf("Número do Item")
  };

  return dados
    .filter(linha => linha[colunas.chaveAcesso] === chaveAcesso)
    .map(linha => ({
      descricaoProduto: linha[colunas.descricaoProduto],
      valorTotalBrutoItem: _RateioCrud_parsearValorNumerico(linha[colunas.valorTotalBrutoItem]),
      numeroItem: linha[colunas.numeroItem]
    }));
}

/**
 * Obtém os dados de totais de tributos de uma NF específica.
 */
function RateioCrud_obterTotaisDaNF(chaveAcesso) {
  const planilhaNF = SpreadsheetApp.openById(ID_PLANILHA_NF);
  const aba = planilhaNF.getSheetByName(ABA_NF_TRIBUTOS_TOTAIS);
  const dados = aba.getDataRange().getValues();
  const cabecalhos = dados.shift();
  const colChave = cabecalhos.indexOf("Chave de Acesso");

  const linhaTotal = dados.find(linha => linha[colChave] === chaveAcesso);
  if (!linhaTotal) throw new Error(`Totais não encontrados para a chave ${chaveAcesso}`);
  
  const totais = {};
  cabecalhos.forEach((cabecalho, index) => {
    const chaveObjeto = cabecalho.replace(/\s+/g, '');
    if (chaveObjeto !== "ChavedeAcesso") {
      totais[chaveObjeto] = _RateioCrud_parsearValorNumerico(linhaTotal[index]);
    }
  });

  return {
    totalValorProdutos: totais.TotalValorProdutos,
    totalValorFrete: totais.TotalValorFrete,
    totalValorIcmsSt: totais.TotalValorICMSST,
    totalValorIpi: totais.TotalValorIPI,
    totalOutrasDespesas: totais.TotalOutrasDespesas,
    totalValorSeguro: totais.TotalValorSeguro,
    totalValorDesconto: totais.TotalValorDesconto,
    valorTotalNf: totais.ValorTotaldaNF
  };
}

/**
 * Obtém as faturas (boletos) de uma NF específica.
 */
function RateioCrud_obterFaturasDaNF(chaveAcesso) {
  const planilhaNF = SpreadsheetApp.openById(ID_PLANILHA_NF);
  const aba = planilhaNF.getSheetByName(ABA_NF_FATURAS);
  const dados = aba.getDataRange().getValues();
  const cabecalhos = dados.shift();
   const colunas = {
      chaveAcesso: cabecalhos.indexOf("Chave de Acesso"),
      numeroFatura: cabecalhos.indexOf("Número da Fatura"),
      numeroParcela: cabecalhos.indexOf("Número da Parcela"),
      dataVencimento: cabecalhos.indexOf("Data de Vencimento"),
      valorParcela: cabecalhos.indexOf("Valor da Parcela")
  };

  return dados
    .filter(linha => linha[colunas.chaveAcesso] === chaveAcesso)
    .map(linha => {
      const dataVenc = linha[colunas.dataVencimento];
      return {
        numeroFatura: linha[colunas.numeroFatura],
        numeroParcela: linha[colunas.numeroParcela],
        dataVencimento: dataVenc instanceof Date ? dataVenc.toISOString() : dataVenc,
        valorParcela: _RateioCrud_parsearValorNumerico(linha[colunas.valorParcela])
      }
  });
}

/**
 * Obtém todas as regras de rateio da planilha Financeiro.
 */
function RateioCrud_obterRegrasRateio() {
  const planilhaFin = SpreadsheetApp.openById(ID_PLANILHA_FINANCEIRO);
  const aba = planilhaFin.getSheetByName(ABA_FINANCEIRO_REGRAS_RATEIO);
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return [];

  const dados = aba.getRange(1, 1, ultimaLinha, aba.getLastColumn()).getValues();
  const cabecalhos = dados.shift();
  const colunas = {
      itemCotacao: cabecalhos.indexOf("Item da Cotação"),
      setor: cabecalhos.indexOf("Setor"),
      porcentagem: cabecalhos.indexOf("Porcentagem")
  };
  
  return dados
    .map(linha => ({
      itemCotacao: linha[colunas.itemCotacao],
      setor: linha[colunas.setor],
      porcentagem: _RateioCrud_parsearValorNumerico(linha[colunas.porcentagem])
    }))
    .filter(r => r.itemCotacao && r.setor);
}


/**
 * Salva as linhas de rateio na aba 'ContasAPagar'.
 */
function RateioCrud_salvarContasAPagar(linhasParaAdicionar) {
  if (!linhasParaAdicionar || linhasParaAdicionar.length === 0) return;

  const planilhaFin = SpreadsheetApp.openById(ID_PLANILHA_FINANCEIRO);
  const aba = planilhaFin.getSheetByName(ABA_FINANCEIRO_CONTAS_A_PAGAR);
  const cabecalhos = CABECALHOS_FINANCEIRO_CONTAS_A_PAGAR;
  
  const chaves = cabecalhos.map(c => c.replace(/\s+/g, ''));
  const dadosFormatados = linhasParaAdicionar.map(obj => 
      chaves.map(chave => obj[chave] ?? '')
  );
  
  aba.getRange(aba.getLastRow() + 1, 1, dadosFormatados.length, dadosFormatados[0].length)
     .setValues(dadosFormatados);
}

/**
 * Atualiza o status do rateio de uma NF para "Concluído".
 */
function RateioCrud_atualizarStatusRateio(chaveAcesso, novoStatus) {
  const planilhaNF = SpreadsheetApp.openById(ID_PLANILHA_NF);
  const aba = planilhaNF.getSheetByName(ABA_NF_NOTAS_FISCAIS);
  const dados = aba.getDataRange().getValues();
  const cabecalhos = dados[0];
  const colChave = cabecalhos.indexOf("Chave de Acesso");
  const colStatusRateio = cabecalhos.indexOf("Status do Rateio");

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][colChave] === chaveAcesso) {
      aba.getRange(i + 1, colStatusRateio + 1).setValue(novoStatus);
      Logger.log(`Status do rateio para a chave ${chaveAcesso} atualizado para "${novoStatus}".`);
      return; 
    }
  }
}

/**
 * Salva novas regras de rateio na planilha, evitando duplicatas.
 */
function RateioCrud_salvarNovasRegrasDeRateio(novasRegras) {
  if (!novasRegras || novasRegras.length === 0) return;

  const planilhaFin = SpreadsheetApp.openById(ID_PLANILHA_FINANCEIRO);
  const aba = planilhaFin.getSheetByName(ABA_FINANCEIRO_REGRAS_RATEIO);
  
  const dadosAtuais = aba.getDataRange().getValues();
  // Cria um Set para checagem rápida de duplicatas. Chave: "ItemDaCotacao#Setor"
  const regrasExistentes = new Set(dadosAtuais.map(linha => `${linha[0]}#${linha[1]}`));

  const linhasParaAdicionar = [];
  novasRegras.forEach(regra => {
    const chaveUnica = `${regra.itemCotacao}#${regra.setor}`;
    if (!regrasExistentes.has(chaveUnica)) {
      linhasParaAdicionar.push([regra.itemCotacao, regra.setor, regra.porcentagem]);
      regrasExistentes.add(chaveUnica); // Adiciona ao Set para evitar duplicatas no mesmo lote
    }
  });

  if (linhasParaAdicionar.length > 0) {
    aba.getRange(aba.getLastRow() + 1, 1, linhasParaAdicionar.length, 3)
       .setValues(linhasParaAdicionar);
    Logger.log(`${linhasParaAdicionar.length} nova(s) regra(s) de rateio foram salvas.`);
  }
}