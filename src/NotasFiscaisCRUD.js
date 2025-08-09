/**
 * @file NotasFiscaisCRUD.js
 * @description Funções de Create, Read, Update, Delete para o módulo de gerenciamento de Notas Fiscais.
 */

/**
 * Obtém todas as notas fiscais da planilha, aplicando filtros de forma performática.
 * Lê a aba 'NotasFiscais' como base e busca dados complementares em outras abas usando mapas para eficiência.
 * @param {object} filtros - Objeto com { status, fornecedor, dataInicio, dataFim, busca }.
 * @returns {Array<object>} Um array de objetos, cada um representando uma NF.
 */
function NotasFiscaisCRUD_obterTodasAsNotas(filtros = {}) {
  const planilhaNF = SpreadsheetApp.openById(ID_PLANILHA_NF);
  const abaNF = planilhaNF.getSheetByName(ABA_NF_NOTAS_FISCAIS);
  if (!abaNF) {
    throw new Error('Aba base de Notas Fiscais não encontrada: ' + ABA_NF_NOTAS_FISCAIS);
  }
  if (abaNF.getLastRow() < 2) return [];

  // Cabeçalhos REAIS da planilha (linha 1)
  const cabecalhosNF = abaNF.getRange(1, 1, 1, abaNF.getLastColumn()).getValues()[0];
  const colMapNF = cabecalhosNF.reduce((acc, h, i) => { acc[h] = i; return acc; }, {});

  const dadosNF = abaNF.getRange(2, 1, abaNF.getLastRow() - 1, abaNF.getLastColumn()).getValues();

  // ---- Tributos (opcional) ----
  const abaTributos = planilhaNF.getSheetByName(ABA_NF_TRIBUTOS_TOTAIS);
  let dadosTributos = [], cabecalhosTributos = [], colChaveTrib = -1, colValorTotal = -1;
  if (abaTributos && abaTributos.getLastRow() > 1) {
    cabecalhosTributos = abaTributos.getRange(1, 1, 1, abaTributos.getLastColumn()).getValues()[0];
    const valoresTributos = abaTributos.getRange(2, 1, abaTributos.getLastRow() - 1, abaTributos.getLastColumn()).getValues();
    colChaveTrib = cabecalhosTributos.indexOf('Chave de Acesso');
    colValorTotal = cabecalhosTributos.indexOf('Valor Total da NF');
    if (colChaveTrib > -1 && colValorTotal > -1) {
      dadosTributos = valoresTributos;
    }
  }
  const mapaValoresTotais = {};
  if (dadosTributos.length && colChaveTrib > -1 && colValorTotal > -1) {
    dadosTributos.forEach(linha => {
      const k = linha[colChaveTrib];
      if (k) mapaValoresTotais[k] = linha[colValorTotal] || 0;
    });
  }

  // ---- Faturas (opcional) ----
  const abaFaturas = planilhaNF.getSheetByName(ABA_NF_FATURAS);
  let dadosFaturas = [], cabecalhosFaturas = [], colChaveFat = -1;
  if (abaFaturas && abaFaturas.getLastRow() > 1) {
    cabecalhosFaturas = abaFaturas.getRange(1, 1, 1, abaFaturas.getLastColumn()).getValues()[0];
    const valoresFaturas = abaFaturas.getRange(2, 1, abaFaturas.getLastRow() - 1, abaFaturas.getLastColumn()).getValues();
    colChaveFat = cabecalhosFaturas.indexOf('Chave de Acesso');
    if (colChaveFat > -1) {
      dadosFaturas = valoresFaturas;
    }
  }
  const mapaContagemFaturas = {};
  if (dadosFaturas.length && colChaveFat > -1) {
    dadosFaturas.forEach(linha => {
      const k = linha[colChaveFat];
      if (k) mapaContagemFaturas[k] = (mapaContagemFaturas[k] || 0) + 1;
    });
  }

  // Filtros de data
  const dataInicioFiltro = filtros.dataInicio ? new Date(filtros.dataInicio) : null;
  const dataFimFiltro = filtros.dataFim ? new Date(filtros.dataFim) : null;
  if (dataFimFiltro && !isNaN(dataFimFiltro)) dataFimFiltro.setHours(23, 59, 59, 999);

  // Colunas esperadas por nome (tolerantes a ordem)
  const idxDataEmissao = colMapNF['Data e Hora Emissão'];
  const idxStatusConc = colMapNF['Status da Conciliação'];
  const idxNomeEmit = colMapNF['Nome Emitente'];
  const idxNumeroNF = colMapNF['Número NF'];
  const idxCNPJEmit = colMapNF['CNPJ Emitente'];
  const idxChave = colMapNF['Chave de Acesso'];

  const resultados = [];
  for (let i = 0; i < dadosNF.length; i++) {
    const linha = dadosNF[i];

    // Data Emissão (tolerante)
    let dataEmissao = null;
    if (idxDataEmissao > -1) {
      const raw = linha[idxDataEmissao];
      const d = raw instanceof Date ? raw : new Date(raw);
      if (!isNaN(d)) dataEmissao = d;
    }

    // 1) Filtro de Data
    if (dataInicioFiltro && dataEmissao && dataEmissao < dataInicioFiltro) continue;
    if (dataFimFiltro && dataEmissao && dataEmissao > dataFimFiltro) continue;

    // 2) Filtro de Status
    const statusNF = (idxStatusConc > -1 ? (linha[idxStatusConc] || '') : '');
    if (filtros.status && statusNF !== filtros.status) continue;

    // 3) Campos textuais
    const nomeEmitente = (idxNomeEmit > -1 ? (linha[idxNomeEmit] || '') : '');
    const numeroNF = String(idxNumeroNF > -1 ? (linha[idxNumeroNF] || '') : '');
    const cnpjEmitente = (idxCNPJEmit > -1 ? (linha[idxCNPJEmit] || '') : '');
    const chaveAcesso = (idxChave > -1 ? (linha[idxChave] || '') : '');

    // 4) Filtro de busca textual
    if (filtros.busca) {
      const termo = String(filtros.busca).toLowerCase();
      const hit = (nomeEmitente.toLowerCase().includes(termo) ||
                   numeroNF.toLowerCase().includes(termo) ||
                   cnpjEmitente.toLowerCase().includes(termo) ||
                   chaveAcesso.toLowerCase().includes(termo));
      if (!hit) continue;
    }

    // 5) Filtro fornecedor
    if (filtros.fornecedor) {
      const f = String(filtros.fornecedor).toLowerCase();
      if (!nomeEmitente.toLowerCase().includes(f)) continue;
    }

    resultados.push({
      chaveAcesso: chaveAcesso,
      numeroNF: numeroNF,
      nomeEmitente: nomeEmitente,
      cnpjEmitente: cnpjEmitente,
      dataEmissao: dataEmissao
        ? dataEmissao.toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' })
        : '',
      statusConciliacao: statusNF || '',
      valorTotalNF: mapaValoresTotais[chaveAcesso] || 0,
      faturasCount: mapaContagemFaturas[chaveAcesso] || 0
    });

    if (resultados.length >= 200) break;
  }

  return resultados;
}

/**
 * Reseta o status de uma NF para 'Pendente' e limpa campos relacionados.
 * @param {string} chaveAcesso A chave de acesso da NF.
 */
function NotasFiscaisCRUD_resetarStatusNF(chaveAcesso) {
  const planilhaNF = SpreadsheetApp.openById(ID_PLANILHA_NF);
  const abaNF = planilhaNF.getSheetByName(ABA_NF_NOTAS_FISCAIS);
  const dados = abaNF.getDataRange().getValues();
  const cabecalhos = dados[0];
  const colChave = cabecalhos.indexOf("Chave de Acesso");
  const colStatusConc = cabecalhos.indexOf("Status da Conciliação");
  const colStatusRateio = cabecalhos.indexOf("Status do Rateio");
  const colIdCotacao = cabecalhos.indexOf("ID da Cotação (Sistema)");

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][colChave] === chaveAcesso) {
      abaNF.getRange(i + 1, colStatusConc + 1).setValue("Pendente");
      abaNF.getRange(i + 1, colStatusRateio + 1).setValue("Pendente");
      abaNF.getRange(i + 1, colIdCotacao + 1).clearContent();
      Logger.log(`CRUD: Status da NF ${chaveAcesso} resetado na linha ${i + 1}.`);
      return;
    }
  }
  throw new Error(`NF com chave ${chaveAcesso} não encontrada para resetar o status.`);
}

/**
 * Exclui todas as linhas de contas a pagar associadas a uma chave de acesso.
 * @param {string} chaveAcesso A chave de acesso da NF.
 */
function NotasFiscaisCRUD_excluirContasAPagarPorChave(chaveAcesso) {
  const planilhaFin = SpreadsheetApp.openById(ID_PLANILHA_FINANCEIRO);
  const abaFin = planilhaFin.getSheetByName(ABA_FINANCEIRO_CONTAS_A_PAGAR);
  if(abaFin.getLastRow() < 2) return;

  const dados = abaFin.getDataRange().getValues();
  const colChave = CABECALHOS_FINANCEIRO_CONTAS_A_PAGAR.indexOf("Chave de Acesso");
  
  // Itera de baixo para cima para evitar problemas com a remoção de linhas
  for (let i = dados.length - 1; i >= 1; i--) {
    if (dados[i][colChave] === chaveAcesso) {
      abaFin.deleteRow(i + 1);
      Logger.log(`CRUD: Linha ${i + 1} de Contas a Pagar excluída para a chave ${chaveAcesso}.`);
    }
  }
}

/**
 * Obtém faturas e contas a pagar de uma NF específica.
 * @param {string} chaveAcesso A chave de acesso da NF.
 * @returns {{faturas: object[], contasAPagar: object[]}}
 */
function NotasFiscaisCRUD_obterResumoFinanceiroDaNF(chaveAcesso) {
  const planilhaNF = SpreadsheetApp.openById(ID_PLANILHA_NF);
  const abaFaturas = planilhaNF.getSheetByName(ABA_NF_FATURAS);
  let faturas = [];
  if (abaFaturas && abaFaturas.getLastRow() > 1) {
    const dadosFaturas = abaFaturas.getDataRange().getValues();
    const cabecalhosFaturas = dadosFaturas.shift() || [];
    faturas = dadosFaturas
      .filter(l => l[0] === chaveAcesso)
      .map(l => {
        const obj = {};
        cabecalhosFaturas.forEach((c, i) => {
          let v = l[i];
          if (v instanceof Date) v = v.toISOString().split('T')[0];
          obj[c] = v;
        });
        return obj;
      });
  }

  const planilhaFin = SpreadsheetApp.openById(ID_PLANILHA_FINANCEIRO);
  const abaContas = planilhaFin.getSheetByName(ABA_FINANCEIRO_CONTAS_A_PAGAR);
  let contasAPagar = [];
  if (abaContas && abaContas.getLastRow() > 1) {
    const dadosContas = abaContas.getDataRange().getValues();
    const cabecalhosContas = dadosContas.shift() || [];
    contasAPagar = dadosContas
      .filter(l => l[0] === chaveAcesso)
      .map(l => {
        const obj = {};
        cabecalhosContas.forEach((c, i) => {
          let v = l[i];
          if (v instanceof Date) v = v.toISOString().split('T')[0];
          obj[c] = v;
        });
        return obj;
      });
  }

  return { faturas, contasAPagar };
}

/**
 * Substitui todas as faturas de uma NF por um novo conjunto.
 * @param {string} chaveAcesso A chave de acesso da NF.
 * @param {Array<object>} faturas O novo array de objetos de fatura.
 */
function NotasFiscaisCRUD_substituirFaturasDaNF(chaveAcesso, faturas) {
  const planilhaNF = SpreadsheetApp.openById(ID_PLANILHA_NF);
  const aba = planilhaNF.getSheetByName(ABA_NF_FATURAS);
  const dados = aba.getDataRange().getValues();
  const colChave = CABECALHOS_NF_FATURAS.indexOf("Chave de Acesso");

  // Excluir antigas
  for (let i = dados.length - 1; i >= 1; i--) {
    if (dados[i][colChave] === chaveAcesso) {
      aba.deleteRow(i + 1);
    }
  }

  // Inserir novas
  if (faturas && faturas.length > 0) {
    const cabecalhos = CABECALHOS_NF_FATURAS;
    const linhasParaAdicionar = faturas.map(fatura => {
      return cabecalhos.map(cabecalho => fatura[cabecalho] || '');
    });
    aba.getRange(aba.getLastRow() + 1, 1, linhasParaAdicionar.length, cabecalhos.length).setValues(linhasParaAdicionar);
  }
}

/**
 * Substitui todos os lançamentos de contas a pagar de uma NF por um novo conjunto.
 * Primeiro apaga todas as linhas existentes para a chave e depois insere as novas.
 * @param {string} chaveAcesso A chave de acesso da NF.
 * @param {Array<object>} linhas O novo array de objetos de contas a pagar.
 */
function NotasFiscaisCRUD_substituirContasAPagarDaNF(chaveAcesso, linhas) {
  const planilhaFin = SpreadsheetApp.openById(ID_PLANILHA_FINANCEIRO);
  const aba = planilhaFin.getSheetByName(ABA_FINANCEIRO_CONTAS_A_PAGAR);
  const cabecalhos = CABECALHOS_FINANCEIRO_CONTAS_A_PAGAR;
  const colChave = cabecalhos.indexOf("Chave de Acesso");
  
  const dadosAtuais = aba.getDataRange().getValues();

  // 1. Excluir antigas: itera de baixo para cima para evitar problemas com a remoção de linhas.
  for (let i = dadosAtuais.length - 1; i >= 1; i--) {
    if (dadosAtuais[i][colChave] === chaveAcesso) {
      aba.deleteRow(i + 1);
    }
  }

  // 2. Inserir novas, se houver.
  if (linhas && linhas.length > 0) {
    const linhasParaAdicionar = linhas.map(linha => {
      // Mapeia o objeto de entrada para um array na ordem correta dos cabeçalhos
      return cabecalhos.map(cabecalho => {
          let valor = linha[cabecalho];
          // Normaliza os tipos de dados
          if ((cabecalho === 'Data de Vencimento') && valor) {
              return new Date(valor); // Garante que a data seja salva como objeto Date
          }
          if ((cabecalho === 'Valor da Parcela' || cabecalho === 'Valor por Setor') && (typeof valor === 'string')) {
              return parseFloat(valor.replace(',', '.')) || 0; // Converte string numérica para número
          }
          return valor || ''; // Retorna o valor ou string vazia se for nulo/undefined
      });
    });

    // Insere todas as novas linhas em uma única operação.
    if(linhasParaAdicionar.length > 0) {
      aba.getRange(aba.getLastRow() + 1, 1, linhasParaAdicionar.length, cabecalhos.length).setValues(linhasParaAdicionar);
    }
  }
}
