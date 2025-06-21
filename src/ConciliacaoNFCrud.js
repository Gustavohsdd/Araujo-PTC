/**
 * @file ConciliacaoNFCrud.gs
 * @description Funções CRUD (Create, Read, Update, Delete) para manipular os dados
 * das notas fiscais na planilha e fazer o parsing dos arquivos XML.
 */

// Namespace para o XML da NF-e, essencial para o XmlService funcionar corretamente.
const NFE_NAMESPACE = XmlService.getNamespace('http://www.portalfiscal.inf.br/nfe');

function _ConciliacaoNFCrud_obterValor(elementoPai, tags) {
  let elementoAtual = elementoPai;
  for (const tag of tags) {
    if (!elementoAtual) return null;
    elementoAtual = elementoAtual.getChild(tag, NFE_NAMESPACE);
  }
  return elementoAtual ? elementoAtual.getText() : null;
}

function _ConciliacaoNFCrud_obterValorComoNumero(elementoPai, tags) {
    const valorString = _ConciliacaoNFCrud_obterValor(elementoPai, tags);
    if (valorString === null || valorString.trim() === '') return null;
    const numero = parseFloat(valorString.trim());
    return isNaN(numero) ? null : numero;
}

function _ConciliacaoNFCrud_aplicarFormatacaoNumerica(aba, linhaInicial, numLinhas) {
    if (numLinhas === 0) return;

    // Formatos para colunas numéricas para garantir a correta interpretação pelo Sheets
    const formatos = {
        'Quantidade Comercial': '#,##0.0000', 'Valor Unitário Comercial': '#,##0.00',
        'Valor Total Bruto Item': '#,##0.00', 'Valor do Frete (Item)': '#,##0.00',
        'Valor do Seguro (Item)': '#,##0.00', 'Valor do Desconto (Item)': '#,##0.00',
        'Outras Despesas (Item)': '#,##0.00', 'Base de Cálculo (ICMS)': '#,##0.00',
        'Alíquota (ICMS)': '#,##0.00', 'Valor (ICMS)': '#,##0.00',
        'Valor (ICMS ST)': '#,##0.00', 'Base de Cálculo (IPI)': '#,##0.00',
        'Alíquota (IPI)': '#,##0.00', 'Valor (IPI)': '#,##0.00',
        'Valor (PIS)': '#,##0.00', 'Valor (COFINS)': '#,##0.00',
        'Valor da Parcela': '#,##0.00', 'Quantidade Volumes': '#,##0.000',
        'Peso Líquido Total': '#,##0.000', 'Peso Bruto Total': '#,##0.000',
        'Total Base Cálculo ICMS': '#,##0.00', 'Total Valor ICMS': '#,##0.00',
        'Total Valor ICMS ST': '#,##0.00', 'Total Valor Produtos': '#,##0.00',
        'Total Valor Frete': '#,##0.00', 'Total Valor Seguro': '#,##0.00',
        'Total Valor Desconto': '#,##0.00', 'Total Valor IPI': '#,##0.00',
        'Total Valor PIS': '#,##0.00', 'Total Valor COFINS': '#,##0.00',
        'Total Outras Despesas': '#,##0.00', 'Valor Total da NF': '#,##0.00'
    };

    const cabecalhos = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];

    for (const colunaNome in formatos) {
        const indiceColuna = cabecalhos.indexOf(colunaNome);
        if (indiceColuna !== -1) {
            aba.getRange(linhaInicial, indiceColuna + 1, numLinhas, 1)
               .setNumberFormat(formatos[colunaNome]);
        }
    }
}

function ConciliacaoNFCrud_garantirPastaProcessados() {
  try {
    const pastaPrincipal = DriveApp.getFolderById(ID_PASTA_XML);
    const subpastas = pastaPrincipal.getFoldersByName('Processados');

    if (subpastas.hasNext()) {
      return subpastas.next();
    } else {
      return pastaPrincipal.createFolder('Processados');
    }
  } catch (e) {
    Logger.log(`Erro ao garantir a existência da pasta "Processados": ${e.message}`);
    return null;
  }
}

function ConciliacaoNFCrud_obterChavesDeAcessoExistentes() {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_NF);
    const aba = planilha.getSheetByName(ABA_NF_NOTAS_FISCAIS);
    const ultimaLinha = aba.getLastRow();

    if (ultimaLinha < 2) {
      return new Set();
    }
    const range = aba.getRange(2, 1, ultimaLinha - 1, 1);
    const valores = range.getValues().flat().filter(String);
    
    return new Set(valores);

  } catch (e) {
    Logger.log(`Erro ao obter chaves de acesso existentes: ${e.message}`);
    return new Set();
  }
}

function ConciliacaoNFCrud_parsearConteudoXml(conteudoXml) {
  // Esta função permanece exatamente a mesma da versão anterior.
  // Ela já é otimizada para ler um único XML.
  const documento = XmlService.parse(conteudoXml);
  const root = documento.getRootElement();

  const nfeElement = root.getChild('NFe', NFE_NAMESPACE);
  const protNFe = root.getChild('protNFe', NFE_NAMESPACE);
  
  if (!nfeElement || !protNFe) {
      throw new Error('Estrutura do XML inválida. Tags <NFe> ou <protNFe> não encontradas na raiz do documento.');
  }
  const infNFe = nfeElement.getChild('infNFe', NFE_NAMESPACE);

  if (!infNFe) {
    throw new Error('Estrutura do XML inválida. Tag <infNFe> não encontrada.');
  }

  const ide = infNFe.getChild('ide', NFE_NAMESPACE);
  const emit = infNFe.getChild('emit', NFE_NAMESPACE);
  const dest = infNFe.getChild('dest', NFE_NAMESPACE);
  const total = infNFe.getChild('total', NFE_NAMESPACE);
  const transp = infNFe.getChild('transp', NFE_NAMESPACE);
  const cobr = infNFe.getChild('cobr', NFE_NAMESPACE);
  const infAdic = infNFe.getChild('infAdic', NFE_NAMESPACE);
  
  const chaveAcesso = _ConciliacaoNFCrud_obterValor(protNFe, ['infProt', 'chNFe']);
  
  const dadosNotasFiscais = {
    chaveAcesso: chaveAcesso, numeroNf: _ConciliacaoNFCrud_obterValor(ide, ['nNF']), serieNf: _ConciliacaoNFCrud_obterValor(ide, ['serie']), dataHoraEmissao: _ConciliacaoNFCrud_obterValor(ide, ['dhEmi']), naturezaOperacao: _ConciliacaoNFCrud_obterValor(ide, ['natOp']), cnpjEmitente: _ConciliacaoNFCrud_obterValor(emit, ['CNPJ']), nomeEmitente: _ConciliacaoNFCrud_obterValor(emit, ['xNome']), ieEmitente: _ConciliacaoNFCrud_obterValor(emit, ['IE']), logradouroEmitente: _ConciliacaoNFCrud_obterValor(emit, ['enderEmit', 'xLgr']), numEndEmitente: _ConciliacaoNFCrud_obterValor(emit, ['enderEmit', 'nro']), bairroEmitente: _ConciliacaoNFCrud_obterValor(emit, ['enderEmit', 'xBairro']), municipioEmitente: _ConciliacaoNFCrud_obterValor(emit, ['enderEmit', 'xMun']), ufEmitente: _ConciliacaoNFCrud_obterValor(emit, ['enderEmit', 'UF']), cepEmitente: _ConciliacaoNFCrud_obterValor(emit, ['enderEmit', 'CEP']), cnpjDestinatario: _ConciliacaoNFCrud_obterValor(dest, ['CNPJ']), nomeDestinatario: _ConciliacaoNFCrud_obterValor(dest, ['xNome']), infoAdicionais: _ConciliacaoNFCrud_obterValor(infAdic, ['infCpl'])
  };

  const dadosItensNf = [];
  const dets = infNFe.getChildren('det', NFE_NAMESPACE);
  dets.forEach(det => {
    const prod = det.getChild('prod', NFE_NAMESPACE);
    const imposto = det.getChild('imposto', NFE_NAMESPACE);
    const icms = imposto ? imposto.getChild('ICMS', NFE_NAMESPACE) : null;
    const ipi = imposto ? imposto.getChild('IPI', NFE_NAMESPACE) : null;
    const pis = imposto ? imposto.getChild('PIS', NFE_NAMESPACE) : null;
    const cofins = imposto ? imposto.getChild('COFINS', NFE_NAMESPACE) : null;
    const icmsContent = icms ? icms.getChildren()[0] : null;
    let pisContent = pis ? pis.getChildren()[0] : null;
    let cofinsContent = cofins ? cofins.getChildren()[0] : null;

    dadosItensNf.push({
      chaveAcesso: chaveAcesso, numeroItem: det.getAttribute('nItem').getValue(), codigoProdutoForn: _ConciliacaoNFCrud_obterValor(prod, ['cProd']), gtin: _ConciliacaoNFCrud_obterValor(prod, ['cEAN']), descricaoProduto: _ConciliacaoNFCrud_obterValor(prod, ['xProd']), ncm: _ConciliacaoNFCrud_obterValor(prod, ['NCM']), cfop: _ConciliacaoNFCrud_obterValor(prod, ['CFOP']), unidadeComercial: _ConciliacaoNFCrud_obterValor(prod, ['uCom']), quantidadeComercial: _ConciliacaoNFCrud_obterValorComoNumero(prod, ['qCom']), valorUnitarioComercial: _ConciliacaoNFCrud_obterValorComoNumero(prod, ['vUnCom']), valorTotalBrutoItem: _ConciliacaoNFCrud_obterValorComoNumero(prod, ['vProd']), valorFreteItem: _ConciliacaoNFCrud_obterValorComoNumero(prod, ['vFrete']), valorSeguroItem: _ConciliacaoNFCrud_obterValorComoNumero(prod, ['vSeg']), valorDescontoItem: _ConciliacaoNFCrud_obterValorComoNumero(prod, ['vDesc']), valorOutrasDespesasItem: _ConciliacaoNFCrud_obterValorComoNumero(prod, ['vOutro']), cstCsosnIcms: _ConciliacaoNFCrud_obterValor(icmsContent, ['CST']) || _ConciliacaoNFCrud_obterValor(icmsContent, ['CSOSN']), baseCalculoIcms: _ConciliacaoNFCrud_obterValorComoNumero(icmsContent, ['vBC']), aliquotaIcms: _ConciliacaoNFCrud_obterValorComoNumero(icmsContent, ['pICMS']), valorIcms: _ConciliacaoNFCrud_obterValorComoNumero(icmsContent, ['vICMS']), valorIcmsSt: _ConciliacaoNFCrud_obterValorComoNumero(icmsContent, ['vICMSST']), cstIpi: _ConciliacaoNFCrud_obterValor(ipi, ['IPITrib', 'CST']), baseCalculoIpi: _ConciliacaoNFCrud_obterValorComoNumero(ipi, ['IPITrib', 'vBC']), aliquotaIpi: _ConciliacaoNFCrud_obterValorComoNumero(ipi, ['IPITrib', 'pIPI']), valorIpi: _ConciliacaoNFCrud_obterValorComoNumero(ipi, ['IPITrib', 'vIPI']), cstPis: _ConciliacaoNFCrud_obterValor(pisContent, ['CST']), valorPis: _ConciliacaoNFCrud_obterValorComoNumero(pisContent, ['vPIS']), cstCofins: _ConciliacaoNFCrud_obterValor(cofinsContent, ['CST']), valorCofins: _ConciliacaoNFCrud_obterValorComoNumero(cofinsContent, ['vCOFINS'])
    });
  });

  const dadosFaturasNf = [];
  if (cobr) {
    const dups = cobr.getChildren('dup', NFE_NAMESPACE);
    dups.forEach(dup => {
      dadosFaturasNf.push({
        chaveAcesso: chaveAcesso, numeroFatura: _ConciliacaoNFCrud_obterValor(cobr, ['fat', 'nFat']), numeroParcela: _ConciliacaoNFCrud_obterValor(dup, ['nDup']), dataVencimento: _ConciliacaoNFCrud_obterValor(dup, ['dVenc']), valorParcela: _ConciliacaoNFCrud_obterValorComoNumero(dup, ['vDup']),
      });
    });
  }

  const dadosTransporteNf = {
    chaveAcesso: chaveAcesso, modalidadeFrete: _ConciliacaoNFCrud_obterValor(transp, ['modFrete']), cnpjTransportadora: _ConciliacaoNFCrud_obterValor(transp, ['transporta', 'CNPJ']), nomeTransportadora: _ConciliacaoNFCrud_obterValor(transp, ['transporta', 'xNome']), ieTransportadora: _ConciliacaoNFCrud_obterValor(transp, ['transporta', 'IE']), enderecoTransportadora: _ConciliacaoNFCrud_obterValor(transp, ['transporta', 'xEnder']), placaVeiculo: _ConciliacaoNFCrud_obterValor(transp, ['veicTransp', 'placa']), quantidadeVolumes: _ConciliacaoNFCrud_obterValorComoNumero(transp, ['vol', 'qVol']), especieVolumes: _ConciliacaoNFCrud_obterValor(transp, ['vol', 'esp']), pesoLiquidoTotal: _ConciliacaoNFCrud_obterValorComoNumero(transp, ['vol', 'pesoL']), pesoBrutoTotal: _ConciliacaoNFCrud_obterValorComoNumero(transp, ['vol', 'pesoB']),
  };
  
  const icmsTot = total ? total.getChild('ICMSTot', NFE_NAMESPACE) : null;
  const dadosTributosTotais = {
    chaveAcesso: chaveAcesso, totalBaseCalculoIcms: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vBC']), totalValorIcms: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vICMS']), totalValorIcmsSt: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vST']), totalValorProdutos: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vProd']), totalValorFrete: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vFrete']), totalValorSeguro: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vSeg']), totalValorDesconto: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vDesc']), totalValorIpi: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vIPI']), totalValorPis: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vPIS']), totalValorCofins: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vCOFINS']), totalOutrasDespesas: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vOutro']), valorTotalNf: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vNF']),
  };

  return {
    notasFiscais: dadosNotasFiscais, itensNf: dadosItensNf, faturasNf: dadosFaturasNf, transporteNf: [dadosTransporteNf], tributosTotaisNf: [dadosTributosTotais],
  };
}

/**
 * NOVA FUNÇÃO OTIMIZADA: Salva todos os dados acumulados de várias NF-e de uma só vez.
 * @param {object} todosOsDados Um objeto contendo arrays de dados para cada aba.
 */
function ConciliacaoNFCrud_salvarDadosEmLote(todosOsDados) {
  const planilha = SpreadsheetApp.openById(ID_PLANILHA_NF);

  const mapeamento = [
    { aba: ABA_NF_NOTAS_FISCAIS, cabecalhos: CABECALHOS_NF_NOTAS_FISCAIS, dados: todosOsDados.notasFiscais },
    { aba: ABA_NF_ITENS, cabecalhos: CABECALHOS_NF_ITENS, dados: todosOsDados.itensNf },
    { aba: ABA_NF_FATURAS, cabecalhos: CABECALHOS_NF_FATURAS, dados: todosOsDados.faturasNf },
    { aba: ABA_NF_TRANSPORTE, cabecalhos: CABECALHOS_NF_TRANSPORTE, dados: todosOsDados.transporteNf },
    { aba: ABA_NF_TRIBUTOS_TOTAIS, cabecalhos: CABECALHOS_NF_TRIBUTOS_TOTAIS, dados: todosOsDados.tributosTotaisNf }
  ];

  const mapaChaves = {
    'NotasFiscais': { 'Chave de Acesso': 'chaveAcesso', 'ID da Cotação (Sistema)': '', 'Status da Conciliação': '', 'Número NF': 'numeroNf', 'Série NF': 'serieNf', 'Data e Hora Emissão': 'dataHoraEmissao', 'Natureza da Operação': 'naturezaOperacao', 'CNPJ Emitente': 'cnpjEmitente', 'Nome Emitente': 'nomeEmitente', 'Inscrição Estadual Emitente': 'ieEmitente', 'Logradouro Emitente': 'logradouroEmitente', 'Número End. Emitente': 'numEndEmitente', 'Bairro Emitente': 'bairroEmitente', 'Município Emitente': 'municipioEmitente', 'UF Emitente': 'ufEmitente', 'CEP Emitente': 'cepEmitente', 'CNPJ Destinatário': 'cnpjDestinatario', 'Nome Destinatário': 'nomeDestinatario', 'Informações Adicionais': 'infoAdicionais', 'Número do Pedido (Extraído)': '' },
    'ItensNF': { 'Chave de Acesso': 'chaveAcesso', 'Número do Item': 'numeroItem', 'Código Produto (Forn)': 'codigoProdutoForn', 'GTIN/EAN (Cód. Barras)': 'gtin', 'Descrição Produto (NF)': 'descricaoProduto', 'NCM': 'ncm', 'CFOP': 'cfop', 'Unidade Comercial': 'unidadeComercial', 'Quantidade Comercial': 'quantidadeComercial', 'Valor Unitário Comercial': 'valorUnitarioComercial', 'Valor Total Bruto Item': 'valorTotalBrutoItem', 'Valor do Frete (Item)': 'valorFreteItem', 'Valor do Seguro (Item)': 'valorSeguroItem', 'Valor do Desconto (Item)': 'valorDescontoItem', 'Outras Despesas (Item)': 'valorOutrasDespesasItem', 'CST/CSOSN (ICMS)': 'cstCsosnIcms', 'Base de Cálculo (ICMS)': 'baseCalculoIcms', 'Alíquota (ICMS)': 'aliquotaIcms', 'Valor (ICMS)': 'valorIcms', 'Valor (ICMS ST)': 'valorIcmsSt', 'CST (IPI)': 'cstIpi', 'Base de Cálculo (IPI)': 'baseCalculoIpi', 'Alíquota (IPI)': 'aliquotaIpi', 'Valor (IPI)': 'valorIpi', 'CST (PIS)': 'cstPis', 'Valor (PIS)': 'valorPis', 'CST (COFINS)': 'cstCofins', 'Valor (COFINS)': 'valorCofins' },
    'FaturasNF': { 'Chave de Acesso': 'chaveAcesso', 'Número da Fatura': 'numeroFatura', 'Número da Parcela': 'numeroParcela', 'Data de Vencimento': 'dataVencimento', 'Valor da Parcela': 'valorParcela' },
    'TransporteNF': { 'Chave de Acesso': 'chaveAcesso', 'Modalidade Frete': 'modalidadeFrete', 'CNPJ Transportadora': 'cnpjTransportadora', 'Nome Transportadora': 'nomeTransportadora', 'IE Transportadora': 'ieTransportadora', 'Endereço Transportadora': 'enderecoTransportadora', 'Placa Veículo': 'placaVeiculo', 'Quantidade Volumes': 'quantidadeVolumes', 'Espécie Volumes': 'especieVolumes', 'Peso Líquido Total': 'pesoLiquidoTotal', 'Peso Bruto Total': 'pesoBrutoTotal' },
    'TributosTotaisNF': { 'Chave de Acesso': 'chaveAcesso', 'Total Base Cálculo ICMS': 'totalBaseCalculoIcms', 'Total Valor ICMS': 'totalValorIcms', 'Total Valor ICMS ST': 'totalValorIcmsSt', 'Total Valor Produtos': 'totalValorProdutos', 'Total Valor Frete': 'totalValorFrete', 'Total Valor Seguro': 'totalValorSeguro', 'Total Valor Desconto': 'totalValorDesconto', 'Total Valor IPI': 'totalValorIpi', 'Total Valor PIS': 'totalValorPis', 'Total Valor COFINS': 'totalValorCofins', 'Total Outras Despesas': 'totalOutrasDespesas', 'Valor Total da NF': 'valorTotalNf' }
  };

  for (const item of mapeamento) {
    if (item.dados && item.dados.length > 0) {
      const aba = planilha.getSheetByName(item.aba);
      const linhaInicial = aba.getLastRow() + 1;
      const numLinhasAdicionadas = item.dados.length;

      const linhasParaAdicionar = item.dados.map(objetoDado => {
        return item.cabecalhos.map(cabecalho => {
          const chave = mapaChaves[item.aba][cabecalho];
          if (chave === '') return ''; 
          return objetoDado[chave] !== undefined && objetoDado[chave] !== null ? objetoDado[chave] : '';
        });
      });

      aba.getRange(linhaInicial, 1, numLinhasAdicionadas, item.cabecalhos.length)
         .setValues(linhasParaAdicionar);
      
      _ConciliacaoNFCrud_aplicarFormatacaoNumerica(aba, linhaInicial, numLinhasAdicionadas);
    }
  }
}

/**
 * @file ConciliacaoNFCrud.gs
 * @description Funções CRUD Corrigidas para o processo de conciliação.
 */

/**
 * [CORRIGIDO] Obtém uma lista de cotações com status "Aguardando Faturamento".
 * Trata a combinação de ID e Fornecedor como única e usa o CNPJ para futuras comparações.
 * @returns {Array<object>|null} Array de objetos de cotação ou null em caso de erro.
 */
function ConciliacaoNFCrud_obterCotacoesAbertas() {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA_COTACOES);
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha <= 1) return [];

    const cabecalhos = Utilities_obterCabecalhos(ABA_COTACOES);
    const colId = cabecalhos.indexOf("ID da Cotação");
    const colData = cabecalhos.indexOf("Data Abertura");
    const colForn = cabecalhos.indexOf("Fornecedor");
    const colStatus = cabecalhos.indexOf("Status da Cotação");

    const dados = aba.getRange(2, 1, ultimaLinha - 1, aba.getLastColumn()).getValues();
    
    // Mapeia Fornecedor para CNPJ para uso posterior
    const fornecedoresData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_FORNECEDORES).getDataRange().getValues();
    const cabecalhosForn = fornecedoresData.shift();
    const colFornNome = cabecalhosForn.indexOf("Fornecedor");
    const colFornCnpj = cabecalhosForn.indexOf("CNPJ");
    const mapaFornecedores = fornecedoresData.reduce((map, row) => {
        const nome = row[colFornNome];
        const cnpj = row[colFornCnpj];
        if (nome) {
            map[nome.toString().trim()] = cnpj ? cnpj.toString().trim() : '';
        }
        return map;
    }, {});

    const cotacoesUnicas = {};
    dados.forEach(linha => {
      const status = linha[colStatus];
      const id = linha[colId];
      const fornecedor = linha[colForn];

      // [NOVA REGRA] A cotação só é válida se tiver ID, Fornecedor, e o status for "Aguardando Faturamento".
      if (id && fornecedor && status === 'Aguardando Faturamento') {
        const compositeKey = `${id}-${fornecedor}`; // Cria chave composta (ex: "54-Ambev")
        
        if (!cotacoesUnicas[compositeKey]) { // Verifica a unicidade usando a chave composta
          const nomeFornecedorTrim = fornecedor.toString().trim();
          cotacoesUnicas[compositeKey] = {
            compositeKey: compositeKey, // Chave para o valor do <option> no HTML
            idCotacao: id,
            fornecedor: fornecedor,
            fornecedorCnpj: mapaFornecedores[nomeFornecedorTrim] || '', // Garante que o CNPJ é encontrado
            dataAbertura: new Date(linha[colData]).toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' })
          };
        }
      }
    });
    
    const resultado = Object.values(cotacoesUnicas);
    
    // Ordena por ID (mais novo primeiro) e depois por nome do fornecedor
    resultado.sort((a, b) => {
        if (b.idCotacao !== a.idCotacao) {
            return b.idCotacao - a.idCotacao;
        }
        return a.fornecedor.localeCompare(b.fornecedor);
    });

    return resultado;

  } catch(e) {
    Logger.log(`Erro em ConciliacaoNFCrud_obterCotacoesAbertas: ${e.message}\n${e.stack}`);
    return null;
  }
}

/**
 * Obtém uma lista de NFs que ainda não foram conciliadas.
 * @returns {Array<object>|null} Array de objetos de NF ou null em caso de erro.
 */
function ConciliacaoNFCrud_obterNFsNaoConciliadas() {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_NF);
    const aba = planilha.getSheetByName(ABA_NF_NOTAS_FISCAIS);
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha <= 1) return [];

    const cabecalhos = CABECALHOS_NF_NOTAS_FISCAIS;
    const colChave = cabecalhos.indexOf("Chave de Acesso");
    const colStatus = cabecalhos.indexOf("Status da Conciliação");
    const colNumNF = cabecalhos.indexOf("Número NF");
    const colEmitente = cabecalhos.indexOf("Nome Emitente");
    const colCnpj = cabecalhos.indexOf("CNPJ Emitente");
    const colData = cabecalhos.indexOf("Data e Hora Emissão");

    const dados = aba.getRange(2, 1, ultimaLinha - 1, aba.getLastColumn()).getValues();
    const nfs = [];
    dados.forEach(linha => {
      if (linha[colStatus] !== 'Conciliada') {
        nfs.push({
          chaveAcesso: linha[colChave],
          numeroNF: linha[colNumNF],
          nomeEmitente: linha[colEmitente],
          cnpjEmitente: linha[colCnpj],
          dataEmissao: new Date(linha[colData]).toLocaleDateString('pt-BR')
        });
      }
    });
    return nfs;
  } catch(e) {
    Logger.log(`Erro em ConciliacaoNFCrud_obterNFsNaoConciliadas: ${e.message}`);
    return null;
  }
}

/**
 * [CORRIGIDO] Obtém os itens de uma cotação específica, filtrando por ID e Fornecedor.
 * @param {string} idCotacao - O ID da cotação.
 * @param {string} nomeFornecedor - O nome do fornecedor.
 * @returns {Array<object>} Array de itens da cotação.
 */
function ConciliacaoNFCrud_obterItensDaCotacao(idCotacao, nomeFornecedor) {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA_COTACOES);
    const cabecalhos = Utilities_obterCabecalhos(ABA_COTACOES);
    const colId = cabecalhos.indexOf("ID da Cotação");
    const colSubProd = cabecalhos.indexOf("SubProduto");
    const colComprar = cabecalhos.indexOf("Comprar");
    const colPreco = cabecalhos.indexOf("Preço");
    const colForn = cabecalhos.indexOf("Fornecedor");
    const colData = cabecalhos.indexOf("Data Abertura");
    
    const dados = aba.getDataRange().getValues();
    const itens = [];
    dados.forEach(linha => {
        const qtdComprar = parseFloat(linha[colComprar]);
        // Adiciona a verificação do fornecedor ao filtro
        if (linha[colId] == idCotacao && linha[colForn] == nomeFornecedor && !isNaN(qtdComprar) && qtdComprar > 0) {
            itens.push({
                subProduto: linha[colSubProd],
                qtdComprar: qtdComprar,
                preco: parseFloat(linha[colPreco]) || 0,
                fornecedor: linha[colForn],
                dataAbertura: linha[colData]
            });
        }
    });
    return itens;
}

/**
 * Obtém os itens das NFs especificadas.
 * @param {Array<string>} chavesAcessoNF - Array de chaves de acesso.
 * @returns {Array<object>} Array de itens de NF.
 */
function ConciliacaoNFCrud_obterItensDasNFs(chavesAcessoNF) {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_NF);
    const aba = planilha.getSheetByName(ABA_NF_ITENS);
    const cabecalhos = CABECALHOS_NF_ITENS;
    const colChave = cabecalhos.indexOf("Chave de Acesso");
    const colDesc = cabecalhos.indexOf("Descrição Produto (NF)");
    const colQtd = cabecalhos.indexOf("Quantidade Comercial");
    const colPreco = cabecalhos.indexOf("Valor Unitário Comercial");
    
    const dados = aba.getDataRange().getValues();
    const itens = [];
    dados.forEach(linha => {
        if (chavesAcessoNF.includes(linha[colChave])) {
            itens.push({
                descricaoNF: linha[colDesc],
                qtdNF: parseFloat(linha[colQtd]) || 0,
                precoNF: parseFloat(linha[colPreco]) || 0
            });
        }
    });
    return itens;
}

/**
 * Obtém os dados gerais agregados das NFs.
 * @param {Array<string>} chavesAcessoNF - Array de chaves de acesso.
 * @returns {object} Objeto com dados gerais.
 */
function ConciliacaoNFCrud_obterDadosGeraisDasNFs(chavesAcessoNF) {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_NF);
    const abaNF = planilha.getSheetByName(ABA_NF_NOTAS_FISCAIS);
    const abaTrib = planilha.getSheetByName(ABA_NF_TRIBUTOS_TOTAIS);
    
    const dadosNF = abaNF.getDataRange().getValues();
    const dadosTrib = abaTrib.getDataRange().getValues();

    let dataEmissao = null;
    let valorTotalNF = 0;

    const cabNF = CABECALHOS_NF_NOTAS_FISCAIS;
    const colChaveNF = cabNF.indexOf("Chave de Acesso");
    const colDataNF = cabNF.indexOf("Data e Hora Emissão");
    
    dadosNF.forEach(linha => {
        if (chavesAcessoNF.includes(linha[colChaveNF])) {
            if (!dataEmissao) dataEmissao = linha[colDataNF]; // Pega a data da primeira NF
        }
    });

    const cabTrib = CABECALHOS_NF_TRIBUTOS_TOTAIS;
    const colChaveTrib = cabTrib.indexOf("Chave de Acesso");
    const colTotalTrib = cabTrib.indexOf("Valor Total da NF");
    
    dadosTrib.forEach(linha => {
        if (chavesAcessoNF.includes(linha[colChaveTrib])) {
            valorTotalNF += parseFloat(linha[colTotalTrib]) || 0;
        }
    });

    return {
        dataEmissao: dataEmissao,
        valorTotalNF: valorTotalNF
    };
}

/**
 * Obtém o prazo de entrega de um fornecedor.
 * @param {string} nomeFornecedor - O nome do fornecedor.
 * @returns {number} O prazo de entrega em dias.
 */
function ConciliacaoNFCrud_obterPrazoFornecedor(nomeFornecedor) {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA_FORNECEDORES);
    const cabecalhos = Utilities_obterCabecalhos(ABA_FORNECEDORES);
    const colNome = cabecalhos.indexOf("Fornecedor");
    const colPrazo = cabecalhos.indexOf("Dias de Entrega");
    
    const dados = aba.getDataRange().getValues();
    for (let i = 0; i < dados.length; i++) {
        if (dados[i][colNome] === nomeFornecedor) {
            return parseInt(dados[i][colPrazo]) || 0;
        }
    }
    return 0; // Padrão
}

/**
 * [CORRIGIDO] Atualiza o status na planilha de Cotações, filtrando por ID e Fornecedor.
 * @param {string} idCotacao - O ID da cotação.
 * @param {string} nomeFornecedor - O nome do fornecedor.
 * @param {Array<object>} itensConciliados - Itens que deram match.
 * @param {Array<object>} itensSomenteCotacao - Itens que foram cortados.
 * @returns {boolean} True se sucesso.
 */
function ConciliacaoNFCrud_atualizarStatusCotacao(idCotacao, nomeFornecedor, itensConciliados, itensSomenteCotacao) {
    try {
        const planilha = SpreadsheetApp.getActiveSpreadsheet();
        const aba = planilha.getSheetByName(ABA_COTACOES);
        const range = aba.getDataRange();
        const dados = range.getValues();
        const cabecalhos = dados[0];

        const colId = cabecalhos.indexOf("ID da Cotação");
        const colForn = cabecalhos.indexOf("Fornecedor");
        const colSubProd = cabecalhos.indexOf("SubProduto");
        const colStatusCot = cabecalhos.indexOf("Status da Cotação");
        const colStatusSub = cabecalhos.indexOf("Status do SubProduto");

        const nomesConciliados = new Set(itensConciliados.map(i => i.subProduto));
        const nomesCortados = new Set(itensSomenteCotacao.map(i => i.subProduto));

        for (let i = 1; i < dados.length; i++) {
            // Adiciona a verificação do fornecedor ao filtro
            if (dados[i][colId] == idCotacao && dados[i][colForn] == nomeFornecedor) {
                const subProdutoLinha = dados[i][colSubProd];
                if (nomesConciliados.has(subProdutoLinha)) {
                    dados[i][colStatusCot] = "Faturado";
                } else if (nomesCortados.has(subProdutoLinha)) {
                    dados[i][colStatusSub] = "Cortado";
                }
            }
        }
        range.setValues(dados);
        return true;
    } catch(e) {
        Logger.log(`Erro em ConciliacaoNFCrud_atualizarStatusCotacao: ${e.message}`);
        return false;
    }
}

/**
 * Atualiza o status na planilha de Notas Fiscais após a conciliação.
 * @param {Array<string>} chavesAcessoNF - As chaves das NFs conciliadas.
 * @param {string} idCotacao - O ID da cotação vinculada.
 * @returns {boolean} True se sucesso.
 */
function ConciliacaoNFCrud_atualizarStatusNF(chavesAcessoNF, idCotacao) {
    try {
        const planilha = SpreadsheetApp.openById(ID_PLANILHA_NF);
        const aba = planilha.getSheetByName(ABA_NF_NOTAS_FISCAIS);
        const range = aba.getDataRange();
        const dados = range.getValues();
        const cabecalhos = dados[0];
        
        const colChave = cabecalhos.indexOf("Chave de Acesso");
        const colStatus = cabecalhos.indexOf("Status da Conciliação");
        const colIdCot = cabecalhos.indexOf("ID da Cotação (Sistema)");

        for (let i = 1; i < dados.length; i++) {
            if (chavesAcessoNF.includes(dados[i][colChave])) {
                dados[i][colStatus] = "Conciliada";
                dados[i][colIdCot] = idCotacao;
            }
        }
        range.setValues(dados);
        return true;
    } catch(e) {
        Logger.log(`Erro em ConciliacaoNFCrud_atualizarStatusNF: ${e.message}`);
        return false;
    }
}
