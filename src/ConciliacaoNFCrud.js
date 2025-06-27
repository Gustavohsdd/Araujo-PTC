/**
 * @file ConciliacaoNFCrud.gs
 * @description Funções CRUD (Create, Read, Update, Delete) para manipular os dados
 * das notas fiscais na planilha e fazer o parsing dos arquivos XML.
 */

// Namespace para o XML da NF-e, essencial para o XmlService funcionar corretamente.
const NFE_NAMESPACE = XmlService.getNamespace('http://www.portalfiscal.inf.br/nfe');

// Funções auxiliares internas
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
    const formatos = {
        'Quantidade Comercial': '#,##0.0000', 'Valor Unitário Comercial': '#,##0.0000000000',
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

// Funções CRUD Principais
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
    if (ultimaLinha < 2) return new Set();
    
    const range = aba.getRange(2, 1, ultimaLinha - 1, 1);
    const valores = range.getValues().flat().filter(String);
    return new Set(valores);
  } catch (e) {
    Logger.log(`Erro ao obter chaves de acesso existentes: ${e.message}`);
    return new Set();
  }
}

function ConciliacaoNFCrud_parsearConteudoXml(conteudoXml) {
  const documento = XmlService.parse(conteudoXml);
  const root = documento.getRootElement();
  const nfeElement = root.getChild('NFe', NFE_NAMESPACE);
  const protNFe = root.getChild('protNFe', NFE_NAMESPACE);
  if (!nfeElement || !protNFe) throw new Error('Estrutura do XML inválida: <NFe> ou <protNFe> não encontradas.');
  
  const infNFe = nfeElement.getChild('infNFe', NFE_NAMESPACE);
  if (!infNFe) throw new Error('Estrutura do XML inválida: <infNFe> não encontrada.');

  const ide = infNFe.getChild('ide', NFE_NAMESPACE);
  const emit = infNFe.getChild('emit', NFE_NAMESPACE);
  const dest = infNFe.getChild('dest', NFE_NAMESPACE);
  const total = infNFe.getChild('total', NFE_NAMESPACE);
  const transp = infNFe.getChild('transp', NFE_NAMESPACE);
  const cobr = infNFe.getChild('cobr', NFE_NAMESPACE);
  const infAdic = infNFe.getChild('infAdic', NFE_NAMESPACE);
  const chaveAcesso = _ConciliacaoNFCrud_obterValor(protNFe, ['infProt', 'chNFe']);
  
  const dadosNotasFiscais = { chaveAcesso: chaveAcesso, numeroNf: _ConciliacaoNFCrud_obterValor(ide, ['nNF']), serieNf: _ConciliacaoNFCrud_obterValor(ide, ['serie']), dataHoraEmissao: _ConciliacaoNFCrud_obterValor(ide, ['dhEmi']), naturezaOperacao: _ConciliacaoNFCrud_obterValor(ide, ['natOp']), cnpjEmitente: _ConciliacaoNFCrud_obterValor(emit, ['CNPJ']), nomeEmitente: _ConciliacaoNFCrud_obterValor(emit, ['xNome']), ieEmitente: _ConciliacaoNFCrud_obterValor(emit, ['IE']), logradouroEmitente: _ConciliacaoNFCrud_obterValor(emit, ['enderEmit', 'xLgr']), numEndEmitente: _ConciliacaoNFCrud_obterValor(emit, ['enderEmit', 'nro']), bairroEmitente: _ConciliacaoNFCrud_obterValor(emit, ['enderEmit', 'xBairro']), municipioEmitente: _ConciliacaoNFCrud_obterValor(emit, ['enderEmit', 'xMun']), ufEmitente: _ConciliacaoNFCrud_obterValor(emit, ['enderEmit', 'UF']), cepEmitente: _ConciliacaoNFCrud_obterValor(emit, ['enderEmit', 'CEP']), cnpjDestinatario: _ConciliacaoNFCrud_obterValor(dest, ['CNPJ']), nomeDestinatario: _ConciliacaoNFCrud_obterValor(dest, ['xNome']), infoAdicionais: _ConciliacaoNFCrud_obterValor(infAdic, ['infCpl']) };

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
      dadosFaturasNf.push({ chaveAcesso: chaveAcesso, numeroFatura: _ConciliacaoNFCrud_obterValor(cobr, ['fat', 'nFat']), numeroParcela: _ConciliacaoNFCrud_obterValor(dup, ['nDup']), dataVencimento: _ConciliacaoNFCrud_obterValor(dup, ['dVenc']), valorParcela: _ConciliacaoNFCrud_obterValorComoNumero(dup, ['vDup']), });
    });
  }

  const dadosTransporteNf = { chaveAcesso: chaveAcesso, modalidadeFrete: _ConciliacaoNFCrud_obterValor(transp, ['modFrete']), cnpjTransportadora: _ConciliacaoNFCrud_obterValor(transp, ['transporta', 'CNPJ']), nomeTransportadora: _ConciliacaoNFCrud_obterValor(transp, ['transporta', 'xNome']), ieTransportadora: _ConciliacaoNFCrud_obterValor(transp, ['transporta', 'IE']), enderecoTransportadora: _ConciliacaoNFCrud_obterValor(transp, ['transporta', 'xEnder']), placaVeiculo: _ConciliacaoNFCrud_obterValor(transp, ['veicTransp', 'placa']), quantidadeVolumes: _ConciliacaoNFCrud_obterValorComoNumero(transp, ['vol', 'qVol']), especieVolumes: _ConciliacaoNFCrud_obterValor(transp, ['vol', 'esp']), pesoLiquidoTotal: _ConciliacaoNFCrud_obterValorComoNumero(transp, ['vol', 'pesoL']), pesoBrutoTotal: _ConciliacaoNFCrud_obterValorComoNumero(transp, ['vol', 'pesoB']), };
  
  const icmsTot = total ? total.getChild('ICMSTot', NFE_NAMESPACE) : null;
  const dadosTributosTotais = { chaveAcesso: chaveAcesso, totalBaseCalculoIcms: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vBC']), totalValorIcms: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vICMS']), totalValorIcmsSt: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vST']), totalValorProdutos: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vProd']), totalValorFrete: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vFrete']), totalValorSeguro: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vSeg']), totalValorDesconto: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vDesc']), totalValorIpi: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vIPI']), totalValorPis: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vPIS']), totalValorCofins: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vCOFINS']), totalOutrasDespesas: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vOutro']), valorTotalNf: _ConciliacaoNFCrud_obterValorComoNumero(icmsTot, ['vNF']), };

  return { notasFiscais: dadosNotasFiscais, itensNf: dadosItensNf, faturasNf: dadosFaturasNf, transporteNf: [dadosTransporteNf], tributosTotaisNf: [dadosTributosTotais], };
}

function ConciliacaoNFCrud_salvarDadosEmLote(todosOsDados) {
  const planilha = SpreadsheetApp.openById(ID_PLANILHA_NF);
  const mapeamento = [
    { aba: ABA_NF_NOTAS_FISCAIS, cabecalhos: CABECALHOS_NF_NOTAS_FISCAIS, dados: todosOsDados.notasFiscais }, { aba: ABA_NF_ITENS, cabecalhos: CABECALHOS_NF_ITENS, dados: todosOsDados.itensNf }, { aba: ABA_NF_FATURAS, cabecalhos: CABECALHOS_NF_FATURAS, dados: todosOsDados.faturasNf }, { aba: ABA_NF_TRANSPORTE, cabecalhos: CABECALHOS_NF_TRANSPORTE, dados: todosOsDados.transporteNf }, { aba: ABA_NF_TRIBUTOS_TOTAIS, cabecalhos: CABECALHOS_NF_TRIBUTOS_TOTAIS, dados: todosOsDados.tributosTotaisNf }
  ];
  const mapaChaves = {
    'NotasFiscais': { 'Chave de Acesso': 'chaveAcesso', 'ID da Cotação (Sistema)': '', 'Status da Conciliação': 'Pendente', 'Número NF': 'numeroNf', 'Série NF': 'serieNf', 'Data e Hora Emissão': 'dataHoraEmissao', 'Natureza da Operação': 'naturezaOperacao', 'CNPJ Emitente': 'cnpjEmitente', 'Nome Emitente': 'nomeEmitente', 'Inscrição Estadual Emitente': 'ieEmitente', 'Logradouro Emitente': 'logradouroEmitente', 'Número End. Emitente': 'numEndEmitente', 'Bairro Emitente': 'bairroEmitente', 'Município Emitente': 'municipioEmitente', 'UF Emitente': 'ufEmitente', 'CEP Emitente': 'cepEmitente', 'CNPJ Destinatário': 'cnpjDestinatario', 'Nome Destinatário': 'nomeDestinatario', 'Informações Adicionais': 'infoAdicionais', 'Número do Pedido (Extraído)': '' },
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
      aba.getRange(linhaInicial, 1, numLinhasAdicionadas, item.cabecalhos.length).setValues(linhasParaAdicionar);
      _ConciliacaoNFCrud_aplicarFormatacaoNumerica(aba, linhaInicial, numLinhasAdicionadas);
    }
  }
}

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
    
    const fornecedoresData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_FORNECEDORES).getDataRange().getValues();
    const cabecalhosForn = fornecedoresData.shift();
    const colFornNome = cabecalhosForn.indexOf("Fornecedor");
    const colFornCnpj = cabecalhosForn.indexOf("CNPJ");
    const mapaFornecedores = fornecedoresData.reduce((map, row) => {
        const nome = row[colFornNome];
        const cnpj = row[colFornCnpj];
        if (nome) map[nome.toString().trim()] = cnpj ? cnpj.toString().trim() : '';
        return map;
    }, {});

    const cotacoesUnicas = {};
    dados.forEach(linha => {
      const status = linha[colStatus];
      const id = linha[colId];
      const fornecedor = linha[colForn];
      if (id && fornecedor && status === 'Aguardando Faturamento') {
        const compositeKey = `${id}-${fornecedor}`; 
        if (!cotacoesUnicas[compositeKey]) { 
          const nomeFornecedorTrim = fornecedor.toString().trim();
          cotacoesUnicas[compositeKey] = {
            compositeKey: compositeKey, idCotacao: id, fornecedor: fornecedor,
            fornecedorCnpj: mapaFornecedores[nomeFornecedorTrim] || '',
            dataAbertura: new Date(linha[colData]).toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' })
          };
        }
      }
    });
    
    const resultado = Object.values(cotacoesUnicas);
    resultado.sort((a, b) => {
        if (b.idCotacao !== a.idCotacao) return b.idCotacao - a.idCotacao;
        return a.fornecedor.localeCompare(b.fornecedor);
    });
    return resultado;
  } catch(e) {
    Logger.log(`Erro em ConciliacaoNFCrud_obterCotacoesAbertas: ${e.message}\n${e.stack}`);
    return null;
  }
}

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
      const statusAtual = linha[colStatus];
      if (statusAtual !== 'Conciliada' && statusAtual !== 'Sem Pedido') {
        nfs.push({
          chaveAcesso: linha[colChave], numeroNF: linha[colNumNF], nomeEmitente: linha[colEmitente],
          cnpjEmitente: linha[colCnpj], dataEmissao: new Date(linha[colData]).toLocaleDateString('pt-BR')
        });
      }
    });
    return nfs;
  } catch(e) {
    Logger.log(`Erro em ConciliacaoNFCrud_obterNFsNaoConciliadas: ${e.message}`);
    return null;
  }
}

function ConciliacaoNFCrud_obterItensDaCotacao(idCotacao, nomeFornecedor) {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA_COTACOES);
    const cabecalhos = Utilities_obterCabecalhos(ABA_COTACOES);
    if (!cabecalhos) throw new Error(`Não foi possível obter os cabeçalhos da aba: ${ABA_COTACOES}`);

    const colMap = {};
    const colunasNecessarias = ["ID da Cotação", "SubProduto", "Comprar", "Preço", "Fornecedor", "Data Abertura", "Fator", "Preço por Fator"];
    colunasNecessarias.forEach(nomeColuna => {
        const index = cabecalhos.indexOf(nomeColuna);
        if (index === -1) throw new Error(`A coluna obrigatória "${nomeColuna}" não foi encontrada na aba Cotações.`);
        colMap[nomeColuna] = index;
    });
    
    const dados = aba.getDataRange().getValues();
    const itens = [];
    for (let i = 1; i < dados.length; i++) {
        const linha = dados[i];
        const qtdComprar = parseFloat(linha[colMap["Comprar"]]);
        if (linha[colMap["ID da Cotação"]] == idCotacao && linha[colMap["Fornecedor"]] == nomeFornecedor && !isNaN(qtdComprar) && qtdComprar > 0) {
            const dataAberturaObj = new Date(linha[colMap["Data Abertura"]]);
            itens.push({
                subProduto: linha[colMap["SubProduto"]], qtdComprar: qtdComprar, preco: parseFloat(linha[colMap["Preço"]]) || 0,
                fator: parseFloat(linha[colMap["Fator"]]) || 1, precoPorFator: parseFloat(linha[colMap["Preço por Fator"]]) || 0,
                fornecedor: linha[colMap["Fornecedor"]], dataAbertura: !isNaN(dataAberturaObj.getTime()) ? dataAberturaObj.toISOString() : null
            });
        }
    }
    return itens;
}

function ConciliacaoNFCrud_obterItensDasNFs(chavesAcessoNF) {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_NF);
    const aba = planilha.getSheetByName(ABA_NF_ITENS);
    const cabecalhos = CABECALHOS_NF_ITENS;
    const colChave = cabecalhos.indexOf("Chave de Acesso");
    const colDesc = cabecalhos.indexOf("Descrição Produto (NF)");
    const colQtd = cabecalhos.indexOf("Quantidade Comercial");
    const colPreco = cabecalhos.indexOf("Valor Unitário Comercial");
    const colNumItem = cabecalhos.indexOf("Número do Item");
    
    const dados = aba.getDataRange().getValues();
    const itens = [];
    dados.forEach(linha => {
        const chaveLinha = linha[colChave];
        if (chavesAcessoNF.includes(chaveLinha)) {
            itens.push({
                chaveAcesso: chaveLinha, numeroItem: linha[colNumItem], descricaoNF: linha[colDesc],
                qtdNF: parseFloat(linha[colQtd]) || 0, precoNF: parseFloat(linha[colPreco]) || 0
            });
        }
    });
    return itens;
}

function ConciliacaoNFCrud_obterDadosGeraisDasNFs(chavesAcessoNF) {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_NF);
    const abaTrib = planilha.getSheetByName(ABA_NF_TRIBUTOS_TOTAIS);
    const dadosCompletos = abaTrib.getDataRange().getValues();
    const cabecalhos = dadosCompletos.shift();
    const chavesSet = new Set(chavesAcessoNF);
    const resultados = [];
    const colMap = {};
    cabecalhos.forEach((header, index) => { colMap[header] = index; });

    for (const linha of dadosCompletos) {
        const chaveAtual = linha[colMap["Chave de Acesso"]];
        if (chavesSet.has(chaveAtual)) {
            resultados.push({
              chaveAcesso: chaveAtual,
              totalBaseCalculoIcms: parseFloat(linha[colMap["Total Base Cálculo ICMS"]]) || 0,
              totalValorIcms: parseFloat(linha[colMap["Total Valor ICMS"]]) || 0, totalValorIcmsSt: parseFloat(linha[colMap["Total Valor ICMS ST"]]) || 0,
              totalValorProdutos: parseFloat(linha[colMap["Total Valor Produtos"]]) || 0, totalValorFrete: parseFloat(linha[colMap["Total Valor Frete"]]) || 0,
              totalValorSeguro: parseFloat(linha[colMap["Total Valor Seguro"]]) || 0, totalValorDesconto: parseFloat(linha[colMap["Total Valor Desconto"]]) || 0,
              totalValorIpi: parseFloat(linha[colMap["Total Valor IPI"]]) || 0, totalValorPis: parseFloat(linha[colMap["Total Valor PIS"]]) || 0,
              totalValorCofins: parseFloat(linha[colMap["Total Valor COFINS"]]) || 0, totalOutrasDespesas: parseFloat(linha[colMap["Total Outras Despesas"]]) || 0,
              valorTotalNf: parseFloat(linha[colMap["Valor Total da NF"]]) || 0,
            });
        }
    }
    return resultados;
}

/**
 * [NOVA FUNÇÃO - CORREÇÃO]
 * Adicionada para suportar a funcionalidade "Marcar como Sem Pedido".
 * @param {Array<string>} chavesAcesso - As chaves de acesso das NFs a serem atualizadas.
 * @param {string | null} idCotacao - O ID da cotação a ser inserido (pode ser null).
 * @param {string} novoStatus - O novo status para a conciliação ("Conciliada", "Sem Pedido", etc.).
 * @returns {boolean} - True se sucesso, false se erro.
 */
function ConciliacaoNFCrud_atualizarStatusNF(chavesAcesso, idCotacao, novoStatus) {
    try {
        const planilhaNF = SpreadsheetApp.openById(ID_PLANILHA_NF);
        const abaNF = planilhaNF.getSheetByName(ABA_NF_NOTAS_FISCAIS);
        const rangeNF = abaNF.getDataRange();
        const dadosNF = rangeNF.getValues();
        const cabecalhosNF = dadosNF[0];

        const colMapNF = {
            chave: cabecalhosNF.indexOf("Chave de Acesso"),
            status: cabecalhosNF.indexOf("Status da Conciliação"),
            idCot: cabecalhosNF.indexOf("ID da Cotação (Sistema)")
        };

        if (Object.values(colMapNF).includes(-1)) {
            throw new Error("Não foi possível encontrar colunas essenciais (Chave, Status, ID Cotação) na aba de Notas Fiscais.");
        }

        const chavesSet = new Set(chavesAcesso);
        for (let i = 1; i < dadosNF.length; i++) {
            if (chavesSet.has(dadosNF[i][colMapNF.chave])) {
                dadosNF[i][colMapNF.status] = novoStatus;
                if (idCotacao !== null) {
                   dadosNF[i][colMapNF.idCot] = idCotacao;
                }
            }
        }

        rangeNF.setValues(dadosNF);
        Logger.log(`Status de ${chavesAcesso.length} NF(s) atualizado para '${novoStatus}'.`);
        return true;
    } catch (e) {
        Logger.log(`ERRO CRÍTICO em ConciliacaoNFCrud_atualizarStatusNF: ${e.toString()}\n${e.stack}`);
        return false;
    }
}

/**
 * Lê a aba 'Conciliacao' e retorna um array com os mapeamentos existentes.
 * @returns {Array<Object>} Um array de objetos, onde cada objeto é um mapeamento. Ex: [{itemCotacao: 'PROD A', descricaoNF: 'PRODUTO A FORNECEDOR'}]
 */
function ConciliacaoNFCrud_obterMapeamentoConciliacao() {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA_CONCILIACAO);
    if (!aba) {
      Logger.log(`Aba de mapeamento "${ABA_CONCILIACAO}" não encontrada. Retornando array vazio.`);
      return [];
    }
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha <= 1) {
      return [];
    }
    
    const dados = aba.getRange(2, 1, ultimaLinha - 1, 2).getValues(); // Lê apenas as duas primeiras colunas
    const cabecalhos = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];

    const colItemCotacao = cabecalhos.indexOf("Item da Cotação");
    const colDescricaoNF = cabecalhos.indexOf("Descrição Produto (NF)");

    if (colItemCotacao === -1 || colDescricaoNF === -1) {
      Logger.log(`ERRO: Colunas "Item da Cotação" ou "Descrição Produto (NF)" não encontradas na aba "${ABA_CONCILIACAO}".`);
      return [];
    }

    const mapeamento = dados.map(linha => ({
      itemCotacao: linha[colItemCotacao],
      descricaoNF: linha[colDescricaoNF]
    })).filter(item => item.itemCotacao && item.descricaoNF); // Garante que não haja linhas vazias

    return mapeamento;

  } catch (e) {
    Logger.log(`Erro em ConciliacaoNFCrud_obterMapeamentoConciliacao: ${e.message}`);
    return []; // Retorna um array vazio em caso de erro para não quebrar a interface
  }
}


/**
 * Recebe novos mapeamentos e os adiciona na aba 'Conciliacao', evitando duplicatas.
 * @param {Array<Object>} novosMapeamentos - Um array de objetos com as novas associações.
 */
function ConciliacaoNFCrud_atualizarMapeamentoConciliacao(novosMapeamentos) {
  if (!novosMapeamentos || novosMapeamentos.length === 0) {
    return;
  }
  
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA_CONCILIACAO);
    if (!aba) {
      Logger.log(`Aba "${ABA_CONCILIACAO}" não encontrada. Impossível atualizar mapeamento.`);
      return;
    }
    
    const dadosAtuais = aba.getDataRange().getValues();
    const cabecalhos = dadosAtuais.shift(); // Remove o cabeçalho
    
    const colItemCotacao = cabecalhos.indexOf("Item da Cotação");
    const colDescricaoNF = cabecalhos.indexOf("Descrição Produto (NF)");
    const colGtin = cabecalhos.indexOf("GTIN/EAN (Cód. Barras)");

    // Cria um Set para consulta rápida de mapeamentos existentes (Item+Descricao)
    const mapaExistente = new Set(dadosAtuais.map(linha => `${linha[colItemCotacao]}#${linha[colDescricaoNF]}`));

    const linhasParaAdicionar = [];
    
    novosMapeamentos.forEach(item => {
      const chaveUnica = `${item.itemCotacao}#${item.descricaoNF}`;
      if (!mapaExistente.has(chaveUnica)) {
        const novaLinha = cabecalhos.map(() => ''); // Cria uma linha vazia com o tamanho dos cabeçalhos
        novaLinha[colItemCotacao] = item.itemCotacao;
        novaLinha[colDescricaoNF] = item.descricaoNF;
        novaLinha[colGtin] = item.gtin; // Adiciona o GTIN se disponível
        
        linhasParaAdicionar.push(novaLinha);
        mapaExistente.add(chaveUnica); // Adiciona ao Set para evitar duplicatas na mesma execução
      }
    });

    if (linhasParaAdicionar.length > 0) {
      aba.getRange(aba.getLastRow() + 1, 1, linhasParaAdicionar.length, cabecalhos.length)
         .setValues(linhasParaAdicionar);
      Logger.log(`${linhasParaAdicionar.length} novo(s) mapeamento(s) adicionado(s) à aba '${ABA_CONCILIACAO}'.`);
    }

  } catch(e) {
    Logger.log(`ERRO em ConciliacaoNFCrud_atualizarMapeamentoConciliacao: ${e.message}\n${e.stack}`);
  }
}

function ConciliacaoNFCrud_salvarAlteracoesEmLote(conciliacoes, itensCortados, novosMapeamentos) {
    Logger.log(`Iniciando salvamento em lote. Conciliações: ${conciliacoes.length}, Itens cortados: ${itensCortados.length}`);
    try {
        const planilhaCotacoes = SpreadsheetApp.getActiveSpreadsheet();
        const abaCotacoes = planilhaCotacoes.getSheetByName(ABA_COTACOES);
        const rangeCotacoes = abaCotacoes.getDataRange();
        const dadosCotacoes = rangeCotacoes.getValues();
        const cabecalhosCotacoes = dadosCotacoes[0];

        const colMapCotacoes = { id: cabecalhosCotacoes.indexOf("ID da Cotação"), fornecedor: cabecalhosCotacoes.indexOf("Fornecedor"), subProduto: cabecalhosCotacoes.indexOf("SubProduto"), statusSub: cabecalhosCotacoes.indexOf("Status do SubProduto"), divergencia: cabecalhosCotacoes.indexOf("Divergencia da Nota"), qtdNota: cabecalhosCotacoes.indexOf("Quantidade na Nota"), precoNota: cabecalhosCotacoes.indexOf("Preço da Nota"), numeroNota: cabecalhosCotacoes.indexOf("Número da Nota") };
        if (Object.values(colMapCotacoes).includes(-1)) throw new Error("Não foi possível encontrar todas as colunas necessárias na aba de Cotações.");

        conciliacoes.forEach(conc => {
            const mapaItens = new Map(conc.itensConciliados.map(item => [item.subProduto, item]));
            for (let i = 1; i < dadosCotacoes.length; i++) {
                if (dadosCotacoes[i][colMapCotacoes.id] == conc.idCotacao && dadosCotacoes[i][colMapCotacoes.fornecedor] == conc.nomeFornecedor) {
                    const subProdutoLinha = dadosCotacoes[i][colMapCotacoes.subProduto];
                    if (mapaItens.has(subProdutoLinha)) {
                        const itemConciliado = mapaItens.get(subProdutoLinha);
                        dadosCotacoes[i][colMapCotacoes.statusSub] = "Faturado";
                        dadosCotacoes[i][colMapCotacoes.divergencia] = itemConciliado.divergenciaNota;
                        dadosCotacoes[i][colMapCotacoes.qtdNota] = itemConciliado.quantidadeNota;
                        dadosCotacoes[i][colMapCotacoes.precoNota] = itemConciliado.precoNota;
                        dadosCotacoes[i][colMapCotacoes.numeroNota] = conc.numeroNF;
                    }
                }
            }
        });

        const mapaCortados = new Map();
        itensCortados.forEach(item => {
            const key = `${item.idCotacao}-${item.nomeFornecedor}`;
            if (!mapaCortados.has(key)) mapaCortados.set(key, new Set());
            mapaCortados.get(key).add(item.subProduto);
        });
        for (let i = 1; i < dadosCotacoes.length; i++) {
            const key = `${dadosCotacoes[i][colMapCotacoes.id]}-${dadosCotacoes[i][colMapCotacoes.fornecedor]}`;
            if (mapaCortados.has(key)) {
                if (mapaCortados.get(key).has(dadosCotacoes[i][colMapCotacoes.subProduto])) {
                    dadosCotacoes[i][colMapCotacoes.statusSub] = "Cortado";
                }
            }
        }
        rangeCotacoes.setValues(dadosCotacoes);
        Logger.log("Planilha de Cotações atualizada.");

        const chavesAcessoConciliadas = conciliacoes.flatMap(c => c.chavesAcessoNF);
        const idCotacaoPorChave = conciliacoes.reduce((acc, c) => {
          c.chavesAcessoNF.forEach(chave => acc[chave] = c.idCotacao);
          return acc;
        }, {});
        if(chavesAcessoConciliadas.length > 0) {
          const chavesUnicas = [...new Set(chavesAcessoConciliadas)];
          ConciliacaoNFCrud_atualizarStatusNF(chavesUnicas, null, "Conciliada");
          
           const planilhaNF = SpreadsheetApp.openById(ID_PLANILHA_NF);
           const abaNF = planilhaNF.getSheetByName(ABA_NF_NOTAS_FISCAIS);
           const rangeNF = abaNF.getDataRange();
           const dadosNF = rangeNF.getValues();
           const cabecalhosNF = dadosNF[0];
           const colChave = cabecalhosNF.indexOf("Chave de Acesso");
           const colIdCot = cabecalhosNF.indexOf("ID da Cotação (Sistema)");

           for(let i=1; i<dadosNF.length; i++) {
             const chave = dadosNF[i][colChave];
             if(idCotacaoPorChave[chave]){
               dadosNF[i][colIdCot] = idCotacaoPorChave[chave];
             }
           }
           rangeNF.setValues(dadosNF);
           Logger.log("IDs de Cotação atualizados nas Notas Fiscais.");
        }
        
        // ADIÇÃO: Chamar a função para atualizar o mapeamento da aba "Conciliacao"
        if (novosMapeamentos && novosMapeamentos.length > 0) {
          ConciliacaoNFCrud_atualizarMapeamentoConciliacao(novosMapeamentos);
        }
        
        return true;
    } catch (e) {
        Logger.log(`ERRO CRÍTICO em ConciliacaoNFCrud_salvarAlteracoesEmLote: ${e.toString()}\n${e.stack}`);
        return false;
    }
}

function ConciliacaoNFCrud_obterTodosItensCotacoesAbertas(chavesCotacoes) {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA_COTACOES);
    const dados = aba.getDataRange().getValues();
    const cabecalhos = dados.shift();
    const colMap = {};
    const colunasNecessarias = ["ID da Cotação", "Fornecedor", "SubProduto", "Comprar", "Preço", "Fator", "Preço por Fator"];
    colunasNecessarias.forEach(nome => { colMap[nome] = cabecalhos.indexOf(nome); });
    const setCotacoes = new Set(chavesCotacoes.map(c => `${c.idCotacao}-${c.fornecedor}`));
    const itens = [];
    for (const linha of dados) {
        const id = linha[colMap["ID da Cotação"]];
        const fornecedor = linha[colMap["Fornecedor"]];
        const compositeKey = `${id}-${fornecedor}`;
        if (setCotacoes.has(compositeKey)) {
            const qtdComprar = parseFloat(linha[colMap["Comprar"]]);
            if (!isNaN(qtdComprar) && qtdComprar > 0) {
                 itens.push({
                    idCotacao: id, fornecedor: fornecedor, subProduto: linha[colMap["SubProduto"]],
                    qtdComprar: qtdComprar, preco: parseFloat(linha[colMap["Preço"]]) || 0,
                    fator: parseFloat(linha[colMap["Fator"]]) || 1, precoPorFator: parseFloat(linha[colMap["Preço por Fator"]]) || 0
                });
            }
        }
    }
    return itens;
}

/**
 * Lê a aba 'Conciliacao' e retorna um array com os mapeamentos existentes.
 * @returns {Array<Object>} Um array de objetos, onde cada objeto é um mapeamento. Ex: [{itemCotacao: 'PROD A', descricaoNF: 'PRODUTO A FORNECEDOR'}]
 */
function ConciliacaoNFCrud_obterMapeamentoConciliacao() {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA_CONCILIACAO);
    if (!aba) {
      Logger.log(`Aba de mapeamento "${ABA_CONCILIACAO}" não encontrada. Retornando array vazio.`);
      return [];
    }
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha <= 1) {
      return [];
    }
    
    const dados = aba.getRange(2, 1, ultimaLinha - 1, 2).getValues(); // Lê apenas as duas primeiras colunas
    const cabecalhos = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];

    const colItemCotacao = cabecalhos.indexOf("Item da Cotação");
    const colDescricaoNF = cabecalhos.indexOf("Descrição Produto (NF)");

    if (colItemCotacao === -1 || colDescricaoNF === -1) {
      Logger.log(`ERRO: Colunas "Item da Cotação" ou "Descrição Produto (NF)" não encontradas na aba "${ABA_CONCILIACAO}".`);
      return [];
    }

    const mapeamento = dados.map(linha => ({
      itemCotacao: linha[colItemCotacao],
      descricaoNF: linha[colDescricaoNF]
    })).filter(item => item.itemCotacao && item.descricaoNF); // Garante que não haja linhas vazias

    return mapeamento;

  } catch (e) {
    Logger.log(`Erro em ConciliacaoNFCrud_obterMapeamentoConciliacao: ${e.message}`);
    return []; // Retorna um array vazio em caso de erro para não quebrar a interface
  }
}


/**
 * Recebe novos mapeamentos e os adiciona na aba 'Conciliacao', evitando duplicatas.
 * @param {Array<Object>} novosMapeamentos - Um array de objetos com as novas associações.
 */
function ConciliacaoNFCrud_atualizarMapeamentoConciliacao(novosMapeamentos) {
  if (!novosMapeamentos || novosMapeamentos.length === 0) {
    return;
  }
  
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA_CONCILIACAO);
    if (!aba) {
      Logger.log(`Aba "${ABA_CONCILIACAO}" não encontrada. Impossível atualizar mapeamento.`);
      return;
    }
    
    const dadosAtuais = aba.getDataRange().getValues();
    const cabecalhos = dadosAtuais.shift(); // Remove o cabeçalho
    
    const colItemCotacao = cabecalhos.indexOf("Item da Cotação");
    const colDescricaoNF = cabecalhos.indexOf("Descrição Produto (NF)");
    const colGtin = cabecalhos.indexOf("GTIN/EAN (Cód. Barras)");

    // Cria um Set para consulta rápida de mapeamentos existentes (Item+Descricao)
    const mapaExistente = new Set(dadosAtuais.map(linha => `${linha[colItemCotacao]}#${linha[colDescricaoNF]}`));

    const linhasParaAdicionar = [];
    
    novosMapeamentos.forEach(item => {
      const chaveUnica = `${item.itemCotacao}#${item.descricaoNF}`;
      if (!mapaExistente.has(chaveUnica)) {
        const novaLinha = cabecalhos.map(() => ''); // Cria uma linha vazia com o tamanho dos cabeçalhos
        novaLinha[colItemCotacao] = item.itemCotacao;
        novaLinha[colDescricaoNF] = item.descricaoNF;
        novaLinha[colGtin] = item.gtin; // Adiciona o GTIN se disponível
        
        linhasParaAdicionar.push(novaLinha);
        mapaExistente.add(chaveUnica); // Adiciona ao Set para evitar duplicatas na mesma execução
      }
    });

    if (linhasParaAdicionar.length > 0) {
      aba.getRange(aba.getLastRow() + 1, 1, linhasParaAdicionar.length, cabecalhos.length)
         .setValues(linhasParaAdicionar);
      Logger.log(`${linhasParaAdicionar.length} novo(s) mapeamento(s) adicionado(s) à aba '${ABA_CONCILIACAO}'.`);
    }

  } catch(e) {
    Logger.log(`ERRO em ConciliacaoNFCrud_atualizarMapeamentoConciliacao: ${e.message}\n${e.stack}`);
  }
}