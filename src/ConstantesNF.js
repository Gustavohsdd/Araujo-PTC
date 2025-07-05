// =================================================================================
// ARQUIVO: ConstantesNF.gs
// DESCRIÇÃO: Constantes para o sistema de leitura e conciliação de Notas Fiscais (NF-e)
// =================================================================================

// --- IDs GLOBAIS ---
// ID da Planilha que armazena os dados extraídos das Notas Fiscais.
const ID_PLANILHA_NF = '1-YKacsqlFJ7ijRY1Vsba4t5OKfyj0Gmc8l1W0POEJJo';

// ID da Pasta no Google Drive onde os arquivos XML das NF-e são armazenados.
const ID_PASTA_XML = '188nkMfRmqwqQ5kwZeviTjdv8v19g1VyQ';


// --- NOMES DAS ABAS ---
const ABA_NF_NOTAS_FISCAIS = 'NotasFiscais';
const ABA_NF_ITENS = 'ItensNF';
const ABA_NF_FATURAS = 'FaturasNF';
const ABA_NF_TRANSPORTE = 'TransporteNF';
const ABA_NF_TRIBUTOS_TOTAIS = 'TributosTotaisNF';


// --- CABEÇALHOS DAS ABAS ---

/**
 * Cabeçalhos para a aba 'NotasFiscais'.
 * Contém os dados gerais de cada nota fiscal processada.
 */
const CABECALHOS_NF_NOTAS_FISCAIS = [
  "Chave de Acesso",
  "ID da Cotação (Sistema)",
  "Status da Conciliação",
  "Número NF",
  "Série NF",
  "Data e Hora Emissão",
  "Natureza da Operação",
  "CNPJ Emitente",
  "Nome Emitente",
  "Inscrição Estadual Emitente",
  "Logradouro Emitente",
  "Número End. Emitente",
  "Bairro Emitente",
  "Município Emitente",
  "UF Emitente",
  "CEP Emitente",
  "CNPJ Destinatário",
  "Nome Destinatário",
  "Informações Adicionais",
  "Número do Pedido (Extraído)",
  "Status do Rateio"
];

/**
 * Cabeçalhos para a aba 'ItensNF'.
 * Contém o detalhamento de cada produto/serviço da nota fiscal.
 */
const CABECALHOS_NF_ITENS = [
  "Chave de Acesso",
  "Número do Item",
  "Código Produto (Forn)",
  "GTIN/EAN (Cód. Barras)",
  "Descrição Produto (NF)",
  "NCM",
  "CFOP",
  "Unidade Comercial",
  "Quantidade Comercial",
  "Valor Unitário Comercial",
  "Valor Total Bruto Item",
  "Valor do Frete (Item)",
  "Valor do Seguro (Item)",
  "Valor do Desconto (Item)",
  "Outras Despesas (Item)",
  "CST/CSOSN (ICMS)",
  "Base de Cálculo (ICMS)",
  "Alíquota (ICMS)",
  "Valor (ICMS)",
  "Valor (ICMS ST)",
  "CST (IPI)",
  "Base de Cálculo (IPI)",
  "Alíquota (IPI)",
  "Valor (IPI)",
  "CST (PIS)",
  "Valor (PIS)",
  "CST (COFINS)",
  "Valor (COFINS)"
];

/**
 * Cabeçalhos para a aba 'FaturasNF'.
 * Contém os dados de cobrança e parcelas (duplicatas).
 */
const CABECALHOS_NF_FATURAS = [
  "Chave de Acesso",
  "Número da Fatura",
  "Número da Parcela",
  "Data de Vencimento",
  "Valor da Parcela"
];

/**
 * Cabeçalhos para a aba 'TransporteNF'.
 * Contém os dados da transportadora e do frete.
 */
const CABECALHOS_NF_TRANSPORTE = [
  "Chave de Acesso",
  "Modalidade Frete",
  "CNPJ Transportadora",
  "Nome Transportadora",
  "IE Transportadora",
  "Endereço Transportadora",
  "Placa Veículo",
  "Quantidade Volumes",
  "Espécie Volumes",
  "Peso Líquido Total",
  "Peso Bruto Total"
];

/**
 * Cabeçalhos para a aba 'TributosTotaisNF'.
 * Contém a totalização dos impostos da nota fiscal.
 */
const CABECALHOS_NF_TRIBUTOS_TOTAIS = [
  "Chave de Acesso",
  "Total Base Cálculo ICMS",
  "Total Valor ICMS",
  "Total Valor ICMS ST",
  "Total Valor Produtos",
  "Total Valor Frete",
  "Total Valor Seguro",
  "Total Valor Desconto",
  "Total Valor IPI",
  "Total Valor PIS",
  "Total Valor COFINS",
  "Total Outras Despesas",
  "Valor Total da NF"
];