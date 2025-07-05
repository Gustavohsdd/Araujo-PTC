// =================================================================================
// ARQUIVO: ConstantesFinanceiro.gs
// DESCRIÇÃO: Constantes para o sistema de Rateio e Contas a pagar
// =================================================================================

// --- IDs GLOBAIS ---
// ID da Planilha que armazena os dados de Boletos e Rateios.
<<<<<<< HEAD
const ID_PLANILHA_FINANCEIRO = '1JNJrbZqXUdG28KlZMhmgOzxEuVKDrEpi5t08yBTFnzw';
=======
const ID_PLANILHA_FINANCEIRO = '1WXcYsH3rl6px5zVPLCiHI4Vfo9y_563lqcc9wEk63eA';
>>>>>>> 1415d4a6ea5bc830f28cbb5bd785bc2e6e55d279

// --- NOMES DAS ABAS ---
const ABA_FINANCEIRO_REGRAS_RATEIO = 'RegrasRateio';
const ABA_FINANCEIRO_CONTAS_A_PAGAR = 'ContasAPagar';

// --- CABEÇALHOS DAS ABAS ---

/**
 * Cabeçalhos para a aba 'RegrasRateio'.
 */
const CABECALHOS_FINANCEIRO_REGRAS_RATEIO = [
  "Item da Cotação",
  "Setor",
  "Porcentagem"
];

/**
 * Cabeçalhos para a aba 'ContasAPagar'.
 */
const CABECALHOS_FINANCEIRO_CONTAS_A_PAGAR = [
  "Chave de Acesso",
  "Número da Fatura",
  "Número da Parcela",
  "Resumo dos Itens",
  "Data de Vencimento",
  "Valor da Parcela",
  "Setor",
  "Valor por Setor"
];