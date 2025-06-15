// @ts-nocheck
// Constantes.gs

// Nome da planilha principal
const NOME_PLANILHA = "CadastroGestao";

// Aba Fornecedores
const ABA_FORNECEDORES = "Fornecedores";
const CABECALHOS_FORNECEDORES = [
  "Data de Cadastro",
  "ID",
  "Fornecedor",
  "CNPJ",
  "Categoria",
  "Vendedor",
  "Telefone",
  "Email",
  "Dias de Pedido",
  "Dias de Entrega",
  "Condições de Pagamento",
  "Dia de Faturamento",
  "Pedido Mínimo (R$)",
  "Regime Tributário",
  "Contato Financeiro"
];

// Aba Produtos
const ABA_PRODUTOS = "Produtos";
const CABECALHOS_PRODUTOS = [
  "Data de Cadastro",
  "ID",
  "Produto",
  "ABC",
  "Categoria",
  "Tamanho",
  "UN",
  "Estoque Minimo",
  "Status"
];

// Aba SubProdutos
const ABA_SUBPRODUTOS = "SubProdutos";
const CABECALHOS_SUBPRODUTOS = [
  "Data de Cadastro",
  "ID",
  "SubProduto",
  "Produto Vinculado",
  "Categoria",
  "Fornecedor",
  "Tamanho",
  "UN",
  "Fator",
  "NCM",
  "CST",
  "CFOP",
  "Status"
];

// Aba Cotações
const ABA_COTACOES = "Cotacoes";
const CABECALHOS_COTACOES = [
  "ID da Cotação",
  "Data Abertura",
  "Produto",
  "SubProduto",
  "Categoria",
  "Fornecedor",
  "Tamanho",
  "UN",
  "Fator",
  "Estoque Mínimo",
  "Estoque Atual",
  "Preço",
  "Preço por Fator",
  "Comprar",
  "Valor Total",
  "Economia em Cotação",
  "NCM",
  "CST",
  "CFOP",
  "Empresa Faturada",
  "Condição de Pagamento",
  "Status da Cotação"
];

// Aba Pedidos
const ABA_PEDIDOS = "Pedidos";
const CABECALHOS_PEDIDOS = [
  "ID do Pedido",
  "Data Abertura",
  "Produto",
  "SubProduto",
  "Categoria",
  "Fornecedor",
  "Tamanho",
  "UN",
  "Fator",
  "Estoque Mínimo",
  "Estoque Atual",
  "Preço",
  "Preço por Fator",
  "Comprar",
  "Valor Total",
  "Economia em Cotação",
  "NCM",
  "CST",
  "CFOP",
  "Empresa Faturada",
  "Condição de Pagamento",
  "Status do Pedido"
];

// Aba Portal
const ABA_PORTAL = "Portal";
const CABECALHOS_PORTAL = [
  "ID da Cotação",
  "Nome Fornecedor",
  "Token Acesso",
  "Link Acesso",
  "Status",
  "Data Envio",
  "Data Resposta",
  "Texto Personalizado Link"
];

// Aba Cadastros
const ABA_CADASTROS = "Cadastros";
const CABECALHOS_CADASTROS = [
  "Empresas",
  "CNPJ",
  "Endereço",
  "Telefone",
  "Email",
  "Contato"
];

// Constantes para o Portal do Fornecedor
const STATUS_PORTAL = {
  LINK_GERADO: "Link Gerado", // Link criado, aguardando primeira resposta
  RESPONDIDO: "Respondido",   // Fornecedor enviou/finalizou os preços
  EM_PREENCHIMENTO: "Em Preenchimento", // (Opcional) Se quiser rastrear edições parciais antes de finalizar
  FECHADO: "Fechado",         // Cotação encerrada pelo comprador
  ERRO_PORTAL: "Erro no Portal", // Erro ao processar o link/dados
  EXPIRADO: "Expirado"        // (Opcional) Se links tiverem validade
};
Object.freeze(STATUS_PORTAL);
