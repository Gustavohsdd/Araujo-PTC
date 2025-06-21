// @ts-nocheck

// Constantes globais de App.gs
// --- MODIFICAÇÃO INÍCIO ---
const App_VIEWS_PERMITIDAS = ["fornecedores", "produtos", "subprodutos", "cotacoes", "cotacaoIndividual", "contagemdeestoque", "EnviarManualmenteView", "marcarprodutos"]; 
// --- MODIFICAÇÃO FIM ---

const App_VIEW_FILENAME_MAP = {
  "fornecedores": "FornecedoresView",
  "produtos": "ProdutosView",
  "subprodutos": "SubProdutosView",
  "cotacoes": "CotacoesView",
  "cotacaoIndividual": "CotacaoIndividualView",
  "contagemdeestoque": "ContagemDeEstoqueView",
  "EnviarManualmenteView": "EnviarManualmenteView",
  // --- MODIFICAÇÃO INÍCIO ---
  "marcarprodutos": "MarcacaoProdutosView"
  // --- MODIFICAÇÃO FIM ---
}

  const WEB_APP_URL_PROJETO_ATUAL = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL'); 

/**
 * Função principal para servir o Web App.
 * Chamada quando um usuário acessa a URL do Web App.
 * @param {GoogleAppsScript.Events.DoGet} e O objeto de evento.
 * @return {HtmlOutput} O HTML a ser renderizado.
 */
function doGet(e) {
  Logger.log(`App.gs: Requisição GET recebida: ${JSON.stringify(e)}`);
  
  const token = e.parameter.token;
  const view = e.parameter.view;

  // Lógica para o Portal do Fornecedor
  if (token) {
    Logger.log(`App.gs: Token encontrado na URL: ${token}. Tentando carregar Portal do Fornecedor.`);
    try {
      // Esta função agora retorna também a 'dataAberturaFormatada'
      const dadosPortal = PortalController_buscarDadosParaPortalWebService(token);
      
      if (!dadosPortal.valido) {
        Logger.log(`App.gs: Token inválido ou erro ao buscar dados para o portal. Token: ${token}. Mensagem: ${dadosPortal.mensagemErro}`);
        // Você pode criar uma página de erro mais elaborada aqui se desejar
        return HtmlService.createHtmlOutput(`<html><body><h1>Acesso Inválido</h1><p>${dadosPortal.mensagemErro}</p></body></html>`)
                         .setTitle("Acesso Inválido ao Portal");
      }

      const htmlTemplate = HtmlService.createTemplateFromFile('PortalFornecedorView'); 
      
      // Propriedades existentes sendo passadas para o template
      htmlTemplate.nomeFornecedor = dadosPortal.nomeFornecedor;
      htmlTemplate.idCotacao = dadosPortal.idCotacao;
      htmlTemplate.produtos = dadosPortal.produtos; 
      htmlTemplate.token = token; 
      htmlTemplate.status = dadosPortal.status; 
      htmlTemplate.pedidoFinalizado = dadosPortal.pedidoFinalizado; 

      // <<< LINHA ADICIONADA PARA A PRÉ-VISUALIZAÇÃO >>>
      // Passa a nova data de abertura formatada para o template HTML.
      htmlTemplate.dataAberturaFormatada = dadosPortal.dataAberturaFormatada;
      
      const finalHtmlOutput = htmlTemplate.evaluate();
      finalHtmlOutput.setTitle(`Cotação ${dadosPortal.idCotacao} - ${dadosPortal.nomeFornecedor}`)
                     .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      return finalHtmlOutput;

    } catch (err) {
      Logger.log(`App.gs: ERRO CRÍTICO ao tentar carregar o Portal do Fornecedor para token ${token}: ${err}\n${err.stack}`);
      return HtmlService.createHtmlOutput('<html><body><h1>Erro Crítico</h1><p>Ocorreu um erro inesperado ao carregar o portal. Por favor, contate o suporte.</p></body></html>')
                       .setTitle("Erro no Portal");
    }
  } 
  
  // Roteamento baseado no parâmetro 'view'
  switch (view) {
    case 'cotacaoIndividual':
      Logger.log("App.gs: Carregando CotacaoIndividualView.");
      let templateCotacao = HtmlService.createTemplateFromFile('CotacaoIndividualView');
      templateCotacao.idCotacao = e.parameter.idCotacao || null;
      templateCotacao.modo = e.parameter.modo || 'editar';
      return templateCotacao.evaluate()
        .setTitle("Cotação: " + e.parameter.idCotacao + " - CotaçãoPRO")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'contagemdeestoque':
      Logger.log("App.js: Carregando ContagemDeEstoqueView.");
      let templateContagem = HtmlService.createTemplateFromFile('ContagemDeEstoqueView');
      templateContagem.idCotacao = e.parameter.idCotacao || null;
      return templateContagem.evaluate()
        .setTitle("Contagem de Estoque - Cotação " + (e.parameter.idCotacao || ""))
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'ImprimirPedidosView':
      Logger.log("App.gs: Carregando ImprimirPedidosView.");
      return App_carregarPaginaDeImpressao(e);

    case 'EnviarManualmenteView':
      Logger.log("App.gs: Carregando EnviarManualmenteView.");
      const templateEnvioManual = HtmlService.createTemplateFromFile('EnviarManualmenteView');
      return templateEnvioManual.evaluate()
        .setTitle("Enviar Pedidos Manualmente")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);

    case 'RelatorioAnaliseCompraView':
      Logger.log("App.gs: Carregando RelatorioAnaliseCompraView.");
      const templateRelatorio = HtmlService.createTemplateFromFile('RelatorioAnaliseCompraView');
      templateRelatorio.idCotacao = e.parameter.idCotacao || null;
      return templateRelatorio.evaluate()
        .setTitle("Relatório de Análise - Cotação " + e.parameter.idCotacao)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);

    // --- MODIFICAÇÃO INÍCIO ---
    case 'marcarprodutos':
      Logger.log("App.gs: Carregando MarcacaoProdutosView.");
      const templateMarcacao = HtmlService.createTemplateFromFile('MarcacaoProdutosView');
      return templateMarcacao.evaluate()
        .setTitle("Marcação de Recebimento")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1') // Garante responsividade
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
    // --- MODIFICAÇÃO FIM ---
        
    default:
      // Comportamento padrão: carregar a página principal
      Logger.log("App.gs: Carregando PaginaPrincipal (comportamento padrão).");
      return HtmlService.createTemplateFromFile('PaginaPrincipal')
        .evaluate()
        .setTitle("CotaçãoPRO")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Inclui o conteúdo de um arquivo HTML.
 */
function App_incluirHtml(nomeArquivo, isScriptContent = false) {
  try {
    return HtmlService.createHtmlOutputFromFile(nomeArquivo).getContent();
  } catch (e) {
    console.error("Erro ao incluir arquivo HTML '" + nomeArquivo + "': " + e.toString());
    if (isScriptContent) {
      return `console.error("Falha ao carregar o conteúdo de '${nomeArquivo}'. Detalhes: ${e.message.replace(/"/g, '\\"').replace(/\n/g, '\\n')}");`;
    }
    return `<div class="error-message p-2">Erro ao incluir ${nomeArquivo}: ${e.message}</div>`;
  }
}

function App_obterView(viewName) {
  if (!App_VIEWS_PERMITIDAS.includes(viewName)) {
    console.error("Tentativa de acesso a view inválida: " + viewName);
    return `<div class="error-message p-4">View inválida solicitada: ${viewName}</div>`;
  }

  const nomeArquivoHtml = App_VIEW_FILENAME_MAP[viewName];

  if (!nomeArquivoHtml) {
    console.error("Mapeamento de nome de arquivo não encontrado para a view: " + viewName);
    return `<div class="error-message p-4">Configuração interna de view não encontrada para: ${viewName}</div>`;
  }

  try {
    const template = HtmlService.createTemplateFromFile(nomeArquivoHtml);
    return template.evaluate().getContent();
  } catch (error) {
    console.error("Erro ao carregar o arquivo HTML '" + nomeArquivoHtml + "' para a view '" + viewName + "': " + error.toString() + " Stack: " + error.stack);
    return `<div class="error-message p-4">Erro ao carregar a view '${viewName}'. Verifique se o arquivo '${nomeArquivoHtml}.html' existe e está correto. Detalhes: ${error.message}</div>`;
  }
}

/**
 * Retorna um objeto contendo constantes selecionadas para uso no lado do cliente.
 * @return {object} Um objeto com as constantes.
 */
function App_obterConstantes() {
  console.log("App_obterConstantes: Solicitado.");
  try {
    if (typeof CABECALHOS_COTACOES === 'undefined') {
      console.error("App_obterConstantes: A constante CABECALHOS_COTACOES não está definida ou acessível.");
      throw new Error("Constante CABECALHOS_COTACOES não definida.");
    }
    const colunasSubProdutos = (typeof COLUNAS_PARA_ABA_SUBPRODUTOS !== 'undefined') ? COLUNAS_PARA_ABA_SUBPRODUTOS : [];

    const constantesParaCliente = {
      CABECALHOS_COTACOES: CABECALHOS_COTACOES,
      COLUNAS_PARA_ABA_SUBPRODUTOS: colunasSubProdutos 
    };
    console.log("App_obterConstantes: Retornando constantes:", JSON.stringify(constantesParaCliente));
    return constantesParaCliente;
  } catch (e) {
    console.error("Erro em App_obterConstantes: " + e.toString());
    return { error: true, message: "Erro ao obter constantes: " + e.message };
  }
}

/**
 * Função para gerar a URL para a página de cotação individual.
 */
function App_obterUrlCotacaoIndividual(idCotacao, modo) {
  const urlBase = ScriptApp.getService().getUrl();
  let urlComParametros = `${urlBase}?view=cotacaoIndividual&modo=${encodeURIComponent(modo || 'editar')}`;
  if (idCotacao) {
    urlComParametros += `&idCotacao=${encodeURIComponent(idCotacao)}`;
  }
  console.log("App_obterUrlCotacaoIndividual: Gerada URL:", urlComParametros);
  return urlComParametros;
}

/**
 * Retorna a URL para a página de Contagem de Estoque Mobile.
 */
function App_obterUrlContagemDeEstoque(idCotacao) {
  const urlBase = ScriptApp.getService().getUrl();
  let url = `${urlBase}?view=contagemdeestoque`;
  if (idCotacao) url += `&idCotacao=${encodeURIComponent(idCotacao)}`;
  console.log("App_obterUrlContagemDeEstoque: Gerada URL:", url);
  return url;
}

/**
 * Função para configurar a URL do Web App nas propriedades do script MANUALMENTE.
 * Cole a URL do seu Web App na variável 'urlWebAppManualmente' e execute esta função uma vez.
 */
function App_configurarUrlWebApp() {
  // ########## COLE A URL DO SEU WEB APP AQUI ENTRE AS ASPAS ##########
  const urlWebAppManualmente = "https://script.google.com/macros/s/AKfycbx5oXaCsdkHlXTIicJTCTW_VVNIv9T979WsLYatuBBVRxPh8aIbh1YI1GXxfwWMtBDT/exec"; 
  // #####################################################################

  if (urlWebAppManualmente && urlWebAppManualmente.startsWith("https://script.google.com/macros/s/")) {
    PropertiesService.getScriptProperties().setProperty('WEB_APP_URL', urlWebAppManualmente);
    Logger.log(`URL do Web App configurada com sucesso nas Propriedades do Script: ${urlWebAppManualmente}`);
    // Removida a interação com SpreadsheetApp.getUi().alert
    // A confirmação será apenas pelo Logger.log.
  } else {
    const mensagemErro = 'URL Inválida. Verifique se a URL fornecida na função App_configurarUrlWebApp é uma URL válida de Web App do Google Apps Script (deve começar com "https://script.google.com/macros/s/").';
    Logger.log(mensagemErro); // Alterado para Logger.log em vez de console.error para consistência
  }
}

/**
 * Renderiza a página de impressão de pedidos.
 * @param {object} e O objeto de evento do Apps Script.
 * @returns {HtmlOutput} A página HTML renderizada.
 */
function App_carregarPaginaDeImpressao(e) {
  // Passa os parâmetros da URL (como idCotacao) para o template, se necessário
  const template = HtmlService.createTemplateFromFile('ImprimirPedidosView');
  template.idCotacao = e.parameter.idCotacao || null;

  return template.evaluate()
      .setTitle('Impressão de Pedidos - Cotação ' + e.parameter.idCotacao)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Cria a URL completa para uma determinada visualização do Web App com parâmetros.
 * Esta função é chamada pelo client-side para obter a URL que será aberta em uma nova aba.
 * @param {object} params Objeto com os parâmetros para a URL. Ex: {view: '...', idCotacao: '...'}
 * @returns {{success: boolean, url?: string, message?: string}}
 */
function App_obterUrlWebApp(params) {
  try {
    let url = ScriptApp.getService().getUrl();
    
    // ===== INÍCIO DA CORREÇÃO =====
    // A condição foi ajustada para apenas adicionar o '?' se houver chaves no objeto de parâmetros.
    if (params && Object.keys(params).length > 0) {
      const query = Object.keys(params)
        .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`)
        .join('&');
      url += `?${query}`;
    }
    // ===== FIM DA CORREÇÃO =====

    Logger.log(`App_obterUrlWebApp: URL gerada -> ${url}`);
    return { success: true, url: url };

  } catch(e) {
    Logger.log(`ERRO em App_obterUrlWebApp: ${e.toString()}`);
    return { success: false, message: e.message };
  }
}
