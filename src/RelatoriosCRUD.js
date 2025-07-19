// @ts-nocheck
/**
 * @file RelatoriosCRUD.gs
 * @description Funções de acesso a dados para gerar relatórios.
 */

// =================================================================================
// CONSTANTES NECESSÁRIAS PARA O RELATÓRIO
// Para garantir que o script funcione, todas as constantes estão definidas aqui.
// =================================================================================

// ID da Planilha que armazena os dados extraídos das Notas Fiscais.
const RelatoriosCRUD_ID_PLANILHA_NF = '1-YKacsqlFJ7ijRY1Vsba4t5OKfyj0Gmc8l1W0POEJJo';

// --- CONSTANTES DAS ABAS DE NOTAS FISCAIS (NF-e) ---
const RelatoriosCRUD_ABA_NF_NOTAS_FISCAIS = 'NotasFiscais';
const RelatoriosCRUD_ABA_NF_ITENS = 'ItensNF';
const RelatoriosCRUD_CABECALHOS_NF_NOTAS_FISCAIS = ["Chave de Acesso", "ID da Cotação (Sistema)", "Status da Conciliação", "Número NF", "Série NF", "Data e Hora Emissão", "Natureza da Operação", "CNPJ Emitente", "Nome Emitente", "Inscrição Estadual Emitente", "Logradouro Emitente", "Número End. Emitente", "Bairro Emitente", "Município Emitente", "UF Emitente", "CEP Emitente", "CNPJ Destinatário", "Nome Destinatário", "Informações Adicionais", "Número do Pedido (Extraído)", "Status do Rateio"];
const RelatoriosCRUD_CABECALHOS_NF_ITENS = ["Chave de Acesso", "Número do Item", "Código Produto (Forn)", "GTIN/EAN (Cód. Barras)", "Descrição Produto (NF)", "NCM", "CFOP", "Unidade Comercial", "Quantidade Comercial", "Valor Unitário Comercial", "Valor Total Bruto Item", "Valor do Frete (Item)", "Valor do Seguro (Item)", "Valor do Desconto (Item)", "Outras Despesas (Item)", "CST/CSOSN (ICMS)", "Base de Cálculo (ICMS)", "Alíquota (ICMS)", "Valor (ICMS)", "Valor (ICMS ST)", "CST (IPI)", "Base de Cálculo (IPI)", "Alíquota (IPI)", "Valor (IPI)", "CST (PIS)", "Valor (PIS)", "CST (COFINS)", "Valor (COFINS)"];

// --- CONSTANTES DAS ABAS DE GESTÃO (PLANILHA ATIVA) ---
const RelatoriosCRUD_ABA_COTACOES = "Cotacoes";
const RelatoriosCRUD_CABECALHOS_COTACOES = ["ID da Cotação", "Data Abertura", "Produto", "SubProduto", "Categoria", "Fornecedor", "Tamanho", "UN", "Fator", "Estoque Mínimo", "Estoque Atual", "Preço", "Preço por Fator", "Comprar", "Valor Total", "Economia em Cotação", "NCM", "CST", "CFOP", "Empresa Faturada", "Condição de Pagamento", "Status da Cotação", "Status do SubProduto", "Quantidade Recebida", "Divergencia da Nota", "Quantidade na Nota", "Preço da Nota", "Número da Nota"];

const RelatoriosCRUD_ABA_CONCILIACAO = "Conciliacao";
const RelatoriosCRUD_CABECALHOS_CONCILIACAO = ["Item da Cotação", "Descrição Produto (NF)", "GTIN/EAN (Cód. Barras)"];


/**
 * Função auxiliar CORRIGIDA para ler dados de uma aba, seja na planilha ativa ou em uma externa.
 * @param {string} nomeAba O nome da aba da planilha.
 * @param {Array<string>} cabecalhosEsperados O array com os nomes das colunas.
 * @param {string|null} spreadsheetId O ID da planilha externa, se houver.
 * @return {Array<object>} Um array de objetos, onde cada objeto representa uma linha.
 */
function RelatoriosCRUD_lerDadosDeAba(nomeAba, cabecalhosEsperados, spreadsheetId = null) {
    try {
        let planilha;
        // Se um ID for fornecido, abre a planilha externa. Caso contrário, usa a ativa.
        if (spreadsheetId) {
            planilha = SpreadsheetApp.openById(spreadsheetId);
            if (!planilha) {
                Logger.log(`Erro Crítico: Não foi possível abrir a planilha com ID: ${spreadsheetId}`);
                return [];
            }
        } else {
            planilha = SpreadsheetApp.getActiveSpreadsheet();
        }
        
        const aba = planilha.getSheetByName(nomeAba);
        if (!aba) {
            Logger.log(`Aviso Crítico: A aba "${nomeAba}" não foi encontrada na planilha de referência. (ID: ${spreadsheetId || 'Ativa'})`);
            return [];
        }
        const dados = aba.getDataRange().getValues();
        const cabecalhosReais = dados.shift() || [];
        
        return dados.map(linha => {
            const objetoLinha = {};
            cabecalhosEsperados.forEach(cabecalho => {
                const index = cabecalhosReais.indexOf(cabecalho);
                objetoLinha[cabecalho] = (index !== -1) ? linha[index] : null;
            });
            return objetoLinha;
        });
    } catch (e) {
        Logger.log(`Erro ao ler a aba "${nomeAba}": ${e.toString()}`);
        return [];
    }
}

/**
 * Gera os dados para o Relatório de Análise de Compra.
 * ATUALIZADO: Calcula e inclui o valor unitário de cada imposto e seu percentual
 * sobre o preço unitário da NF.
 */
function RelatoriosCRUD_gerarDadosAnaliseCompra(idCotacaoAlvo) {
    Logger.log(`Iniciando geração do Relatório de Análise de Compra para Cotação ID: ${idCotacaoAlvo}`);
    try {
        // 1. Leitura das abas
        const dadosCotacoes = RelatoriosCRUD_lerDadosDeAba(RelatoriosCRUD_ABA_COTACOES, RelatoriosCRUD_CABECALHOS_COTACOES);
        const dadosConciliacao = RelatoriosCRUD_lerDadosDeAba(RelatoriosCRUD_ABA_CONCILIACAO, RelatoriosCRUD_CABECALHOS_CONCILIACAO);
        const dadosNotasFiscais = RelatoriosCRUD_lerDadosDeAba(RelatoriosCRUD_ABA_NF_NOTAS_FISCAIS, RelatoriosCRUD_CABECALHOS_NF_NOTAS_FISCAIS, RelatoriosCRUD_ID_PLANILHA_NF);
        const dadosItensNF = RelatoriosCRUD_lerDadosDeAba(RelatoriosCRUD_ABA_NF_ITENS, RelatoriosCRUD_CABECALHOS_NF_ITENS, RelatoriosCRUD_ID_PLANILHA_NF);

        // 2. Identifica os produtos únicos da cotação alvo.
        const produtosDaCotacao = new Set(
            dadosCotacoes
            .filter(linha => String(linha["ID da Cotação"]) === String(idCotacaoAlvo))
            .map(linha => linha["Produto"])
        );

        // 3. Mapeamentos essenciais
        const mapaNotas = dadosNotasFiscais.reduce((map, nf) => {
            if (nf["Chave de Acesso"]) {
                map[nf["Chave de Acesso"]] = {
                    nomeEmitente: nf["Nome Emitente"],
                    dataEmissao: new Date(nf["Data e Hora Emissão"]),
                    numeroNF: nf["Número NF"]
                };
            }
            return map;
        }, {});

        const mapaConciliacaoNFparaProduto = dadosConciliacao.reduce((map, item) => {
            if (item["Descrição Produto (NF)"] && item["Item da Cotação"]) {
                map[item["Descrição Produto (NF)"].trim()] = item["Item da Cotação"].trim();
            }
            return map;
        }, {});

        // 4. Encontrar as últimas compras por PRODUTO e NOME BRUTO DO FORNECEDOR DA NF.
        const ultimasComprasReais = {};
        dadosItensNF.forEach(itemNF => {
            const descProdutoNF = itemNF["Descrição Produto (NF)"];
            if (!descProdutoNF) return;

            const produtoPrincipal = mapaConciliacaoNFparaProduto[descProdutoNF.trim()];
            const notaFiscal = mapaNotas[itemNF["Chave de Acesso"]];

            if (produtoPrincipal && notaFiscal && notaFiscal.nomeEmitente) {
                const nomeFornecedorNF = notaFiscal.nomeEmitente.trim();
                const chaveFinal = `${produtoPrincipal}|${nomeFornecedorNF}`;

                if (!ultimasComprasReais[chaveFinal] || notaFiscal.dataEmissao > ultimasComprasReais[chaveFinal].dataEmissao) {
                    ultimasComprasReais[chaveFinal] = {
                        ...itemNF,
                        dataEmissao: notaFiscal.dataEmissao,
                        numeroNF: notaFiscal.numeroNF
                    };
                }
            }
        });
        Logger.log(`Mapa de últimas compras REAIS (ultimasComprasReais) preenchido com ${Object.keys(ultimasComprasReais).length} entradas.`);

        // 5. Montagem do relatório final
        const relatorioFinal = [];
        const seisMesesAtras = new Date();
        seisMesesAtras.setMonth(seisMesesAtras.getMonth() - 6);

        for (const produto of produtosDaCotacao) {
            // Lógica dos cards superiores (histórico de cotações)
            const historicoProduto = dadosCotacoes.filter(l => l["Produto"] === produto && parseFloat(String(l["Comprar"] || "0").replace(",", ".")) > 0).map(l => ({data: new Date(l["Data Abertura"]),preco: parseFloat(String(l["Preço"]).replace(",", ".")) || 0,quantidade: parseFloat(String(l["Comprar"]).replace(",", ".")) || 0,precoPorFator: parseFloat(String(l["Preço por Fator"]).replace(",", ".")) || 0,un: l["UN"] || 'N/A'})).filter(item => !isNaN(item.data.getTime())).sort((a, b) => b.data - a.data);
            const unidadeDoProduto = historicoProduto.length > 0 ? historicoProduto[0].un : (dadosCotacoes.find(l => l["Produto"] === produto) || {}).UN || 'N/A';
            const ultimos6Pedidos = historicoProduto.slice(0, 6).map(p => ({ data: p.data.toLocaleDateString('pt-BR'), preco: p.preco, precoPorFator: p.precoPorFator }));
            const historico6Meses = historicoProduto.filter(p => p.data >= seisMesesAtras);
            let precoFatorMin6M = null, precoFatorMax6M = null, precoFatorMedioPonderado6M = null;
            if (historico6Meses.length > 0) {
                const somaQuantidade = historico6Meses.reduce((acc, p) => acc + p.quantidade, 0);
                precoFatorMin6M = Math.min(...historico6Meses.map(p => p.precoPorFator));
                precoFatorMax6M = Math.max(...historico6Meses.map(p => p.precoPorFator));
                const somaValorTotalFator = historico6Meses.reduce((acc, p) => acc + (p.precoPorFator * p.quantidade), 0);
                precoFatorMedioPonderado6M = somaQuantidade > 0 ? somaValorTotalFator / somaQuantidade : 0;
            }
            const volumesPorPedido = historicoProduto.map(p => ({ data: p.data.toLocaleDateString('pt-BR'), quantidade: p.quantidade, un: p.un }));
            let intervaloMedioDias = null;
            if (historicoProduto.length > 1) {
                let somaDiferencasDias = 0;
                for (let i = 0; i < historicoProduto.length - 1; i++) {
                    somaDiferencasDias += (historicoProduto[i].data - historicoProduto[i+1].data) / (1000 * 60 * 60 * 24);
                }
                intervaloMedioDias = somaDiferencasDias / (historicoProduto.length - 1);
            }

            // LÓGICA DE ANÁLISE TRIBUTÁRIA (MODIFICADA)
            const analiseTributaria = [];
            for (const chaveCompra of Object.keys(ultimasComprasReais)) {
                const [produtoCompra, fornecedorNF] = chaveCompra.split('|');

                if (produtoCompra === produto.trim()) {
                    const ultimaCompra = ultimasComprasReais[chaveCompra];
                    
                    const qtdComercial = parseFloat(String(ultimaCompra["Quantidade Comercial"]).replace(",", ".")) || 1;
                    const vlrUnitario = parseFloat(String(ultimaCompra["Valor Unitário Comercial"]).replace(",", ".")) || 0;
                    
                    // Pega o valor TOTAL do imposto para o item da nota
                    const icmsTotal = parseFloat(String(ultimaCompra["Valor (ICMS)"] || "0").replace(",", ".")) || 0;
                    const icmsStTotal = parseFloat(String(ultimaCompra["Valor (ICMS ST)"] || "0").replace(",", ".")) || 0;
                    const ipiTotal = parseFloat(String(ultimaCompra["Valor (IPI)"] || "0").replace(",", ".")) || 0;
                    const pisTotal = parseFloat(String(ultimaCompra["Valor (PIS)"] || "0").replace(",", ".")) || 0;
                    const cofinsTotal = parseFloat(String(ultimaCompra["Valor (COFINS)"] || "0").replace(",", ".")) || 0;
                    
                    // Calcula o valor UNITÁRIO de cada imposto
                    const icmsUnitario = qtdComercial > 0 ? icmsTotal / qtdComercial : 0;
                    const icmsStUnitario = qtdComercial > 0 ? icmsStTotal / qtdComercial : 0;
                    const ipiUnitario = qtdComercial > 0 ? ipiTotal / qtdComercial : 0;
                    const pisUnitario = qtdComercial > 0 ? pisTotal / qtdComercial : 0;
                    const cofinsUnitario = qtdComercial > 0 ? cofinsTotal / qtdComercial : 0;

                    // Cria o objeto de impostos com valores unitários e seus percentuais
                    const impostos = {
                        icms: icmsUnitario,
                        icms_percent: vlrUnitario > 0 ? (icmsUnitario / vlrUnitario) * 100 : 0,
                        icms_st: icmsStUnitario,
                        icms_st_percent: vlrUnitario > 0 ? (icmsStUnitario / vlrUnitario) * 100 : 0,
                        ipi: ipiUnitario,
                        ipi_percent: vlrUnitario > 0 ? (ipiUnitario / vlrUnitario) * 100 : 0,
                        pis: pisUnitario,
                        pis_percent: vlrUnitario > 0 ? (pisUnitario / vlrUnitario) * 100 : 0,
                        cofins: cofinsUnitario,
                        cofins_percent: vlrUnitario > 0 ? (cofinsUnitario / vlrUnitario) * 100 : 0,
                    };

                    const custoTotalNF = vlrUnitario + icmsUnitario + icmsStUnitario + ipiUnitario + pisUnitario + cofinsUnitario;

                    analiseTributaria.push({
                        fornecedor: fornecedorNF,
                        numeroNF: ultimaCompra.numeroNF,
                        dataUltimaNF: ultimaCompra.dataEmissao.toLocaleDateString('pt-BR'),
                        precoUnitarioNF: vlrUnitario,
                        impostos: impostos,
                        custoTotalNF: custoTotalNF
                    });
                }
            }
            Logger.log(`Para o produto '${produto}', foram listadas ${analiseTributaria.length} últimas compras reais.`);

            relatorioFinal.push({
                produto: produto,
                unidade: unidadeDoProduto,
                ultimos6Pedidos: ultimos6Pedidos,
                precoFatorMin6M: precoFatorMin6M,
                precoFatorMax6M: precoFatorMax6M,
                precoFatorMedioPonderado6M: precoFatorMedioPonderado6M,
                volumesPorPedido: volumesPorPedido,
                intervaloMedioDias: intervaloMedioDias,
                analiseTributaria: analiseTributaria.sort((a, b) => a.fornecedor.localeCompare(b.fornecedor))
            });
        }
        
        return relatorioFinal;

    } catch (error) {
        Logger.log(`ERRO CRÍTICO em RelatoriosCRUD_gerarDadosAnaliseCompra: ${error.toString()} \nStack: ${error.stack}`);
        return null;
    }
}