// FornecedoresController.gs

// Define as colunas que queremos exibir na tabela principal
const FornecedoresController_COLUNAS_EXIBICAO_NOMES = [
  "Fornecedor",
  "CNPJ",
  "Vendedor",
  "Telefone",
  "Email"
];

// Funções de normalização para verificação de duplicidade
function FornecedoresController_normalizarTextoComparacao(texto) {
  if (!texto || typeof texto !== 'string') return "";
  return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
}

function FornecedoresController_normalizarCnpjComparacao(cnpj) {
  if (!cnpj || typeof cnpj !== 'string') return "";
  return cnpj.replace(/\D/g, ''); // Remove todos os não dígitos
}

/**
 * Obtém os dados dos fornecedores de forma paginada e com capacidade de busca.
 * @param {object} options Objeto com { pagina: number, itensPorPagina: number, termoBusca: string|null }.
 * @return {object} { cabecalhosParaExibicao: Array<string>, fornecedoresPaginados: Array<Object>, totalItens: number, paginaAtual: number, totalPaginas: number, error?: boolean, message?: string }
 */
function FornecedoresController_obterDadosCompletosFornecedores(options) {
  try {
    const pagina = (options && options.pagina) ? parseInt(options.pagina, 10) : 1;
    const itensPorPagina = (options && options.itensPorPagina) ? parseInt(options.itensPorPagina, 10) : 5;
    const termoBusca = (options && options.termoBusca && typeof options.termoBusca === 'string') ? 
                       FornecedoresController_normalizarTextoComparacao(options.termoBusca) : null;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Planilha ativa não acessível.");

    const abaFornecedores = ss.getSheetByName(ABA_FORNECEDORES);
    if (!abaFornecedores) throw new Error(`Aba '${ABA_FORNECEDORES}' não encontrada.`);

    const range = abaFornecedores.getDataRange();
    const todasAsLinhas = range.getValues();

    if (todasAsLinhas.length <= 1) {
      return { 
        cabecalhosParaExibicao: FornecedoresController_COLUNAS_EXIBICAO_NOMES, 
        fornecedoresPaginados: [],
        totalItens: 0,
        paginaAtual: 1,
        totalPaginas: 0
      };
    }

    const cabecalhosDaPlanilha = todasAsLinhas[0].map(String);
    let dadosBrutosDaPlanilha = todasAsLinhas.slice(1);

    let todosFornecedoresObj = dadosBrutosDaPlanilha.map(linha => {
      const fornecedorObj = {};
      cabecalhosDaPlanilha.forEach((cabecalho, index) => {
        let valorCelula = linha[index];
        if (valorCelula instanceof Date) {
          if (cabecalho === CABECALHOS_FORNECEDORES[0]) { // "Data de Cadastro"
            valorCelula = Utilities.formatDate(valorCelula, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
          } else {
            valorCelula = Utilities.formatDate(valorCelula, Session.getScriptTimeZone(), "dd/MM/yyyy");
          }
        }
        fornecedorObj[cabecalho] = valorCelula !== null && valorCelula !== undefined ? String(valorCelula) : "";
      });
      return fornecedorObj;
    });

    if (termoBusca && termoBusca.length > 0) {
      todosFornecedoresObj = todosFornecedoresObj.filter(fornecedor => {
        return Object.values(fornecedor).some(valor => 
          FornecedoresController_normalizarTextoComparacao(String(valor)).includes(termoBusca)
        );
      });
    }
    
    const totalItens = todosFornecedoresObj.length;
    const totalPaginas = Math.ceil(totalItens / itensPorPagina);
    const offset = (pagina - 1) * itensPorPagina;
    const fornecedoresPaginados = todosFornecedoresObj.slice(offset, offset + itensPorPagina);

    return { 
      cabecalhosParaExibicao: FornecedoresController_COLUNAS_EXIBICAO_NOMES, 
      fornecedoresPaginados: fornecedoresPaginados,
      totalItens: totalItens,
      paginaAtual: pagina,
      totalPaginas: totalPaginas
    };

  } catch (e) {
    console.error("ERRO em FornecedoresController_obterDadosCompletosFornecedores: " + e.toString() + " Stack: " + e.stack);
    return { error: true, message: "Falha ao buscar dados dos fornecedores. Detalhes: " + e.message };
  }
}

/**
* Cria um novo fornecedor na planilha.
* @param {object} dadosNovoFornecedor Objeto contendo os dados do novo fornecedor (sem ID).
* @return {object} { success: boolean, message: string, novoId?: string }
*/
function FornecedoresController_criarNovoFornecedor(dadosNovoFornecedor) {
  try {
    console.log("FornecedoresController: Iniciando criarNovoFornecedor com dados:", JSON.stringify(dadosNovoFornecedor));
    if (!dadosNovoFornecedor || !dadosNovoFornecedor[CABECALHOS_FORNECEDORES[2]]) { // "Fornecedor"
      throw new Error("Nome do Fornecedor é obrigatório.");
    }

    const nomeNovoFornecedor = dadosNovoFornecedor[CABECALHOS_FORNECEDORES[2]];
    const cnpjNovoFornecedor = dadosNovoFornecedor[CABECALHOS_FORNECEDORES[3]]; // "CNPJ"

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Planilha ativa não acessível.");

    const abaFornecedores = ss.getSheetByName(ABA_FORNECEDORES);
    if (!abaFornecedores) throw new Error(`Aba '${ABA_FORNECEDORES}' não encontrada.`);

    const range = abaFornecedores.getDataRange();
    const todasAsLinhas = range.getValues();
    const cabecalhosDaPlanilha = todasAsLinhas[0].map(String);

    const idColunaIndex = cabecalhosDaPlanilha.indexOf(CABECALHOS_FORNECEDORES[1]); 
    const nomeFornecedorColunaIndex = cabecalhosDaPlanilha.indexOf(CABECALHOS_FORNECEDORES[2]);
    const cnpjColunaIndex = cabecalhosDaPlanilha.indexOf(CABECALHOS_FORNECEDORES[3]);

    if (idColunaIndex === -1) throw new Error("Coluna 'ID' não encontrada na planilha. Não é possível gerar novo ID.");
    if (nomeFornecedorColunaIndex === -1) throw new Error("Coluna 'Fornecedor' não encontrada na planilha.");
    if (cnpjColunaIndex === -1) throw new Error("Coluna 'CNPJ' não encontrada na planilha.");

    for (let i = 1; i < todasAsLinhas.length; i++) {
      const nomeLinhaAtual = todasAsLinhas[i][nomeFornecedorColunaIndex];
      const cnpjLinhaAtual = todasAsLinhas[i][cnpjColunaIndex];

      if (nomeNovoFornecedor && FornecedoresController_normalizarTextoComparacao(nomeLinhaAtual) === FornecedoresController_normalizarTextoComparacao(nomeNovoFornecedor)) {
        throw new Error(`O nome de fornecedor '${nomeNovoFornecedor}' já está cadastrado.`);
      }
      if (cnpjNovoFornecedor && FornecedoresController_normalizarCnpjComparacao(cnpjLinhaAtual) === FornecedoresController_normalizarCnpjComparacao(cnpjNovoFornecedor) && FornecedoresController_normalizarCnpjComparacao(cnpjNovoFornecedor) !== "") {
        throw new Error(`O CNPJ '${cnpjNovoFornecedor}' já está cadastrado.`);
      }
    }

    let proximoId = 1;
    if (todasAsLinhas.length > 1) {
      for (let i = 1; i < todasAsLinhas.length; i++) {
        const idAtual = parseInt(todasAsLinhas[i][idColunaIndex]);
        if (!isNaN(idAtual) && idAtual >= proximoId) {
          proximoId = idAtual + 1;
        }
      }
    }
    const novoIdGerado = String(proximoId);

    const novaLinhaArray = [];
    CABECALHOS_FORNECEDORES.forEach(nomeCabecalho => {
      if (nomeCabecalho === CABECALHOS_FORNECEDORES[0]) {
        novaLinhaArray.push(new Date());
      } else if (nomeCabecalho === CABECALHOS_FORNECEDORES[1]) {
        novaLinhaArray.push(novoIdGerado);
      } else if (nomeCabecalho === CABECALHOS_FORNECEDORES[12]) { // "Pedido Mínimo (R$)"
        let valorPedido = dadosNovoFornecedor[nomeCabecalho] || "0";
        if (typeof valorPedido === 'string') {
          valorPedido = valorPedido.replace('R$', '').replace(/\./g, '').replace(',', '.').trim();
        }
        novaLinhaArray.push(parseFloat(valorPedido) || 0);
      } else {
        novaLinhaArray.push(dadosNovoFornecedor[nomeCabecalho] || "");
      }
    });

    abaFornecedores.appendRow(novaLinhaArray);
    SpreadsheetApp.flush();
    return { success: true, message: "Fornecedor criado com sucesso!", novoId: novoIdGerado };

  } catch (e) {
    console.error("ERRO em FornecedoresController_criarNovoFornecedor: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: e.message };
  }
}

/**
* Atualiza um fornecedor na planilha.
* @param {object} fornecedorParaAtualizar Objeto contendo o "ID" do fornecedor e os campos a serem atualizados.
* @return {object} { success: boolean, message: string }
*/
function FornecedoresController_atualizarFornecedor(fornecedorParaAtualizar) {
  try {
    const idParaAtualizar = String(fornecedorParaAtualizar["ID"]);
    if (!idParaAtualizar) throw new Error("ID do fornecedor é obrigatório para atualização.");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaFornecedores = ss.getSheetByName(ABA_FORNECEDORES);
    if (!abaFornecedores) throw new Error(`Aba '${ABA_FORNECEDORES}' não encontrada.`);

    const range = abaFornecedores.getDataRange();
    const todasAsLinhas = range.getValues();
    const cabPlan = todasAsLinhas[0].map(String);

    // Índices
    const idxId = cabPlan.indexOf("ID");
    const idxNome = cabPlan.indexOf("Fornecedor");
    
    // Encontrar linha e capturar nome antigo
    let linhaPlan = -1;
    for (let i = 1; i < todasAsLinhas.length; i++) {
      if (String(todasAsLinhas[i][idxId]) === idParaAtualizar) {
        linhaPlan = i + 1; // 1-based
        break;
      }
    }
    if (linhaPlan === -1) throw new Error(`Fornecedor com ID '${idParaAtualizar}' não encontrado.`);
    const nomeAntigo = String(todasAsLinhas[linhaPlan - 1][idxNome]);
    const nomeAtualizado = fornecedorParaAtualizar["Fornecedor"];
    if (!nomeAtualizado) throw new Error("Nome do Fornecedor é obrigatório.");

    // Verificar duplicidade em outros IDs
    function _norm(txt) { return txt.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim(); }
    const nomeAtualNorm = _norm(nomeAtualizado);
    for (let i = 1; i < todasAsLinhas.length; i++) {
      if ((i + 1) !== linhaPlan && _norm(String(todasAsLinhas[i][idxNome])) === nomeAtualNorm) {
        throw new Error(`O nome '${nomeAtualizado}' já está cadastrado para outro ID.`);
      }
    }

    // Montar linha final
    const novaLinha = cabPlan.map((cab, k) => {
      if (k === cabPlan.indexOf("Data de Cadastro") || k === idxId) {
        return todasAsLinhas[linhaPlan - 1][k];
      }
      return fornecedorParaAtualizar[cab] !== undefined ? fornecedorParaAtualizar[cab] : todasAsLinhas[linhaPlan - 1][k];
    });

    // Atualizar na planilha
    abaFornecedores.getRange(linhaPlan, 1, 1, novaLinha.length).setValues([novaLinha]);
    SpreadsheetApp.flush();

    // --- NOVO: Propagar novo nome para SubProdutos vinculados ---
    const abaSub = ss.getSheetByName(ABA_SUBPRODUTOS);
    if (abaSub) {
      const dadosSub = abaSub.getDataRange().getValues();
      const idxSubFor = CABECALHOS_SUBPRODUTOS.indexOf("Fornecedor");
      for (let j = 1; j < dadosSub.length; j++) {
        if (String(dadosSub[j][idxSubFor]) === nomeAntigo) {
          abaSub.getRange(j + 1, idxSubFor + 1).setValue(nomeAtualizado);
        }
      }
      SpreadsheetApp.flush();
    }

    return { success: true, message: "Fornecedor atualizado e SubProdutos propagados com sucesso!" };

  } catch (e) {
    console.error("ERRO em FornecedoresController_atualizarFornecedor: " + e.toString());
    return { success: false, message: e.message };
  }
}