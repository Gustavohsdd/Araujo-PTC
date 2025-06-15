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
    if (!fornecedorParaAtualizar || !fornecedorParaAtualizar["ID"]) {
      throw new Error("ID do fornecedor é obrigatório para atualização.");
    }

    const idParaAtualizar = String(fornecedorParaAtualizar["ID"]);
    const nomeAtualizado = fornecedorParaAtualizar[CABECALHOS_FORNECEDORES[2]];
    const cnpjAtualizado = fornecedorParaAtualizar[CABECALHOS_FORNECEDORES[3]];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaFornecedores = ss.getSheetByName(ABA_FORNECEDORES);
    if (!abaFornecedores) throw new Error(`Aba '${ABA_FORNECEDORES}' não encontrada.`);

    const range = abaFornecedores.getDataRange();
    const todasAsLinhas = range.getValues();
    const cabecalhosDaPlanilha = todasAsLinhas[0].map(String);

    const idColunaIndex = cabecalhosDaPlanilha.indexOf(CABECALHOS_FORNECEDORES[1]);
    const nomeFornecedorColunaIndex = cabecalhosDaPlanilha.indexOf(CABECALHOS_FORNECEDORES[2]);
    const cnpjColunaIndex = cabecalhosDaPlanilha.indexOf(CABECALHOS_FORNECEDORES[3]);

    if (idColunaIndex === -1) throw new Error("Coluna 'ID' não encontrada.");
    if (nomeFornecedorColunaIndex === -1) throw new Error("Coluna 'Fornecedor' não encontrada.");
    if (cnpjColunaIndex === -1) throw new Error("Coluna 'CNPJ' não encontrada.");

    let linhaParaAtualizarIndexNaPlanilha = -1;

    for (let i = 1; i < todasAsLinhas.length; i++) {
      const idLinhaAtual = String(todasAsLinhas[i][idColunaIndex]);
      if (idLinhaAtual === idParaAtualizar) {
        linhaParaAtualizarIndexNaPlanilha = i + 1;
      } else {
        const nomeLinhaAtual = todasAsLinhas[i][nomeFornecedorColunaIndex];
        const cnpjLinhaAtual = todasAsLinhas[i][cnpjColunaIndex];
        if (nomeAtualizado && FornecedoresController_normalizarTextoComparacao(nomeLinhaAtual) === FornecedoresController_normalizarTextoComparacao(nomeAtualizado)) {
          throw new Error(`O nome de fornecedor '${nomeAtualizado}' já está cadastrado para outro ID (${idLinhaAtual}).`);
        }
        if (cnpjAtualizado && FornecedoresController_normalizarCnpjComparacao(cnpjLinhaAtual) === FornecedoresController_normalizarCnpjComparacao(cnpjAtualizado) && FornecedoresController_normalizarCnpjComparacao(cnpjAtualizado) !== "") {
          throw new Error(`O CNPJ '${cnpjAtualizado}' já está cadastrado para outro ID (${idLinhaAtual}).`);
        }
      }
    }

    if (linhaParaAtualizarIndexNaPlanilha === -1) {
      throw new Error(`Fornecedor com ID '${idParaAtualizar}' não encontrado para atualização.`);
    }

    const valorOriginalLinha = todasAsLinhas[linhaParaAtualizarIndexNaPlanilha - 1];
    const linhaFinalParaSalvar = [];
    let alteracoesReais = 0;

    for (let k = 0; k < cabecalhosDaPlanilha.length; k++) {
      const nomeCabecalhoPlanilha = cabecalhosDaPlanilha[k];
      let valorOriginalCelula = valorOriginalLinha[k];

      if (nomeCabecalhoPlanilha === CABECALHOS_FORNECEDORES[0]) { // "Data de Cadastro"
        linhaFinalParaSalvar.push(valorOriginalCelula instanceof Date ? valorOriginalCelula : new Date(valorOriginalCelula)); // Mantém a data original se válida
      } else if (nomeCabecalhoPlanilha === CABECALHOS_FORNECEDORES[1]) { // "ID"
        linhaFinalParaSalvar.push(idParaAtualizar);
      } else if (fornecedorParaAtualizar.hasOwnProperty(nomeCabecalhoPlanilha)) {
        let valorDoForm = fornecedorParaAtualizar[nomeCabecalhoPlanilha];
        if (nomeCabecalhoPlanilha === CABECALHOS_FORNECEDORES[12]) { // "Pedido Mínimo (R$)"
          if (typeof valorDoForm === 'string') {
            valorDoForm = valorDoForm.replace('R$', '').replace(/\./g, '').replace(',', '.').trim();
          }
          valorDoForm = parseFloat(valorDoForm);
          if (isNaN(valorDoForm)) valorDoForm = valorOriginalCelula;
        }
        linhaFinalParaSalvar.push(valorDoForm);
        
        let originalComparavel = valorOriginalCelula;
        let formComparavel = valorDoForm;
        if (valorOriginalCelula instanceof Date && nomeCabecalhoPlanilha !== CABECALHOS_FORNECEDORES[0]) { // Para outras datas
             originalComparavel = Utilities.formatDate(valorOriginalCelula, Session.getScriptTimeZone(), "dd/MM/yyyy");
             formComparavel = valorDoForm; // Assumindo que do form vem como string "dd/MM/yyyy" se for data
        } else if (nomeCabecalhoPlanilha === CABECALHOS_FORNECEDORES[12]) {
            originalComparavel = parseFloat(String(valorOriginalCelula).replace('R$', '').replace(/\./g, '').replace(',', '.').trim());
            if(isNaN(originalComparavel)) originalComparavel = 0;
            formComparavel = parseFloat(String(valorDoForm)); // valorDoForm já é float aqui
            if(isNaN(formComparavel)) formComparavel = 0;
        }

        if (String(originalComparavel) !== String(formComparavel)) {
          if (nomeCabecalhoPlanilha === CABECALHOS_FORNECEDORES[12] && (originalComparavel === 0 || isNaN(originalComparavel)) && formComparavel === 0) {
            // Não conta como alteração se ambos são zero para pedido mínimo
          } else {
            alteracoesReais++;
          }
        }
      } else {
        linhaFinalParaSalvar.push(valorOriginalCelula);
      }
    }

    if (alteracoesReais > 0) {
      abaFornecedores.getRange(linhaParaAtualizarIndexNaPlanilha, 1, 1, linhaFinalParaSalvar.length).setValues([linhaFinalParaSalvar]);
      SpreadsheetApp.flush();
      return { success: true, message: "Fornecedor atualizado com sucesso!" };
    } else {
      return { success: true, message: "Nenhum dado foi modificado." };
    }

  } catch (e) {
    console.error("ERRO em FornecedoresController_atualizarFornecedor: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: e.message };
  }
}