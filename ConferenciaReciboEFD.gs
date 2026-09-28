// --- 1. CONFIGURAÇÕES FIXAS ---
// Este é o ID da pasta mãe "EFD"
const ID_PASTA_RAIZ_EFD = 'ID_EXEMPLO';

function atualizarStatusTransmissaoEFD() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = spreadsheet.getActiveSheet();
  
  // O nome da aba TEM que ser igual ao nome da pasta (ex: "2026-02")
  const nomeAba = sheet.getName(); 
  
  // Colunas da sua planilha (A = 0, K = 10)
  const COLUNA_EMPRESA = 0; 
  const COLUNA_STATUS = 10; 
  
  // Mostra um aviso no canto da tela da planilha
  spreadsheet.toast("Acessando a pasta raiz EFD...", "🤖 Robô EFD", 3);
  
  let pastaRaiz;
  try {
    pastaRaiz = DriveApp.getFolderById(ID_PASTA_RAIZ_EFD);
  } catch (e) {
    ui.alert("❌ Erro ao acessar a pasta mãe EFD.\nVerifique se o ID está correto ou se você tem permissão.\nErro: " + e.message);
    return;
  }

  spreadsheet.toast(`Procurando a pasta da competência: ${nomeAba}...`, "🤖 Robô EFD", 3);
  // Busca a pasta que tem o exato nome da aba dentro da pasta EFD
  const pastasEncontradas = pastaRaiz.searchFolders(`title = '${nomeAba}'`);
  
  if (!pastasEncontradas.hasNext()) {
    ui.alert(`❌ Subpasta não encontrada!\n\nNão encontrei nenhuma pasta chamada "${nomeAba}" dentro da pasta EFD principal.`);
    return;
  }
  
  // Pega a primeira pasta que ele achou com esse nome
  const pastaCompetencia = pastasEncontradas.next(); 
  const idPasta = pastaCompetencia.getId();

  const dados = sheet.getDataRange().getValues();
  
  spreadsheet.toast("Lendo arquivos PDF nas subpastas... Isso pode levar alguns segundos.", "🤖 Robô EFD", 5);
  let todosOsPdfs = obterTodosOsPdfs(idPasta);
  
  if (todosOsPdfs.length === 0) {
    ui.alert(`⚠️ Nenhum PDF encontrado dentro da pasta "${nomeAba}" ou em suas subpastas.`);
    return;
  }
  
  // Criar um "mapa" rápido com os nomes das empresas
  let mapaEmpresas = {};
  for (let i = 1; i < dados.length; i++) { // Começa do 1 para pular o cabeçalho
    let nomeEmpresa = dados[i][COLUNA_EMPRESA];
    if (nomeEmpresa) {
      let nomeLimpo = normalizarTexto(nomeEmpresa.toString());
      mapaEmpresas[nomeLimpo] = i + 1; 
    }
  }

  spreadsheet.toast("Tudo lido! Atualizando a sua planilha agora...", "🤖 Robô EFD", 5);
  let linhasAtualizadas = 0;

  // Passar por cada PDF encontrado e procurar na planilha
  for (let i = 0; i < todosOsPdfs.length; i++) {
    let arquivoPdf = todosOsPdfs[i];
    let nomeArquivo = arquivoPdf.getName();
    
    // Tira o ".pdf" do final e limpa o texto para comparar
    let nomeArquivoLimpo = normalizarTexto(nomeArquivo.replace(/\.pdf$/i, ''));
    
    let linhaEncontrada = mapaEmpresas[nomeArquivoLimpo];
    
    // Busca inteligente: caso o nome do arquivo tenha sido cortado
    if (!linhaEncontrada) {
      for (let chaveEmpresa in mapaEmpresas) {
        if (chaveEmpresa.startsWith(nomeArquivoLimpo) || nomeArquivoLimpo.startsWith(chaveEmpresa)) {
          linhaEncontrada = mapaEmpresas[chaveEmpresa];
          break;
        }
      }
    }
    
    // Se achou a empresa na planilha, escreve "Transmitido"
    if (linhaEncontrada) {
      sheet.getRange(linhaEncontrada, COLUNA_STATUS + 1).setValue("Transmitido");
      linhasAtualizadas++;
    }
  }
  
  spreadsheet.toast("Finalizado!", "🤖 Robô EFD", 2);
  ui.alert(`✅ Processo Concluído!\n\nAba verificada: ${nomeAba}\nEncontrados ${todosOsPdfs.length} PDFs.\n${linhasAtualizadas} empresas marcadas como 'Transmitido'.`);
}

// --- 2. FUNÇÃO QUE LIMPA OS NOMES PARA COMPARAÇÃO ---
function normalizarTexto(texto) {
  return texto.toUpperCase()
              .replace(/[^A-Z0-9]/g, '') 
              .trim();
}

// --- 3. FUNÇÕES DE BUSCA NAS SUBPASTAS (RECURSIVIDADE) ---
function obterTodosOsPdfs(idPasta) {
  const pastaPrincipal = DriveApp.getFolderById(idPasta);
  let todosOsPdfs = [];
  buscarPdfsNaPasta(pastaPrincipal, todosOsPdfs);
  return todosOsPdfs;
}

function buscarPdfsNaPasta(pasta, arrayDeArquivos) {
  const arquivos = pasta.getFilesByType(MimeType.PDF);
  while (arquivos.hasNext()) {
    arrayDeArquivos.push(arquivos.next());
  }
  const subpastas = pasta.getFolders();
  while (subpastas.hasNext()) {
    const subpasta = subpastas.next();
    buscarPdfsNaPasta(subpasta, arrayDeArquivos); 
  }
}

// --- 4. FUNÇÕES DE MENU ---
function limparCNPJ(cnpj) {
  return cnpj.replace(/[^\d]/g, '');
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 Automação Fiscal')
    .addItem('Verificar Recibos EFD', 'atualizarStatusTransmissaoEFD')
    .addToUi();
}