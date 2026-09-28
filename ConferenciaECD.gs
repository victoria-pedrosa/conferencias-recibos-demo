// --- CONFIGURAÇÕES FIXAS ---
const ID_PASTA_RAIZ_ECD = 'ID_EXEMPLO';
const COLUNA_EMPRESA = 1; // 1 = Coluna B. Se mudou para a Coluna A, coloque 0 aqui.

// 1. Menu
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🤖 Automação ECD')
    .addItem('Conferir Recibos (Linha por Linha)', 'iniciarConferenciaLinhaPorLinha')
    .addToUi();
}

// 2. Coleta o Ano
function iniciarConferenciaLinhaPorLinha() {
  var ui = SpreadsheetApp.getUi();
  
  var resposta = ui.prompt(
    'Ano Calendário da ECD',
    'Digite o ano que deseja conferir na leitura do PDF (ex: 2025):',
    ui.ButtonSet.OK_CANCEL
  );

  if (resposta.getSelectedButton() == ui.Button.OK) {
    var anoPesquisa = resposta.getResponseText().trim();
    if(anoPesquisa.length !== 4) {
      ui.alert('Erro', 'Por favor, digite um ano válido com 4 dígitos.', ui.ButtonSet.OK);
      return;
    }
    executarVarreduraLinhaPorLinha(anoPesquisa);
  }
}

// 3. Função Principal
function executarVarreduraLinhaPorLinha(anoPesquisa) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  spreadsheet.toast("Mapeando pastas no Google Drive...", "🤖 Preparando", 3);
  
  try {
    var pastaRaiz = DriveApp.getFolderById(ID_PASTA_RAIZ_ECD);
  } catch(e) {
    SpreadsheetApp.getUi().alert("Erro ao acessar a pasta mãe. Verifique as permissões.");
    return;
  }

  // Mapeia todas as pastas mantendo os espaços no nome para podermos contar as palavras
  var mapPastas = {};
  var subpastas = pastaRaiz.getFolders();
  while (subpastas.hasNext()) {
    var pasta = subpastas.next();
    mapPastas[normalizarTexto(pasta.getName())] = pasta.getId();
  }

  spreadsheet.toast("Iniciando conferência das selecionadas!", "🤖 Robô em Ação", 3);

  for (var i = 1; i < data.length; i++) { 
    var caixaMarcada = data[i][10]; // Coluna K
    
    if (caixaMarcada === true) {
      var linha = i + 1;
      var nomeEmpresa = data[i][COLUNA_EMPRESA]; 
      var cnpjPlanilha = String(data[i][2]).replace(/\D/g, ''); 
      
      sheet.getRange(linha, 10).setValue('⌛ Buscando pasta...');
      SpreadsheetApp.flush(); 
      
      var nomeLimpoPlanilha = normalizarTexto(nomeEmpresa.toString());
      
      // Pega as 3 primeiras palavras da empresa na planilha
      var tresPrimeirasPalavras = pegarPrimeirasPalavras(nomeLimpoPlanilha, 3);
      var idPastaEmpresa = null;

      // Procura qual pasta do Drive contém essas 3 palavras
      for (var nomePasta in mapPastas) {
        if (nomePasta.indexOf(tresPrimeirasPalavras) !== -1) {
          idPastaEmpresa = mapPastas[nomePasta];
          break;
        }
      }

      if (!idPastaEmpresa) {
        // Status atualizado
        sheet.getRange(linha, 10).setValue('Pasta Não Encontrada');
        sheet.getRange(linha, 11).setValue(false);
        SpreadsheetApp.flush();
        continue; 
      }

      // Acessa a pasta da empresa e procura o Recibo
      var pastaEmpresa = DriveApp.getFolderById(idPastaEmpresa);
      var arquivosRecibo = pastaEmpresa.searchFiles("title contains 'RECIBO' and mimeType = 'application/pdf'");

      if (arquivosRecibo.hasNext()) {
        sheet.getRange(linha, 10).setValue('⌛ Lendo PDF...');
        SpreadsheetApp.flush();

        var file = arquivosRecibo.next();
        var text = "";
        try {
          text = extrairTextoDoPDF(file);
        } catch(e) {
          // Status atualizado
          sheet.getRange(linha, 10).setValue('Falha ao Ler PDF');
          sheet.getRange(linha, 11).setValue(false);
          SpreadsheetApp.flush();
          continue;
        }

        // Validações de Segurança (CNPJ e Ano)
        var anoCorreto = text.indexOf(anoPesquisa) !== -1;
        var achouCnpjDaLinha = false;
        
        var regexCnpj = /\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}/g;
        var cnpjsEncontrados = text.match(regexCnpj);

        if (cnpjsEncontrados) {
          for (var c = 0; c < cnpjsEncontrados.length; c++) {
            var cnpjRecibo = cnpjsEncontrados[c].replace(/\D/g, ''); 
            if (cnpjRecibo === cnpjPlanilha) {
              achouCnpjDaLinha = true;
              break;
            }
          }
        }

        // Status atualizados conforme solicitado
        if (achouCnpjDaLinha && anoCorreto) {
          sheet.getRange(linha, 10).setValue('OK - ' + anoPesquisa);
        } else if (!achouCnpjDaLinha) {
          sheet.getRange(linha, 10).setValue('CNPJ Diferente');
        } else {
          sheet.getRange(linha, 10).setValue('Ano Incorreto');
        }

      } else {
        // Status atualizado
        sheet.getRange(linha, 10).setValue('Recibo Ausente');
      }

      sheet.getRange(linha, 11).setValue(false);
      SpreadsheetApp.flush();
    }
  }

  spreadsheet.toast("Todas as empresas selecionadas foram processadas!", "Finalizado 🎉", 5);
}

// 4. Limpeza de Texto (MATÉM OS ESPAÇOS para separar as palavras)
function normalizarTexto(texto) {
  if (!texto) return "";
  return texto.toString().toUpperCase()
              .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // Tira acentos
              .replace(/[^A-Z0-9 ]/g, ' ') // Tira caracteres especiais mas MANTÉM espaços
              .replace(/\s+/g, ' ') // Transforma espaços duplos em um só
              .trim();
}

// 5. Função: Pega as X primeiras palavras
function pegarPrimeirasPalavras(texto, quantidade) {
  if (!texto) return "";
  var palavras = texto.split(" "); // Divide a frase onde tiver espaço
  return palavras.slice(0, quantidade).join(" "); // Pega só as 3 primeiras e junta de novo
}

// 6. Função de OCR 
function extrairTextoDoPDF(file) {
  var blob = file.getBlob();
  var fileMetadata = {
    name: 'Temp_OCR_' + file.getName(),
    mimeType: 'application/vnd.google-apps.document'
  };
  var tempDoc = Drive.Files.create(fileMetadata, blob, {supportsAllDrives: true});
  var doc = DocumentApp.openById(tempDoc.id);
  var texto = doc.getBody().getText();
  Drive.Files.remove(tempDoc.id, {supportsAllDrives: true});
  return texto;
}