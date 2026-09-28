function diagnosticar() {
  var idReinf = "ID_EXEMPLO";

  try {
    var pasta = DriveApp.getFolderById(idMit);
    Logger.log("Pasta raiz: " + pasta.getName());

    // Lista TODAS as subpastas para ver o nome exato
    var subs = pasta.getFolders();
    Logger.log("=== SUBPASTAS DENTRO DE REINF ===");
    var count = 0;
    while (subs.hasNext()) {
      var sub = subs.next();
      Logger.log("[" + count + "] nome='" + sub.getName() + "' id=" + sub.getId());
      count++;
    }
    if (count === 0) Logger.log("NENHUMA subpasta encontrada!");

  } catch(e) {
    Logger.log("ERRO: " + e);
  }
}