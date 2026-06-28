/**
 * Synchronisation Excel (Drive) -> Google Sheet, DANS Google (Apps Script).
 *
 * Copie depuis Action_2026-c_New.xlsx (sur ton Drive) :
 *   - onglet "Portefeuille BNC" : colonnes A–H
 *   - onglet "Prospects"        : colonnes A–C
 * vers les onglets de MÊME nom dans ce Google Sheet. Les autres colonnes
 * (I–P, etc.) ne sont JAMAIS touchées.
 *
 * INSTALLATION (une seule fois) :
 *   1. Ouvre le Google Sheet -> menu Extensions -> Apps Script.
 *   2. Colle TOUT ce fichier dans l'éditeur (remplace le contenu par défaut).
 *   3. Active le service avancé Drive : dans l'éditeur, panneau de gauche
 *      « Services » -> + -> « Drive API » -> Ajouter  (identifiant : Drive).
 *   4. Enregistre (icône disquette).
 *   5. Lance une fois la fonction « syncDepuisXlsx » (bouton Exécuter) et
 *      autorise l'accès quand Google le demande.
 *   6. Recharge le Google Sheet : un menu « 🔄 Sync » apparaît.
 */

// ======================== CONFIGURATION ========================
var NOM_XLSX = 'Action_2026-c_New.xlsx';
// Onglet -> nombre de colonnes à synchroniser (depuis A).
var SYNC = {
  'Portefeuille BNC': 8,  // A–H
  'Prospects': 3          // A–C
};
// ===============================================================


/** Menu personnalisé à l'ouverture du Sheet. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔄 Sync')
    .addItem('Importer depuis Excel', 'syncDepuisXlsx')
    .addToUi();
}


/** Synchronisation principale. */
function syncDepuisXlsx() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) Trouver le .xlsx sur le Drive, par nom
  var it = DriveApp.getFilesByName(NOM_XLSX);
  if (!it.hasNext()) {
    SpreadsheetApp.getUi().alert('Fichier introuvable sur le Drive : ' + NOM_XLSX);
    return;
  }
  var xlsxFile = it.next();

  // 2) Convertir le .xlsx en Google Sheet TEMPORAIRE (service avancé Drive, v2)
  var temp = Drive.Files.insert(
    { title: '__temp_import_bnc', mimeType: MimeType.GOOGLE_SHEETS },
    xlsxFile.getBlob()
  );
  var tempSs = SpreadsheetApp.openById(temp.id);

  var resume = [];
  try {
    // 3) Copier les colonnes voulues, onglet par onglet
    Object.keys(SYNC).forEach(function (nomFeuille) {
      var nbCols = SYNC[nomFeuille];
      var src = tempSs.getSheetByName(nomFeuille);
      var dst = ss.getSheetByName(nomFeuille);
      if (!src) { resume.push('⚠ Onglet absent dans le xlsx : ' + nomFeuille); return; }
      if (!dst) { resume.push('⚠ Onglet absent dans le Sheet : ' + nomFeuille); return; }

      var nbLignes = src.getLastRow();
      if (nbLignes < 1) { resume.push('— ' + nomFeuille + ' : vide'); return; }

      var valeurs = src.getRange(1, 1, nbLignes, nbCols).getValues();
      dst.getRange(1, 1, nbLignes, nbCols).setValues(valeurs);
      resume.push('✓ ' + nomFeuille + ' : ' + nbLignes + ' lignes × ' + nbCols + ' col.');
    });
  } finally {
    // 4) Toujours supprimer le fichier temporaire
    DriveApp.getFileById(temp.id).setTrashed(true);
  }

  ss.toast(resume.join('\n'), 'Synchronisation', 8);
}


/**
 * (Optionnel) Crée un déclencheur quotidien à ~6h pour synchroniser
 * automatiquement. Lance cette fonction UNE fois pour l'activer.
 */
function creerDeclencheurQuotidien() {
  ScriptApp.newTrigger('syncDepuisXlsx')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
}
