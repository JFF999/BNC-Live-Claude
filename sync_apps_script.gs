/**
 * Synchronisation Excel (Drive) -> Google Sheet, DANS Google (Apps Script).
 *
 * Copie depuis Action_2026-c_New.xlsx (sur ton Drive) :
 *   - onglet "Portefeuille BNC" : colonnes A–J (source BNC + Pré Aff + MAJ Aff)
 *   - onglet "Prospects"        : colonnes A–E (source + Pré Aff + MAJ Aff)
 * vers les onglets de MÊME nom dans ce Google Sheet. Les colonnes calculées
 * par l'app (Prix $, Pré G %, Pré YF, MAJ YF, …) ne sont JAMAIS touchées.
 *
 * INSTALLATION (une seule fois) :
 *   1. Ouvre le Google Sheet -> menu Extensions -> Apps Script.
 *   2. Colle TOUT ce fichier dans l'éditeur (remplace le contenu par défaut).
 *   3. Active le service avancé Drive : panneau de gauche « Services » -> +
 *      -> « Drive API » -> Ajouter. (v2 ou v3 : le script gère les deux.)
 *   4. Enregistre (icône disquette).
 *   5. Lance une fois la fonction « syncDepuisXlsx » (bouton Exécuter) et
 *      autorise l'accès quand Google le demande.
 *   6. Recharge le Google Sheet : un menu « 🔄 Sync » apparaît.
 */

// ======================== CONFIGURATION ========================
var NOM_XLSX = 'Action_2026-c_New.xlsx';
// Onglet -> nombre de colonnes à synchroniser (depuis A).
var SYNC = {
  'Portefeuille BNC': 10,  // A–J
  'Prospects': 5           // A–E
};
// ===============================================================


/** Menu personnalisé à l'ouverture du Sheet. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔄 Sync')
    .addItem('Importer depuis Excel', 'syncDepuisXlsx')
    .addToUi();
}


/** Convertit un .xlsx (Blob) en Google Sheet temporaire et renvoie son id.
 *  Compatible service avancé Drive v3 (Files.create) ET v2 (Files.insert). */
function xlsxVersSheetTemporaire(blob) {
  var ressource = { name: '__temp_import_bnc', mimeType: MimeType.GOOGLE_SHEETS };
  if (Drive.Files.create) {            // Drive API v3 (par défaut aujourd'hui)
    return Drive.Files.create(ressource, blob).id;
  }
  // Drive API v2 (ancien) : le champ s'appelle "title"
  return Drive.Files.insert({ title: ressource.name, mimeType: ressource.mimeType }, blob).id;
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

  // 2) Convertir le .xlsx en Google Sheet TEMPORAIRE
  var tempId = xlsxVersSheetTemporaire(xlsxFile.getBlob());
  var tempSs = SpreadsheetApp.openById(tempId);

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
    DriveApp.getFileById(tempId).setTrashed(true);
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
