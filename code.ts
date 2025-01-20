function main(workbook: ExcelScript.Workbook) {
  // 1. Récupération des feuilles nécessaires
  const fcuSheet = workbook.getWorksheet("FCU"); // Feuille de destination
  const extractionSheet = workbook.getWorksheet("Extraction FCU"); // Feuille source

  // 2. Suppression des colonnes inutiles dans Extraction FCU
  harmonizeColumns(extractionSheet, fcuSheet);

  // 3. Création de la clé de concaténation dans les deux feuilles
  createConcatKey(extractionSheet, "J", "L"); // OWNER_CODE (J) + Serial number (L)
  createConcatKey(fcuSheet, "E", "F"); // OWNER_CODE (E) + Serial number (F)

  // 4. Récupération des commentaires avec une recherche V
  retrieveComments(extractionSheet, fcuSheet);

  // 5. Mise à jour de la feuille FCU avec les données de Extraction FCU
  updateFcuSheet(fcuSheet, extractionSheet);

  // 6. Nettoyage des colonnes temporaires (concatKey)
  cleanTemporaryColumns(fcuSheet, extractionSheet);
}

// Fonction 1 : Harmonisation des colonnes (suppression des colonnes inutiles)
function harmonizeColumns(extractionSheet: ExcelScript.Worksheet, fcuSheet: ExcelScript.Worksheet): void {
  const fcuColumns = fcuSheet.getRange("1:1").getTexts()[0]; // Colonnes de FCU (ligne d'entête)
  const extractionColumns = extractionSheet.getRange("1:1").getTexts()[0]; // Colonnes de Extraction FCU

  for (let i = extractionColumns.length - 1; i >= 0; i--) {
    if (!fcuColumns.includes(extractionColumns[i])) {
      extractionSheet.getRangeByIndexes(0, i, extractionSheet.getRowCount(), 1)
        .delete(ExcelScript.DeleteShiftDirection.left);
    }
  }
}

// Fonction 2 : Création de la clé de concaténation
function createConcatKey(sheet: ExcelScript.Worksheet, col1: string, col2: string): void {
  // Insérer une colonne pour la clé de concaténation
  sheet.getRange("A:A").insert(ExcelScript.InsertShiftDirection.right);
  sheet.getRange("A1").setValue("concatKey");

  // Formule de concaténation
  sheet.getRange("A2").setFormulaLocal(`=CONCATENER(${col1}2;${col2}2)`);

  // Autofill jusqu'à la dernière ligne
  const lastRow = sheet.getUsedRange().getLastRow().getRowIndex();
  sheet.getRange("A2").autoFill(`A2:A${lastRow + 1}`, ExcelScript.AutoFillType.fillFormulas);
}

// Fonction 3 : Récupération des commentaires (ou autres données)
function retrieveComments(extractionSheet: ExcelScript.Worksheet, fcuSheet: ExcelScript.Worksheet): void {
  // Ajouter une colonne pour les commentaires rapatriés
  const newColIndex = extractionSheet.getUsedRange().getColumnCount() + 1;
  extractionSheet.getRangeByIndexes(0, newColIndex - 1, extractionSheet.getRowCount(), 1)
    .insert(ExcelScript.InsertShiftDirection.right);
  extractionSheet.getCell(0, newColIndex - 1).setValue("Comments_rapatriés");

  // Formule RECHERCHEV
  extractionSheet.getCell(1, newColIndex - 1).setFormulaLocal(
    `=RECHERCHEV(A2;FCU!A:N;14;FAUX)` // Recherche colonne Comments (14ᵉ colonne)
  );

  // Autofill pour toute la colonne
  const lastRow = extractionSheet.getUsedRange().getLastRow().getRowIndex();
  extractionSheet.getRangeByIndexes(1, newColIndex - 1, lastRow, 1)
    .autoFill(`${ExcelScript.Range.toLetter(newColIndex)}2:${ExcelScript.Range.toLetter(newColIndex)}${lastRow + 1}`, ExcelScript.AutoFillType.fillFormulas);
}

// Fonction 4 : Mise à jour de la feuille FCU
function updateFcuSheet(fcuSheet: ExcelScript.Worksheet, extractionSheet: ExcelScript.Worksheet): void {
  // Effacer les données existantes
  fcuSheet.getRange().clear(ExcelScript.ClearApplyTo.all);

  // Copier les données mises à jour de Extraction FCU
  const extractionData = extractionSheet.getUsedRange();
  fcuSheet.getRangeByIndexes(0, 0, extractionData.getRowCount(), extractionData.getColumnCount())
    .copyFrom(extractionData, ExcelScript.RangeCopyType.values, false, false);
}

// Fonction 5 : Nettoyage des colonnes temporaires
function cleanTemporaryColumns(fcuSheet: ExcelScript.Worksheet, extractionSheet: ExcelScript.Worksheet): void {
  fcuSheet.getRange("A:A").delete(ExcelScript.DeleteShiftDirection.left); // Supprimer concatKey de FCU
  extractionSheet.getRange("A:A").delete(ExcelScript.DeleteShiftDirection.left); // Supprimer concatKey de Extraction FCU
}
