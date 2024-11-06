# code-exp-
function main(workbook: ExcelScript.Workbook) {
    // Initialisation des feuilles de travail
    let extractionSheet = workbook.getWorksheet("Extraction MOTMOD");
    let motmodSheet = workbook.getWorksheet("MOTMOD");

    // Processus complet de mise à jour des données
    processData(extractionSheet, motmodSheet);
}

function processData(extractionSheet: ExcelScript.Worksheet, motmodSheet: ExcelScript.Worksheet) {
    let lastRow = extractionSheet.getUsedRange().getLastRow().getRowIndex();

    // Préparation de la feuille d'extraction
    prepareExtractionSheet(extractionSheet, lastRow);

    // Mise à jour des données vers la feuille MOTMOD
    updateDataToMOTMOD(extractionSheet, motmodSheet, lastRow);

    // Mise en forme de la feuille MOTMOD
    formatMOTMODSheet(motmodSheet, lastRow);
}

function prepareExtractionSheet(extractionSheet: ExcelScript.Worksheet, lastRow: number) {
    // Réglages initiaux pour la hauteur des lignes
    extractionSheet.getRange("A:XFD").getFormat().setRowHeight(15.75);

    // Suppression et réorganisation des colonnes
    manageColumns(extractionSheet);

    // Insertion et gestion de clés pour la recherche
    manageSearchKeys(extractionSheet, lastRow);
}

function manageColumns(sheet: ExcelScript.Worksheet) {
    // Suppression des colonnes inutiles
    sheet.getRange("B:C").delete(ExcelScript.DeleteShiftDirection.left);

    // Réorganisation de la colonne 'Amount'
    let amountRange = sheet.getRange("H:H");
    amountRange.insert(ExcelScript.InsertShiftDirection.right);
    amountRange.copyFrom(sheet.getRange("K:K"), ExcelScript.RangeCopyType.all, false, false);
    sheet.getRange("K:K").delete(ExcelScript.DeleteShiftDirection.left);

    // Ajout d'une colonne de commentaires
    sheet.getRange("I:I").insert(ExcelScript.InsertShiftDirection.right);
    sheet.getRange("I1").setValue("commentaires");

    // Suppression des colonnes supplémentaires
    sheet.getRange("O:V").delete(ExcelScript.DeleteShiftDirection.left);

    // Ajustement automatique des colonnes restantes
    sheet.getRange("A:N").getFormat().autofitColumns();
}

function manageSearchKeys(sheet: ExcelScript.Worksheet, lastRow: number) {
    // Ajout de clés de recherche
    sheet.getRange("A:A").insert(ExcelScript.InsertShiftDirection.right);
    sheet.getRange("A1").setFormulaLocal("concat");
    sheet.getRange("A2").setFormulaLocal("=concatener(E2;H2)");
    sheet.getRange("A2:A" + lastRow).autoFill();
}

function updateDataToMOTMOD(extractionSheet: ExcelScript.Worksheet, motmodSheet: ExcelScript.Worksheet, lastRow: number) {
    // Copie des données depuis Extraction MOTMOD vers MOTMOD
    motmodSheet.getRange("A:N").copyFrom(extractionSheet.getRange("B:O"), ExcelScript.RangeCopyType.values, false, false);
    // Nettoyage après copie
    extractionSheet.getRange().clear(ExcelScript.ClearApplyTo.all);
}

function formatMOTMODSheet(motmodSheet: ExcelScript.Worksheet, lastRow: number) {
    // Application des styles de bordure et de couleur
    let range = motmodSheet.getRange("A1:N" + lastRow);
    applyBorders(range);
    applyStyles(range);
}

function applyBorders(range: ExcelScript.Range) {
    // Configuration des bordures internes et externes
    const borderStyle = ExcelScript.BorderLineStyle.continuous;
    const borderColor = "000000";
    const borderWeight = ExcelScript.BorderWeight.thin;

    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(borderStyle);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(borderStyle);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(borderStyle);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(borderStyle);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(borderStyle);
    range.getFormat().getRangeBorder(Excel
