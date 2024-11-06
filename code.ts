function main(workbook: ExcelScript.Workbook) {
    // Obtention des feuilles nécessaires
    let extractionSheet = workbook.getWorksheet("Extraction MOTMOD");
    let motmodSheet = workbook.getWorksheet("MOTMOD");

    // Processus de mise à jour des données
    updateDataProcess(extractionSheet, motmodSheet);
}

function updateDataProcess(extractionSheet: ExcelScript.Worksheet, motmodSheet: ExcelScript.Worksheet) {
    let lastRow = extractionSheet.getUsedRange().getLastRow().getRowIndex();

    // Préparation et nettoyage de la feuille d'extraction
    prepareExtractionSheet(extractionSheet);

    // Transfert et mise à jour des données
    transferData(extractionSheet, motmodSheet, lastRow);

    // Application des formules et mise en forme finale
    finalizeSheet(motmodSheet, lastRow);
}

function prepareExtractionSheet(sheet: ExcelScript.Worksheet) {
    // Réglage de la hauteur des lignes
    sheet.getRange("A:XFD").getFormat().setRowHeight(15.75);

    // Suppression des colonnes inutiles
    sheet.getRange("B:C").delete(ExcelScript.DeleteShiftDirection.left);

    // Déplacement et gestion de la colonne 'Amount'
    let amountColumn = sheet.getRange("H:H");
    amountColumn.insert(ExcelScript.InsertShiftDirection.right);
    amountColumn.copyFrom(sheet.getRange("K:K"), ExcelScript.RangeCopyType.all, false, false);
    sheet.getRange("K:K").delete(ExcelScript.DeleteShiftDirection.left);

    // Insertion de la colonne de commentaires
    sheet.getRange("I:I").insert(ExcelScript.InsertShiftDirection.right);
    sheet.getRange("I1").setValue("commentaires");

    // Suppression des colonnes supplémentaires
    sheet.getRange("O:V").delete(ExcelScript.DeleteShiftDirection.left);

    // Ajustement automatique des colonnes
    sheet.getRange("A:N").getFormat().autofitColumns();

    // Insertion et gestion de clés pour la recherche
    insertSearchKeys(sheet);
}

function insertSearchKeys(sheet: ExcelScript.Worksheet) {
    // Création d'une clé de recherche
    sheet.getRange("A:A").insert(ExcelScript.InsertShiftDirection.right);
    sheet.getRange("A1").setFormulaLocal("concat");
    sheet.getRange("A2").setFormulaLocal("=concatener(E2;H2)");
    sheet.getRange("A2").autoFill("A2:A" + sheet.getUsedRange().getLastRow().getRowIndex(), ExcelScript.AutoFillType.fillDefault);
}

function transferData(extractionSheet: ExcelScript.Worksheet, motmodSheet: ExcelScript.Worksheet, lastRow: number) {
    // Copie des données vers la feuille MOTMOD
    motmodSheet.getRange("A:N").copyFrom(extractionSheet.getRange("B:O"), ExcelScript.RangeCopyType.values, false, false);

    // Application de la formule de recherche V pour les commentaires
    let commentRange = extractionSheet.getRange("J2:J" + lastRow);
    commentRange.setFormulaLocal("=recherchev(A2;'MOTMOD'!A$1:O$" + lastRow + ";10; FAUX)");
    range.autoFill("B2:B" + lastRow, ExcelScript.AutoFillType.fillDown);

    // Nettoyage de la feuille d'extraction
    extractionSheet.getRange().clear(ExcelScript.ClearApplyTo.all);
}

function finalizeSheet(sheet: ExcelScript.Worksheet, lastRow: number) {
    // Mise en forme des bordures et des couleurs
    applyBordersAndColors(sheet, lastRow);
}

function applyBordersAndColors(range: ExcelScript.Range) {
    const borderStyle = ExcelScript.BorderLineStyle.continuous;
    const borderColor = "000000";
    const borderWeight = ExcelScript.BorderWeight.thin;

    // Assurez-vous de passer les bonnes constantes pour les indices de bordure
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(borderStyle);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setColor(borderColor);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(borderWeight);

    // Continuez avec les autres bordures en utilisant la même structure
}

