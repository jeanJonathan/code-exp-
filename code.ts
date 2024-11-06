function main(workbook: ExcelScript.Workbook) {
    // Initialisation des feuilles de travail
    let extractionSheet = workbook.getWorksheet("Extraction MOTMOD");
    let motmodSheet = workbook.getWorksheet("MOTMOD");

    // Calcul du dernier rang utilisé pour la gestion des formules et de la copie
    let lastRow = extractionSheet.getUsedRange().getLastRow().getRowIndex();

    // Processus complet de mise à jour des données
    prepareExtractionSheet(extractionSheet, lastRow);
    transferData(extractionSheet, motmodSheet, lastRow);
    applyStylesAndBorders(motmodSheet, lastRow);
}

function prepareExtractionSheet(sheet: ExcelScript.Worksheet, lastRow: number) {
    // Réglages initiaux pour la hauteur des lignes
    sheet.getRange("A:XFD").getFormat().setRowHeight(15.75);
    // Suppression et réorganisation des colonnes
    manageColumns(sheet, lastRow);
}

function manageColumns(sheet: ExcelScript.Worksheet, lastRow: number) {
    sheet.getRange("B:C").delete(ExcelScript.DeleteShiftDirection.left);

    let amountRange = sheet.getRange("H:H");
    amountRange.insert(ExcelScript.InsertShiftDirection.right);
    amountRange.copyFrom(sheet.getRange("K:K"), ExcelScript.RangeCopyType.all, false, false);
    sheet.getRange("K:K").delete(ExcelScript.DeleteShiftDirection.left);

    sheet.getRange("I:I").insert(ExcelScript.InsertShiftDirection.right);
    sheet.getRange("I1").setValue("commentaires");

    sheet.getRange("O:V").delete(ExcelScript.DeleteShiftDirection.left);
    sheet.getRange("A:N").getFormat().autofitColumns();

    // Insertion et gestion de clés pour la recherche
    manageSearchKeys(sheet, lastRow);
}

function manageSearchKeys(sheet: ExcelScript.Worksheet) {
    // Ajout d'une clé de recherche en insérant une nouvelle colonne à gauche de la colonne A
    let keyRange = sheet.getRange("A:A");
    keyRange.insert(ExcelScript.InsertShiftDirection.right);

    // Définir des formules pour concaténer les données nécessaires
    let newKeyRange = sheet.getRange("A1:A" + sheet.getUsedRange().getLastRow().getRowIndex());
    newKeyRange.getCell(0, 0).setFormulaLocal("concat");
    newKeyRange.getCell(1, 0).setFormulaLocal("=concatener(E2;H2)");

    // Utilisation de autoFill pour remplir la colonne
    let fillRange = sheet.getRange("A2:A" + sheet.getUsedRange().getLastRow().getRowIndex());
    newKeyRange.autoFill(fillRange, ExcelScript.AutoFillType.fillDown);
}


function transferData(extractionSheet: ExcelScript.Worksheet, motmodSheet: ExcelScript.Worksheet, lastRow: number) {
    motmodSheet.getRange().clear(ExcelScript.ClearApplyTo.all);
    motmodSheet.getRange("A:N").copyFrom(extractionSheet.getRange("B:O"), ExcelScript.RangeCopyType.values, false, false);
    extractionSheet.getRange().clear(ExcelScript.ClearApplyTo.all);

    let commentRange = extractionSheet.getRange("J2:J" + lastRow);
    commentRange.setFormulaLocal("=recherchev(A2;'MOTMOD'!A$1:O$" + lastRow + ";10;FAUX)");
    extractionSheet.getRange("J:J").copyFrom(commentRange, ExcelScript.RangeCopyType.values, false, false);
}

function applyStylesAndBorders(sheet: ExcelScript.Worksheet, lastRow: number) {
    let range = sheet.getRange("A1:N" + lastRow);
    applyBorders(range);
    applyColors(sheet, lastRow);
}

function applyBorders(range: ExcelScript.Range) {
    const borderStyle = ExcelScript.BorderLineStyle.continuous;
    const borderColor = "000000";
    const borderWeight = ExcelScript.BorderWeight.thin;

    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(borderStyle);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(borderStyle);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(borderStyle);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(borderStyle);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(borderStyle);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(borderStyle);

    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setColor(borderColor);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setColor(borderColor);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setColor(borderColor);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setColor(borderColor);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setColor(borderColor);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setColor(borderColor);

    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(borderWeight);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(borderWeight);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(borderWeight);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(borderWeight);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(borderWeight);
    range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(borderWeight);
}

function applyColors(sheet: ExcelScript.Worksheet, lastRow: number) {
    sheet.getRange("A1").getFormat().getFill().setColor("FFFF00");
    sheet.getRange("M1").getFormat().getFill().setColor("FFC000");
    sheet.getRange("G1").getFormat().getFill().setColor("FFC000");
    sheet.getRange("I1:I" + lastRow).getFormat().getFill().setColor("DDEBF7");
}
