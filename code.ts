function main(workbook: ExcelScript.Workbook) {
    // Initialisation des feuilles de travail
    let extractionSheet = workbook.getWorksheet("Extraction MOTMOD");
    let motmodSheet = workbook.getWorksheet("MOTMOD");
    let lastRow = extractionSheet.getUsedRange().getLastRow().getRowIndex();

    // Processus complet de mise à jour des données
    prepareExtractionSheet(extractionSheet);
    transferData(extractionSheet, motmodSheet, lastRow);
    applyStylesAndBorders(motmodSheet, lastRow);
}

function prepareExtractionSheet(sheet: ExcelScript.Worksheet) {
    // Réglages initiaux pour la hauteur des lignes
    sheet.getRange("A:XFD").getFormat().setRowHeight(15.75);

    // Suppression et réorganisation des colonnes
    sheet.getRange("B:C").delete(ExcelScript.DeleteShiftDirection.left);
    let amountRange = sheet.getRange("H:H");
    amountRange.insert(ExcelScript.InsertShiftDirection.right);
    amountRange.copyFrom(sheet.getRange("K:K"), ExcelScript.RangeCopyType.all, false, false);
    sheet.getRange("K:K").delete(ExcelScript.DeleteShiftDirection.left);

    // Insertion et gestion de clés pour la recherche
    let keyRange = sheet.getRange("A:A");
    keyRange.insert(ExcelScript.InsertShiftDirection.right);
    keyRange.getRange("A1").setFormulaLocal("concat");
    keyRange.getRange("A2").setFormulaLocal("=concatener(E2;H2)");
    keyRange.getRange("A2:A" + lastRow).autoFill(ExcelScript.AutoFillType.fillDown);

    // Ajout d'une colonne de commentaires
    sheet.getRange("I:I").insert(ExcelScript.InsertShiftDirection.right);
    sheet.getRange("I1").setValue("commentaires");
    sheet.getRange("O:V").delete(ExcelScript.DeleteShiftDirection.left);
    sheet.getRange("A:N").getFormat().autofitColumns();
}

function transferData(extractionSheet: ExcelScript.Worksheet, motmodSheet: ExcelScript.Worksheet, lastRow: number) {
    // Application de la formule de recherche VLOOKUP et copie des données
    let lookupRange = extractionSheet.getRange("J2:J" + lastRow);
    lookupRange.setFormulaLocal("=recherchev(A2;'MOTMOD'!A$1:O$" + lastRow + ";10;FAUX)");
    extractionSheet.getRange("J:J").copyFrom(lookupRange, ExcelScript.RangeCopyType.values, false, false);

    // Effacement des données existantes et copie des nouvelles données
    motmodSheet.getRange().clear(ExcelScript.ClearApplyTo.all);
    motmodSheet.getRange("A:N").copyFrom(extractionSheet.getRange("B:O"), ExcelScript.RangeCopyType.values, false, false);
    extractionSheet.getRange().clear(ExcelScript.ClearApplyTo.all);
}

function applyStylesAndBorders(sheet: ExcelScript.Worksheet, lastRow: number) {
    // Application des styles de bordure et de couleur
    let range = sheet.getRange("A1:N" + lastRow);
    applyBorders(range);
    applyColors(sheet, lastRow);
}

function applyBorders(range: ExcelScript.Range) {
    // Configuration des bordures internes et externes
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
    // Mise en forme des cellules avec des couleurs spécifiques
    sheet.getRange("A1").getFormat().getFill().setColor("FFFF00"); // Jaune pour la première colonne
    sheet.getRange("M1").getFormat().getFill().setColor("FFC000"); // Couleur spécifique pour la cellule M1
    sheet.getRange("G1").getFormat().getFill().setColor("FFC000"); // Couleur spécifique pour la cellule G1
    sheet.getRange("I1:I" + lastRow).getFormat().getFill().setColor("DDEBF7"); // Couleur pour la plage de I1 à la dernière rangée de I
}
