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
    prepareExtractionSheet(extractionSheet);

    // Mise à jour des données vers la feuille MOTMOD
    updateDataToMOTMOD(extractionSheet, motmodSheet);

    // Mise en forme de la feuille MOTMOD
    formatMOTMODSheet(motmodSheet);
}

function prepareExtractionSheet(sheet: ExcelScript.Worksheet) {
    // Réglages initiaux pour la hauteur des lignes
    sheet.getRange("A:XFD").getFormat().setRowHeight(15.75);

    // Suppression et réorganisation des colonnes
    manageColumns(sheet);

    // Insertion et gestion de clés pour la recherche
    manageSearchKeys(sheet);
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

function manageSearchKeys(sheet: ExcelScript.Worksheet) {
    // Ajout de clés de recherche
    sheet.getRange("A:A").insert(ExcelScript.InsertShiftDirection.right);
    sheet.getRange("A1").setFormulaLocal("concat");
    sheet.getRange("A2").setFormulaLocal("=concatener(E2;H2)");
    sheet.getRange("A2:A" + sheet.getUsedRange().getLastRow().getRowIndex()).autoFill();
}

function updateDataToMOTMOD(extractionSheet: ExcelScript.Worksheet, motmodSheet: ExcelScript.Worksheet) {
    // Copie des données depuis Extraction MOTMOD vers MOTMOD
    motmodSheet.getRange("A:N").copyFrom(extractionSheet.getRange("B:O"), ExcelScript.RangeCopyType.values, false, false);
    // Nettoyage après copie
    extractionSheet.getRange().clear(ExcelScript.ClearApplyTo.all);
}

function formatMOTMODSheet(sheet: ExcelScript.Worksheet) {
    // Application des styles de bordure et de couleur
    let range = sheet.getRange("A1:N" + sheet.getUsedRange().getLastRow().getRowIndex());
    applyBorders(range);
    applyStyles(sheet);
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

function applyStyles(sheet: ExcelScript.Worksheet) {
    // Mise en forme des cellules avec des couleurs spécifiques pour certaines cellules
    sheet.getRange("A1").getFormat().getFill().setColor("FFFF00"); // Jaune pour la première colonne
    sheet.getRange("M1").getFormat().getFill().setColor("FFC000"); // Couleur spécifique pour la cellule M1
    sheet.getRange("G1").getFormat().getFill().setColor("FFC000"); // Couleur spécifique pour la cellule G1
    sheet.getRange("I1:I" + sheet.getUsedRange().getLastRow().getRowIndex()).getFormat().getFill().setColor("DDEBF7"); // Couleur pour la plage de I1 à la dernière rangée de I
}


function main(workbook: ExcelScript .Workbook) { let ExtractionMOTMODSheet = workbook.getWorksheet("Extraction MOTMOD"); laissez MOTMODSheet = workbook.getWorksheet("MOTMOD"); let lastrow = ExtractionMOTMODSheet.getUsedRange().getLastRow().getRowIndex(); // Fixe la hauteur des lignes de ExtractionMOTMODSheet à 15.75 ExtractionMOTMODSheet.getRange("A:XFD").getFormat().setRowHeight(15.75); // supprime les colonnes B et C de ExtractionMOTMODSheet ExtractionMOTMODSheet.getRange("B:C").delete(ExcelScript.DeleteShiftDirection.left); // Déplace la colonne Amount de ExtractionMOTMODSheet ExtractionMOTMODSheet.getRange("H:H").insert (ExcelScript.InsertShiftDirection.right); ExtractionMOTMODSheet.getRange("H:H").copyFrom(ExtractionMOTMODSheet.getRange("K:K"), ExcelScript.RangeCopyType.all, false, false); ExtractionMOTMODSheet.getRange("K: K").delete(ExcelScript.DeleteShiftDirection.left); // Insérer une colonne de commentaires après Amount ExtractionMOTMODSheet ExtractionMOTMODSheet.getRange("I:I").insert(ExcelScript.InsertShiftDirection.right); ExtractionMOTMODSheet.getRange("I1").setValue("commentaires"); // Supprime les colonnes O:V de ExtractionMOTMODSheet ExtractionMOTMODSheet.getRange("O:V").delete(ExcelScript.DeleteShiftDirection.left); // Autofit des colonnes de ExtractionMOTMODSheet ExtractionMOTMODSheet.getRange("A:N").getFormat().autofitColumns(); // Insère en colonne A:A de ExtractionMOTMODSheet une concaténation pour créer une clé de recherche ExtractionMOTMODSheet.getRange("A:A").insert(ExcelScript.InsertShiftDirection.right); ExtractionMOTMODSheet.getRange("A1").setFormulaLocal("concat"); ExtractionMOTMODSheet.getRange("A2").setFormulaLocal("=concatener(E2;H2)"); ExtractionMOTMODSheet.getRange("A2").autoFill(); // Insère en colonne A:A de MOTMODSheet une concaténation pour créer une clé de recherche MOTMODSheet.getRange("A:A").insert(ExcelScript.InsertShiftDirection.right); MOTMODSheet.getRange("A1").setFormulaLocal("concat"); MOTMODSheet.getRange("A2").setFormulaLocal("=concatener(E2;H2)"); MOTMODSheet.getRange("A2").autoFill(); // Créé la formule de recherche en colonne commentaires de ExtractionMOTMODSheet ExtractionMOTMODSheet.getRange("J2").setFormulaLocal("=recherchev(A2;'MOTMOD'!A$1:O$" + lastrow.toString() + ";10; FAUX)"); ExtractionMOTMODSheet.getRange("J2").autoFill("J2:J" + lastrow.toString()); // Copie colle en valeur la colonne commentaires de ExtractionMOTMODSheet ExtractionMOTMODSheet.getRange("J:J").copyFrom(ExtractionMOTMODSheet.getRange("J:J"), ExcelScript.RangeCopyType.values, false, false); // Efface les données de MOTMOD MOTMODSheet.getRange().clear(ExcelScript.ClearApplyTo.all); // Copiez les données de Extraction MOTMOD sur MOTMOD et effacez les données de ExtractionMOTMOD MOTMODSheet.getRange("A:N").copyFrom(ExtractionMOTMODSheet.getRange("B:O"), ExcelScript.RangeCopyType.values, false, false) ; ExtractionMOTMODSheet.getRange().clear(ExcelScript.ClearApplyTo.all); //Mise en forme du tableau MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setColor("000000"); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setColor("000000"); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setColor("000000"); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setColor("000000"); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setColor("000000"); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setColor("000000"); MOTMODSheet.getRange("A1:N" + lastrow).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin); //1ere colonne en jaune MOTMODSheet.getRange("A1:A" + lastrow).getFormat().getFill().setColor("FFFF00"); //Mise en forme date MOTMODSheet.getRange("N:N").setNumberFormatLocal("jj/mm/aaaa"); MOTMODSheet.getRange("J:J").setNumberFormatLocal("jj/mm/aaaa"); //Mise Forme Couleurs MOTMODSheet.getRange("M1").getFormat().getFill().setColor("FFC000"); MOTMODSheet.getRange("G1").getFormat().getFill().setColor("FFC000"); MOTMODSheet.getRange("I1:I" + lastrow.toString()).getFormat().getFill().setColor("DDEBF7"); //supprime les espaces en colonne H MOTMODSheet.getRange("H:H").replaceAll(" ", "", { completeMatch: false, matchCase: false }); //fixe le format de Amount MOTMODSheet.getRange("H:H").setNumberFormatLocal("_-* # ##0,00\\ [$€-fr-FR]_-;-* # ##0, 00\\ [$€-fr-FR]_-;_-* \"-\"??\\ [$€-fr-FR]_-;_-@_-"); //largeurs de colonnes MOTMODSheet.getRange("A:B").getFormat().setColumnWidth(41); MOTMODSheet.getRange("C:D").getFormat().setColumnWidth(80); MOTMODSheet.getRange("E:E").getFormat().setColumnWidth(50); MOTMODSheet.getRange("F:F").getFormat().setColumnWidth(160); MOTMODSheet.getRange("G:G").getFormat().setColumnWidth(45); MOTMODSheet.getRange("H:H").getFormat().setColumnWidth(90); MOTMODSheet.getRange("I:I").getFormat().setColumnWidth(150); MOTMODSheet.getRange("J:J").getFormat().setColumnWidth(64); MOTMODSheet.getRange("K:K").getFormat().setColumnWidth(120); MOTMODSheet.getRange("L:L").getFormat().setColumnWidth(70); MOTMODSheet.getRange("M:N").getFormat().setColumnWidth(64); }getFormat().getFill().setColor("FFFF00"); //Mise en forme date MOTMODSheet.getRange("N:N").setNumberFormatLocal("jj/mm/aaaa"); MOTMODSheet.getRange("J:J").setNumberFormatLocal("jj/mm/aaaa"); //Mise Forme Couleurs MOTMODSheet.getRange("M1").getFormat().getFill().setColor("FFC000"); MOTMODSheet.getRange("G1").getFormat().getFill().setColor("FFC000"); MOTMODSheet.getRange("I1:I" + lastrow.toString()).getFormat().getFill().setColor("DDEBF7"); //supprime les espaces en colonne H MOTMODSheet.getRange("H:H").replaceAll(" ", "", { completeMatch: false, matchCase: false }); //fixe le format de Amount MOTMODSheet.getRange("H:H").setNumberFormatLocal("_-* # ##0,00\\ [$€-fr-FR]_-;-* # ##0, 00\\ [$€-fr-FR]_-;_-* \"-\"??\\ [$€-fr-FR]_-;_-@_-"); //largeurs de colonnes MOTMODSheet.getRange("A:B").getFormat().setColumnWidth(41); MOTMODSheet.getRange("C:D").getFormat().setColumnWidth(80); MOTMODSheet.getRange("E:E").getFormat().setColumnWidth(50); MOTMODSheet.getRange("F:F").getFormat().setColumnWidth(160); MOTMODSheet.getRange("G:G").getFormat().setColumnWidth(45); MOTMODSheet.getRange("H:H").getFormat().setColumnWidth(90); MOTMODSheet.getRange("I:I").getFormat().setColumnWidth(150); MOTMODSheet.getRange("J:J").getFormat().setColumnWidth(64); MOTMODSheet.getRange("K:K").getFormat().setColumnWidth(120); MOTMODSheet.getRange("L:L").getFormat().setColumnWidth(70); MOTMODSheet.getRange("M:N").getFormat().setColumnWidth(64); }getFormat().getFill().setColor("FFFF00"); //Mise en forme date MOTMODSheet.getRange("N:N").setNumberFormatLocal("jj/mm/aaaa"); MOTMODSheet.getRange("J:J").setNumberFormatLocal("jj/mm/aaaa"); //Mise Forme Couleurs MOTMODSheet.getRange("M1").getFormat().getFill().setColor("FFC000"); MOTMODSheet.getRange("G1").getFormat().getFill().setColor("FFC000"); MOTMODSheet.getRange("I1:I" + lastrow.toString()).getFormat().getFill().setColor("DDEBF7"); //supprime les espaces en colonne H MOTMODSheet.getRange("H:H").replaceAll(" ", "", { completeMatch: false, matchCase: false }); //fixe le format de Amount MOTMODSheet.getRange("H:H").setNumberFormatLocal("_-* # ##0,00\\ [$€-fr-FR]_-;-* # ##0, 00\\ [$€-fr-FR]_-;_-* \"-\"??\\ [$€-fr-FR]_-;_-@_-"); //largeurs de colonnes MOTMODSheet.getRange("A:B").getFormat().setColumnWidth(41); MOTMODSheet.getRange("C:D").getFormat().setColumnWidth(80); MOTMODSheet.getRange("E:E").getFormat().setColumnWidth(50); MOTMODSheet.getRange("F:F").getFormat().setColumnWidth(160); MOTMODSheet.getRange("G:G").getFormat().setColumnWidth(45); MOTMODSheet.getRange("H:H").getFormat().setColumnWidth(90); MOTMODSheet.getRange("I:I").getFormat().setColumnWidth(150); MOTMODSheet.getRange("J:J").getFormat().setColumnWidth(64); MOTMODSheet.getRange("K:K").getFormat().setColumnWidth(120); MOTMODSheet.getRange("L:L").getFormat().setColumnWidth(70); MOTMODSheet.getRange("M:N").getFormat().setColumnWidth(64); }getRange("F:F").getFormat().setColumnWidth(160); MOTMODSheet.getRange("G:G").getFormat().setColumnWidth(45); MOTMODSheet.getRange("H:H").getFormat().setColumnWidth(90); MOTMODSheet.getRange("I:I").getFormat().setColumnWidth(150); MOTMODSheet.getRange("J:J").getFormat().setColumnWidth(64); MOTMODSheet.getRange("K:K").getFormat().setColumnWidth(120); MOTMODSheet.getRange("L:L").getFormat().setColumnWidth(70); MOTMODSheet.getRange("M:N").getFormat().setColumnWidth(64); }getRange("F:F").getFormat().setColumnWidth(160); MOTMODSheet.getRange("G:G").getFormat().setColumnWidth(45); MOTMODSheet.getRange("H:H").getFormat().setColumnWidth(90); MOTMODSheet.getRange("I:I").getFormat().setColumnWidth(150); MOTMODSheet.getRange("J:J").getFormat().setColumnWidth(64); MOTMODSheet.getRange("K:K").getFormat().setColumnWidth(120); MOTMODSheet.getRange("L:L").getFormat().setColumnWidth(70); MOTMODSheet.getRange("M:N").getFormat().setColumnWidth(64); }

