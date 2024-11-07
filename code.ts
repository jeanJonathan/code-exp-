function main(workbook: ExcelScript.Workbook) {
    const extractSheet = workbook.getWorksheet("Extract FCU");
    const fcuHmuSheet = workbook.getWorksheet("FCU-HMU");
    
    // Préparation des données dans le fichier extract
    prepareExtractData(extractSheet);

    // Transfert et mise à jour des données dans le fichier FCU-HMU
    transferAndUpdateData(extractSheet, fcuHmuSheet);

    // Finalisation et sauvegarde des modifications
    finalizeUpdates(fcuHmuSheet);
}

function prepareExtractData(sheet: ExcelScript.Worksheet) {
    // Supprimer les zéros non nécessaires dans 'ID project SAP'
    adjustProjectIDs(sheet);
    // Ajoutez d'autres préparations de données si nécessaire
}

function adjustProjectIDs(sheet: ExcelScript.Worksheet) {
    let range = sheet.getUsedRange();
    let column = range.getColumn('E'); // Supposons que 'ID project SAP' est dans la colonne E
    let values = column.getValues();
    // Enlever les zéros initiaux
    for (let i = 0; i < values.length; i++) {
        values[i][0] = values[i][0].replace(/^0000/, '');
    }
    column.setValues(values);
}

function transferAndUpdateData(extractSheet: ExcelScript.Worksheet, fcuHmuSheet: ExcelScript.Worksheet) {
    // Utiliser les ID project SAP ajustés pour trouver des correspondances et mettre à jour les données
    const lastRow = extractSheet.getUsedRange().getLastRow().getRowIndex();
    for (let i = 2; i <= lastRow; i++) {
        const extractID = extractSheet.getRange(`E${i}`).getText(); // ID de l'extractSheet
        const comments = extractSheet.getRange(`T${i}`).getText(); // Supposons que les commentaires sont dans la colonne T
        updateFcuHmuSheet(fcuHmuSheet, extractID, comments);
    }
}

function updateFcuHmuSheet(fcuHmuSheet: ExcelScript.Worksheet, extractID: string, comments: string) {
    const rows = fcuHmuSheet.getUsedRange().getRowCount();
    for (let i = 2; i <= rows; i++) {
        const fcuID = fcuHmuSheet.getRange(`E${i}`).getText(); // Supposons que 'ID project SAP' est aussi dans la colonne E
        if (fcuID === extractID) {
            fcuHmuSheet.getRange(`U${i}`).setValue(comments); // Mettre à jour la colonne 'Comments' dans FCU-HMU
            // Mettre à jour d'autres colonnes au besoin
            break;
        }
    }
}

function finalizeUpdates(sheet: ExcelScript.Worksheet) {
    // Enregistrer le workbook ou faire d'autres nettoyages
    console.log('Updates finalized.');
}
