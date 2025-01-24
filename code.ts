NbOffresRefusées =
CALCULATE(
    COUNTROWS(Table),
    Table[SalesDocumentCategoryName] = "Quotation" && -- Type "Quotation"
    ISBLANK(RELATED(SalesDocumentFlow[SalesDocumentFlowSubsequentDocumentID])) && -- Pas de commande associée
    TODAY() > Table[SalesDocumentCreationDate] + 30 && -- Expirée
    Table[SalesDocumentOverallProcessingStatusName] IN {"Not yet processed", "Partially processed"} -- Statut partiellement ou non traité
)
