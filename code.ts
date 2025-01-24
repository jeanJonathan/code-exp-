1.2. Commandes en cours

Description : Commandes non encore finalisées ou partiellement livrées.
Conditions :
SalesDocumentOverallProcessingStatusName = "Not yet processed" ou "Partially processed".
Formule DAX :
NbCommandesEnCours =
CALCULATE(
    COUNTROWS(Table),
    Table[SalesDocumentCategoryName] = "Order" &&
    Table[SalesDocumentOverallProcessingStatusName] IN {"Not yet processed", "Partially processed"}
)
1.3. Commandes finalisées

Description : Commandes complètement traitées et livrées.
Conditions :
SalesDocumentOverallProcessingStatusName = "Completely processed".
Formule DAX :
NbCommandesFinalisées =
CALCULATE(
    COUNTROWS(Table),
    Table[SalesDocumentCategoryName] = "Order" &&
    Table[SalesDocumentOverallProcessingStatusName] = "Completely processed"
)
1.4. Montant total des commandes

Description : Somme des montants associés aux commandes.
Formule DAX :
MontantTotalCommandes =
CALCULATE(
    SUM(Table[SalesDocumentNetValue]),
    Table[SalesDocumentCategoryName] = "Order"
)
1.5. Montant des commandes en cours

Description : Montant total des commandes non encore finalisées.
Formule DAX :
MontantCommandesEnCours =
CALCULATE(
    SUM(Table[SalesDocumentNetValue]),
    Table[SalesDocumentCategoryName] = "Order" &&
    Table[SalesDocumentOverallProcessingStatusName] IN {"Not yet processed", "Partially processed"}
)
1.6. Délai moyen de traitement des commandes

Description : Temps moyen entre la création de la commande et sa finalisation (ou livraison).
Formule DAX (colonne calculée pour les délais en jours) :
DélaiTraitement = 
DATEDIFF(
    Table[SalesDocumentCreationDate],
    Table[DateFinalisation], -- Colonne représentant la date de finalisation ou livraison
    DAY
)
Formule pour la moyenne :
DélaiMoyenTraitement =
AVERAGEX(
    FILTER(
        Table,
        Table[SalesDocumentCategoryName] = "Order" &&
        NOT(ISBLANK(Table[DateFinalisation]))
    ),
    Table[DélaiTraitement]
)
1.7. Commandes annulées

Description : Commandes ayant été créées mais ensuite annulées.
Conditions :
Une commande peut être identifiée comme annulée par un statut ou un montant à zéro.
Formule DAX :
NbCommandesAnnulées =
CALCULATE(
    COUNTROWS(Table),
    Table[SalesDocumentCategoryName] = "Order" &&
    Table[SalesDocumentNetValue] = 0
)
1.8. Répartition des commandes par canal de distribution

Description : Distribution des commandes selon les canaux (ex. : client direct, TUAG, etc.).
Formule DAX :
NbCommandesParCanal =
SUMMARIZE(
    FILTER(Table, Table[SalesDocumentCategoryName] = "Order"),
    Table[CanalDistribution],
    "NbCommandes", COUNTROWS(Table)
)
1.9. Commandes par secteur d’activité

Description : Nombre et montant des commandes par secteur (ex. : moteur neuf, rechange, etc.).
Formule DAX :
NbCommandesParSecteur =
SUMMARIZE(
    FILTER(Table, Table[SalesDocumentCategoryName] = "Order"),
    Table[SecteurActivité],
    "NbCommandes", COUNTROWS(Table),
    "MontantTotal", SUM(Table[SalesDocumentNetValue])
)
1.10. Taux de traitement des commandes

Description : Pourcentage des commandes finalisées par rapport au total des commandes.
Formule DAX :
TauxTraitement =
DIVIDE(
    [NbCommandesFinalisées],
    [NbCommandes],
    0
), partiellement traitées, ou non traitées), leur transformation en commandes, et leur état (en cours, expirées ou refusées). L