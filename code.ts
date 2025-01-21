Étape 2 : Ajouter les colonnes calculées nécessaires
1. Calcul de la date de validité de l’offre

Si la colonne DateValidité n’existe pas, calculez-la à partir de la date de création :
DateValidité = 
Table[SalesDocumentCreationDate] + 30
2. Indiquer si l’offre est en cours

Créez une colonne calculée pour identifier les offres en cours :
OffreEnCours = 
IF(
    Table[SalesDocumentCategoryCode] = "Quotation" &&
    Table[SalesDocumentOverallProcessingStatusName] IN {"Not yet processed", "Partially processed"} &&
    TODAY() <= Table[DateValidité],
    "En cours",
    "Non en cours"
)
3. Calcul des jours restants avant expiration

Ajoutez une colonne calculée pour afficher les jours restants avant expiration :
JoursRestants = 
DATEDIFF(TODAY(), Table[DateValidité], DAY)
Étape 3 : Créer les mesures pour les KPI
1. Nombre total d’offres en cours

NbOffresEnCours = 
CALCULATE(
    COUNTROWS(Table),
    Table[OffreEnCours] = "En cours"
)
2. Montant total des offres en cours

Si une colonne MontantTotal existe :
MontantOffresEnCours = 
CALCULATE(
    SUM(Table[MontantTotal]),
    Table[OffreEnCours] = "En cours"
)
3. Nombre d’offres par statut global

NbOffresParStatut =
CALCULATE(
    COUNTROWS(Table),
    Table[SalesDocumentCategoryCode] = "Quotation",
    Table[OffreEnCours] = "En cours"
)
Étape 4 : Construire les visualisations dans Power BI
1. Histogramme des offres en cours par statut global

Axes :
X : SalesDocumentOverallProcessingStatusName (Not yet processed, Partially processed).
Y : Nombre d’offres.
Filtre :
SalesDocumentCategoryCode = "Quotation".
OffreEnCours = "En cours".
Insights attendus :
Visualiser la répartition des offres en cours selon leur statut.
2. Carte de chaleur des délais restants par client

Axes :
X : CustomerCompanyCustomerName (Nom client).
Y : JoursRestants (jours avant expiration).
Couleurs :
Rouge : Offres expirant dans moins de 7 jours.
Orange : Offres expirant dans 7 à 15 jours.
Vert : Plus de 15 jours restants.
Filtre :
OffreEnCours = "En cours".
Insights attendus :
Identifier les clients ayant des offres urgentes.
3. Tableau détaillé des offres en cours

Colonnes :
SalesDocumentID (identifiant de l’offre).
CustomerCompanyCustomerName (nom du client).
SalesDocumentCreationDate (date de création).
DateValidité (date de validité calculée).
JoursRestants (jours restants).
SalesDocumentOverallProcessingStatusName (statut global).
MontantTotal (si disponible).
Filtres :
OffreEnCours = "En cours".
Insights attendus :
Liste exploitable des offres valides, triée par priorité.
4. KPI

Nombre total d’offres en cours.
Montant total des offres en cours.
Moyenne des jours restants :
MoyenneJoursRestants =
AVERAGEX(
    FILTER(Table, Table[OffreEnCours] = "En cours"),
    Table[JoursRestants]
)
Synthèse des visualisations
Visualisation	Objectif	Insights attendus
Histogramme par statut global	Répartir les offres selon leur statut global.	Identifier les offres partiellement traitées.
Carte de chaleur des délais restants	Prioriser les offres proches de l’expiration.	Connaître les clients avec des offres urgentes.
Tableau détaillé des offres en cours	Liste détaillée des offres valides.	Identifier les clients ou montants à forte priorité.
KPI : Nombre et montant total des offres	Suivi global des offres en cours.	Visualiser rapidement les volumes et montants totaux.
Étape suivante
Configurez les relations dans Power BI, appliquez les formules DAX ci-dessus et construisez les visualisations suggérées. Une fois terminé, vous pourrez analyser vos données et répondre aux priorités des offres en cours.

Si vous avez besoin d’aide pour configurer une visualisation ou ajuster les formules, je suis à votre disposition.