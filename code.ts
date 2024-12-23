Vue 1 : Détail par poste SAP
Description :

Chaque ligne correspond à un poste SAP unique. Les colonnes regroupent les informations sur les Pro Formas, BLs (avec quantités), et Factures. Les valeurs associées sont affichées dans des cellules concaténées.

Poste SAP	P/N	Désignation	Qté Commandée	Qté Livrée	To be delivered	BLs (Quantité)	Pro Formas (Quantité)	Factures (Quantité)	Shipping Date
1234	0000021850	SCREW, CA	53	40	13	85686187(40)	100789790(40)	1088618(40)	2024-07-01
5678	0000022670	SCREW, TA	50	25	25	85686187(25), 85655097(25)	100789790(50)	1093028(50)	2024-08-15
9101	0000023202	SCREW	70	55	15	85734031(55)	100789791(55)	1099870(55)	2024-09-04
1213	0000023520	SCREW	28	20	8	85655097(20)	100789792(20)	1088618(20)	2024-06-15
Points clés :

Les BLs sont concaténés dans une cellule unique pour un poste (avec les quantités livrées).
Les Pro Formas et Factures sont également concaténés.
Les colonnes comme Qté Commandée, Qté Livrée, et To be delivered affichent les quantités associées.
Vue 2 : Regroupement par échéance
Description :

Les données sont regroupées par échéance (Shipping Date ou autre date clé). Chaque ligne représente un regroupement pour une même échéance avec des détails sur les postes SAP associés.

Shipping Date	Poste SAP	P/N	Qté Commandée	Qté Livrée	To be delivered	BLs (Quantité)	Pro Formas (Quantité)	Factures (Quantité)
2024-07-01	1234	0000021850	53	40	13	85686187(40)	100789790(40)	1088618(40)
2024-08-15	5678	0000022670	50	25	25	85686187(25), 85655097(25)	100789790(50)	1093028(50)
2024-09-04	9101	0000023202	70	55	15	85734031(55)	100789791(55)	1099870(55)
2024-06-15	1213	0000023520	28	20	8	85655097(20)	100789792(20)	1088618(20)
Points clés :

Chaque ligne correspond à une date spécifique (Shipping Date).
Les BLs, Pro Formas, et Factures sont liés aux postes SAP concernés pour cette date.
Les colonnes comme Qté Commandée et To be delivered affichent les totaux pour chaque ligne d’échéance.
Résumé des différences entre les vues
Vue 1 :
Focus sur un poste SAP unique.
Les informations des BLs, Pro Formas, et Factures sont concaténées dans des colonnes associées.
Vue 2 :
Regroupement par Shipping Date (ou échéance clé).
Présente une vue synthétique par date, avec des détails consolidés.
