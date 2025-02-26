SELECT 
    "Date on Which Record Was Created" AS OrderDate,
    "Sales Document" AS SalesOrderID,
    "Sales Document Type" AS SalesDocumentType,
    "SD document category" AS SDDocCategory,
    "Shipping Conditions" AS ShippingConditions
FROM "Nom_de_ta_base".VBAK
WHERE "Date on Which Record Was Created" BETWEEN '20200101' AND '20251231'
AND "Shipping Conditions" = 'EX'
