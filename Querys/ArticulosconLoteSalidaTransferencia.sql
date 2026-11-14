SELECT 
    'N' AS "Sel",
    {6} AS "NumLinea",
    {1} AS "Numero Doc",
    T0."ItemCode",
    T0."ItemName" AS "Descripcion",
    '{2}' AS "Almacen",
    '{5}' AS "Almacen destino",
    CAST({3} AS DOUBLE) AS "Cantidad",
    '{4}' AS "Tipo Doc"
FROM OITM T0
INNER JOIN OWHS T1 
    ON T1."WhsCode" = '{2}' 
    AND T1."U_UtilLot" = 'Y'
WHERE 
    T0."ItemCode" = '{0}' 
    AND T0."QryGroup28" = 'Y'