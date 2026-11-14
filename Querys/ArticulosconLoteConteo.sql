
SELECT 
    'N' AS "Sel",
    {5} AS "NumLinea",
    {1} AS "Numero Doc",
    '{4}' AS "Tipo Doc",
    T0."ItemCode",
    '{2}' AS "Almacen",
    CAST(({3}) AS DOUBLE) AS "Cantidad",
    CASE 
        WHEN CAST(({3}) AS DOUBLE) > 0 THEN 'E'
        WHEN CAST(({3}) AS DOUBLE) < 0 THEN 'S'
    END AS "Tipo Conteo"
FROM OITM T0
INNER JOIN OWHS T1 
    ON '{2}' = T1."WhsCode" 
    AND T1."U_UtilLot" = 'Y'
WHERE 
    T0."ItemCode" = '{0}' 
    AND T0."QryGroup28" = 'Y'