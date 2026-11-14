SELECT
    'N' AS "Sel",
    '{0}' AS "ItemCode",
    '{1}' AS "Tipo doc",
    {2} AS "Num Doc",
    CAST('' AS NVARCHAR(50)) AS "Numero lote",
    0 AS "Cantidad disponible",
    CAST({3} AS DECIMAL(19,6)) AS "Cantidad",
    IFNULL(
        (
            SELECT
                TO_NVARCHAR(
                    ADD_DAYS(CURRENT_DATE, T2."U_V_Util"),
                    'YYYY-MM-DD'
                )
            FROM OITM T2
            WHERE T2."ItemCode" = '{0}'
        ),
        ''
    ) AS "Fecha vence",
    CURRENT_TIMESTAMP AS "Fecha Admin",
    '{4}' AS "Almacen",
    '' AS "Index"
FROM OITM T0
INNER JOIN OWHS T1
    ON T1."WhsCode" = '{4}'
    AND T1."U_UtilLot" = 'Y'
WHERE
    T0."ItemCode" = '{0}'
    AND T0."QryGroup28" = 'Y';