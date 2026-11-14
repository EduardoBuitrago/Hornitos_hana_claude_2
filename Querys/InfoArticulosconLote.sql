SELECT
    'N' AS "Sel",
    '{0}' AS "ItemCode",
    '{1}' AS "Tipo doc",
    {2} AS "Num Doc",
    CAST('' AS NVARCHAR(50)) AS "Numero lote",
    CAST(({3}) AS DOUBLE) AS "Cantidad",
    IFNULL(
        ADD_DAYS(
            CURRENT_DATE,
            (SELECT T0."U_V_Util" 
             FROM OITM T0 
             WHERE T0."ItemCode" = '{0}')
        ),
        ''
    ) AS "Fecha vence",
    CURRENT_DATE AS "Fecha Admin",
    '{4}' AS "Almacen",
    '' AS "Index",
    {5} AS "NumLinea"
FROM DUMMY