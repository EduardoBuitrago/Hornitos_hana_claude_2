SELECT
    'N' AS "Sel",
    '{0}' AS "ItemCode",
    {1} AS "Num Doc",
    '{2}' AS "Tipo doc",
    CAST(({3}) AS DOUBLE) AS "Cantidad",
    CAST('{4}' AS NVARCHAR(50)) AS "Numero lote",
    TO_DATE('{5}', 'DD/MM/YYYY') AS "Fecha vence",
    TO_DATE('{6}', 'DD/MM/YYYY') AS "Fecha Admin",
    '{7}' AS "Almacen",
    {8} AS "Index",
    {9} AS "NumLinea"
FROM DUMMY