SELECT
    '{0}' AS "ItemCode",
    '{1}' AS "Tipo doc",
    {2} AS "Num Doc",
    '{3}' AS "Numero lote",
    CAST({4} AS DOUBLE) AS "Cantidad",
    TO_DATE('{5}', 'DD/MM/YYYY') AS "Fecha vence",
    TO_DATE('{6}', 'DD/MM/YYYY') AS "Fecha Admin",
    '{7}' AS "Almacen",
    '{8}' AS "Index",
    {9} AS "NumLinea"
FROM DUMMY