SELECT
    'N' AS "Sel",
    '{0}' AS "ItemCode",
    {1} AS "Num Doc",
    '{2}' AS "Tipo doc",
    '{4}' AS "Numero lote",
    CAST({3} AS DOUBLE) AS "Cantidad",
    TO_DATE('{5}', 'DD/MM/YYYY') AS "Fecha vence",
    '{6}' AS "Almacen",
    {7} AS "Index",
    {8} AS "NumLinea"
FROM DUMMY