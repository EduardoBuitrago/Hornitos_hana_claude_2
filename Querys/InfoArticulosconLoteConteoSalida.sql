SELECT
    'N' AS "Sel",
    MAX(T0."Name") AS "ItemCode",
    MAX(T0."U_TipDoc") AS "Tipo doc",
    MAX(T0."U_NumDoc") AS "Num Doc",
    T0."U_NumLot" AS "Numero lote",
    SUM(T0."U_CantLote") AS "Cantidad disponible",
    CAST(0 AS DECIMAL(19,6)) AS "Cantidad",
    MAX(T0."U_FchVenLt") AS "Fecha vence",
    MAX(T0."U_FchAdmLt") AS "Fecha Admin",
    MAX(T0."U_almacen") AS "Almacen",
    '' AS "Index",
    {2} AS "NumLinea"
FROM "@CLC_LTS" T0
WHERE
    T0."Name" = '{0}'
    AND T0."U_CantLote" > 0
    AND T0."U_almacen" = '{1}'
GROUP BY
    T0."U_NumLot"
ORDER BY
    MAX(T0."U_FchVenLt");