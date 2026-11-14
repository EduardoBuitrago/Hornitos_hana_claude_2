SELECT
    'N' AS "Sel",
    T0."Name"       AS "ItemCode",
    T0."U_TipDoc"   AS "Tipo doc",
    T0."U_NumDoc"   AS "Num Doc",
    T0."U_NumLot"   AS "Numero lote",
    T0."U_CantLote" AS "Cantidad disponible",
    CAST((0) AS DOUBLE) AS "Cantidad",
    T0."U_FchVenLt" AS "Fecha vence",
    T0."U_FchAdmLt" AS "Fecha Admin",
    T0."U_almacen"  AS "Almacen",
    '' AS "Index",
    0 AS "NumLinea"
FROM "@CLC_LTS" T0
WHERE 
    T0."Name" = '{0}'
    AND T0."U_CantLote" > 0
    AND T0."U_almacen" = '{1}'
ORDER BY 
    T0."U_FchVenLt"