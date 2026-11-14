SELECT
    T0."Name" AS "ItemCode",
    T0."U_NumLot" AS "Numero lote",
    'OIQR' AS "Tipo Doc",
    {3} AS "NumDoc conteo",
    {4} AS "NumDoc conta",
    CAST(T0."U_CantLote" AS DECIMAL(19,6)) AS "Cantidad",
    T0."U_almacen",
    T0."U_TipDoc",
    T0."U_NumDoc",
    T0."U_almacen"
FROM "@CLC_LTS" T0
WHERE
    T0."Name" = '{0}'
    AND T0."U_almacen" = '{1}'
    AND T0."U_CantLote" > 0
    AND T0."U_NumLot" NOT IN ({2});