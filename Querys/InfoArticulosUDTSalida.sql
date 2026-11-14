SELECT
    T0."Name"        AS "ItemCode",
    T0."U_TipDoc"    AS "Tipo doc",
    T0."U_NumDoc"    AS "Num Doc",
    T0."U_NumLot"    AS "Numero lote",
    T0."U_Cantidad"  AS "Cantidad",
    T0."U_FchSal"    AS "Fecha salida",
    T0."U_Almacen"   AS "Almacen"
FROM "@CLC_LTSSA" T0
WHERE 
    T0."Name" = '{0}' 
    AND T0."U_NumDoc" = {1} 
    AND T0."U_TipDoc" = '{2}'