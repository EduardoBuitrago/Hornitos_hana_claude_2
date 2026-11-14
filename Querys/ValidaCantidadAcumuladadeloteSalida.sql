SELECT
    T0."Name"        AS "ItemCode",
    SUM(T0."U_CantLote") AS "Cantidad",
    T0."U_almacen"   AS "Almacen"
FROM "@CLC_LTS" T0
WHERE 
    T0."Name" = '{0}'
    AND T0."U_CantLote" > 0
    AND COALESCE(T0."U_Estado", 'C') <> 'P' -- no cuenta lotes de documentos pendientes de aprobacion
    AND T0."U_almacen" = '{1}'
GROUP BY 
    T0."Name",
    T0."U_almacen"
HAVING 
    SUM(T0."U_CantLote") >= {2}