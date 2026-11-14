SELECT
    T0."Name",
    '{3}' AS "TipDoc",
    {1} AS "NumDoc",
    T0."U_NumDoc" AS "NumDoc2",
    T0."U_NumLot",
    T0."U_CantLote" AS "U_CantLote",
    T0."U_FchVenLt",
    T0."U_FchAdmLt",
    T0."U_almacen",
    T0."U_TipoCont",
    T0."U_TpDocEnt",
    T0."U_NmDocEnt",
    T0."U_RowNum"
FROM "@CLC_LTSSATEMP" T0
WHERE
    T0."Name" = '{0}'
    AND T0."U_NumDoc" = {4}
    AND T0."U_almacen" = '{2}'
    AND T0."U_CantLote" > 0;