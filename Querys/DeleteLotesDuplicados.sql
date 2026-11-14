SELECT
    T0."Code"
FROM "@CLC_LTS" T0
WHERE
    T0."Name" = '{0}'
    AND T0."U_NumLot" = '{1}'
    AND T0."U_almacen" = '{2}'
    AND T0."Code" <> '{3}'
ORDER BY
    T0."Code";