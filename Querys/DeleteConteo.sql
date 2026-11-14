DELETE FROM "@CLC_LTSSATEMP"
WHERE
    "Name" = '{0}'
    AND "U_NumDoc" = {1}
    AND "U_almacen" = '{2}'
    AND "U_NumLot" = '{3}';