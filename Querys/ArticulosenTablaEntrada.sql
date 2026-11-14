SELECT
    COUNT(*) AS "Total"
FROM "@CLC_LTS" T0
WHERE 
    T0."Name" = '{0}' 
    AND T0."U_TipDoc" = '{1}' 
    AND T0."U_NumDoc" = {2}