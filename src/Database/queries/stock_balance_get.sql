-- =====================================================================
-- Query   : stock_balance_get
-- Type    : SELECT
-- Purpose : Opening balance of one sparepart in one plant: everything
--           that moved strictly BEFORE p_date. Always returns exactly
--           one row (zeros when nothing moved).
-- Replaces: select_beginning_stock
-- =====================================================================
PARAMETERS
    [p_manufacturing_plant_id] Long,
    [p_sparepart_id] Long,
    [p_date] DateTime;

SELECT
    IIf(Sum(l.quantity_in) IS NULL, 0, Sum(l.quantity_in)) AS quantity_in,
    IIf(Sum(l.quantity_out) IS NULL, 0, Sum(l.quantity_out)) AS quantity_out,
    IIf(Sum(l.quantity_in) IS NULL, 0, Sum(l.quantity_in))
        - IIf(Sum(l.quantity_out) IS NULL, 0, Sum(l.quantity_out)) AS balance
FROM
    vw_stock_ledger AS l
WHERE
    l.manufacturing_plant_id = [p_manufacturing_plant_id]
    AND l.sparepart_id = [p_sparepart_id]
    AND l.movement_date < [p_date];
