-- =====================================================================
-- Query   : manufacturing_plant_stock_set_current_stock
-- Type    : UPDATE
-- Purpose : Overwrite the cached stock with an absolute value
--           (stock take / re-sync from vw_stock_ledger).
--           0 rows affected = no stock row, or p_current_stock < 0.
-- Replaces: update_current_stock_exact
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_sparepart_id] Long,
	[p_current_stock] Long;

UPDATE
	manufacturing_plant_stocks
SET
	current_stock = [p_current_stock],
	updated_at = Now()
WHERE
	manufacturing_plant_id = [p_manufacturing_plant_id]
	AND sparepart_id = [p_sparepart_id]
	AND [p_current_stock] >= 0;
