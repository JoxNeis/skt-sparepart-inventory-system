-- =====================================================================
-- Query   : manufacturing_plant_stock_adjust_current_stock
-- Type    : UPDATE
-- Purpose : Add a signed quantity to the cached stock
--           (+ receive, - request, or the difference of a reversal).
--           0 rows affected = no stock row, or the result would be
--           negative (ck_manufacturing_plant_stocks_current_stocks).
-- Replaces: update_current_stock
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_sparepart_id] Long,
	[p_quantity_change] Long;

UPDATE
	manufacturing_plant_stocks
SET
	current_stock = current_stock + [p_quantity_change],
	updated_at = Now()
WHERE
	manufacturing_plant_id = [p_manufacturing_plant_id]
	AND sparepart_id = [p_sparepart_id]
	AND current_stock + [p_quantity_change] >= 0;
