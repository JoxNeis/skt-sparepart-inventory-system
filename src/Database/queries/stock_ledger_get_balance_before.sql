-- =====================================================================
-- Query   : stock_ledger_get_balance_before
-- Type    : SELECT (one row)
-- Purpose : Stock of one sparepart in one plant just before one stock
--           card row, in the order of stock_ledger_list_by_month
--           (movement_date, direction, reference): the STOK AWAL of
--           the first row of a stock card page.
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_sparepart_id] Long,
	[p_key_date] DateTime,
	[p_key_direction] Text(3),
	[p_key_reference] Text(36);
SELECT
	IIf(Sum(l.quantity_in) IS NULL, 0, Sum(l.quantity_in))
	- IIf(Sum(l.quantity_out) IS NULL, 0, Sum(l.quantity_out)) AS balance
FROM
	vw_stock_ledger AS l
WHERE
	l.manufacturing_plant_id = [p_manufacturing_plant_id]
	AND l.sparepart_id = [p_sparepart_id]
	AND (
		l.movement_date < [p_key_date]
		OR (l.movement_date = [p_key_date]
			AND (l.direction < [p_key_direction]
				OR (l.direction = [p_key_direction]
					AND l.reference_uuid < [p_key_reference])))
	);
