-- =====================================================================
-- Query   : stock_ledger_count_by_month
-- Type    : SELECT (one row)
-- Purpose : Number of rows stock_ledger_list_by_month returns over all
--           its pages.
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_sparepart_id] Long,
	[p_year] Short,
	[p_month] Short;
SELECT
	Count(*) AS total_rows
FROM
	vw_stock_ledger AS l
WHERE
	l.manufacturing_plant_id = [p_manufacturing_plant_id]
	AND l.sparepart_id = [p_sparepart_id]
	AND l.movement_date >= DateSerial([p_year], [p_month], 1)
	AND l.movement_date < DateSerial([p_year], [p_month] + 1, 1);
