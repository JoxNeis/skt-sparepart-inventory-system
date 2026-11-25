-- =====================================================================
-- Query   : manufacturing_plant_stock_count
-- Type    : SELECT (one row)
-- Purpose : Number of rows manufacturing_plant_stock_list returns over
--           all its pages.
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_filter_value] Text(255);
SELECT
	Count(*) AS total_rows
FROM
	manufacturing_plant_stocks AS ps
		INNER JOIN spareparts AS s
			ON s.id = ps.sparepart_id
WHERE
	ps.manufacturing_plant_id = [p_manufacturing_plant_id]
	AND s.deleted_at IS NULL
	AND (
		[p_filter_value] IS NULL
		OR InStr(1, s.code & ' - ' & s.[name] & '|' & s.alias, [p_filter_value]) > 0
	);
