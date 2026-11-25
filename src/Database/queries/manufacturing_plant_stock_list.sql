-- =====================================================================
-- Query   : manufacturing_plant_stock_list
-- Type    : SELECT (one page)
-- Purpose : Stock and ROP planning figures of the catalogue spareparts
--           in a plant (REKAP INVENTORY, dashboard, stock checks).
-- Filter  : p_filter_value searched (contains, ignoring case) in
--           "code - name" and alias; blank = no filter.
-- Paging  : keyset, sorted by (sparepart name, stock row id).
--           p_direction 'NEXT' = rows after  (p_key_text, p_key_id)
--           p_direction 'PREV' = rows before (p_key_text, p_key_id)
--           First page: 'NEXT' from ('', 0).
--           Total rows: manufacturing_plant_stock_count.
-- Replaces: select_plant_stocks
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_filter_value] Text(255),
	[p_direction] Text(4),
	[p_key_text] Text(255),
	[p_key_id] Long;
SELECT
	ps.id AS id,
	ps.current_stock AS current_stock,
	ps.reorder_point AS reorder_point,
	ps.safety_stock AS safety_stock,
	ps.lead_time_days AS lead_time_days,
	ps.average_daily_usage AS average_daily_usage,
	ps.daily_demand_std_dev AS daily_demand_std_dev,
	ps.service_level_z AS service_level_z,
	IIf(
		ps.current_stock <= ps.reorder_point,
		True,
		False
	) AS is_at_reorder_point,
	s.id AS sparepart_id,
	s.code AS sparepart_code,
	s.[name] AS sparepart_name,
	s.unit_of_measurement AS sparepart_unit_of_measurement
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
	)
	AND IIf(
		[p_direction] = 'PREV',
		s.[name] < [p_key_text]
			OR (s.[name] = [p_key_text] AND ps.id < [p_key_id]),
		s.[name] > [p_key_text]
			OR (s.[name] = [p_key_text] AND ps.id > [p_key_id])
	)
ORDER BY
	s.[name],
	ps.id;
