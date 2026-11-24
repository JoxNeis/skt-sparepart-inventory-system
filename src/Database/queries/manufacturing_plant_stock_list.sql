-- =====================================================================
-- Query   : manufacturing_plant_stock_list
-- Type    : SELECT
-- Purpose : Stock and ROP planning figures of every catalogue
--           sparepart in a plant.
-- Replaces: select_plant_stocks
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long;

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
ORDER BY
	s.[name];
