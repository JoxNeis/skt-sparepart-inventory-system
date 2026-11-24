-- =====================================================================
-- Query   : manufacturing_plant_stock_update
-- Type    : UPDATE
-- Purpose : Overwrite every planning column of one plant stock row
--           by hand. current_stock is not included on purpose; use
--           manufacturing_plant_stock_adjust_current_stock / _set_current_stock.
-- Replaces: update_reorder_point (it wrote p_reorder_point into
--           safety_stock and average_daily_usage as well)
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_sparepart_id] Long,
	[p_reorder_point] Long,
	[p_safety_stock] Long,
	[p_lead_time_days] Long,
	[p_average_daily_usage] IEEEDouble,
	[p_daily_demand_std_dev] IEEEDouble,
	[p_service_level_z] IEEEDouble;

UPDATE
	manufacturing_plant_stocks
SET
	reorder_point = [p_reorder_point],
	safety_stock = [p_safety_stock],
	lead_time_days = [p_lead_time_days],
	average_daily_usage = [p_average_daily_usage],
	daily_demand_std_dev = [p_daily_demand_std_dev],
	service_level_z = [p_service_level_z],
	updated_at = Now()
WHERE
	manufacturing_plant_id = [p_manufacturing_plant_id]
	AND sparepart_id = [p_sparepart_id];
