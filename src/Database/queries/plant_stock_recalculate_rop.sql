-- =====================================================================
-- Query   : plant_stock_recalculate_rop
-- Type    : UPDATE
-- Purpose : Store new demand inputs and recalculate the ROP outputs:
--             safety_stock  = Z * std_dev * Sqr(lead_time)
--             reorder_point = safety_stock + usage * lead_time
--           Z is the row's service_level_z (default 1.65).
--           Results are rounded UP to whole units (-Int(-x) is the
--           ceiling); Round(x, 6) first drops floating-point noise
--           such as 11.000000000000002.
-- Replaces: update_rop
-- =====================================================================
PARAMETERS
    [p_manufacturing_plant_id] Long,
    [p_sparepart_id] Long,
    [p_lead_time_days] Long,
    [p_average_daily_usage] IEEEDouble,
    [p_daily_demand_std_dev] IEEEDouble;

UPDATE
    manufacturing_plant_stocks
SET
    lead_time_days = [p_lead_time_days],
    average_daily_usage = [p_average_daily_usage],
    daily_demand_std_dev = [p_daily_demand_std_dev],
    safety_stock = -Int(-Round(
        service_level_z * [p_daily_demand_std_dev] * Sqr([p_lead_time_days]), 6)),
    reorder_point = -Int(-Round(
        service_level_z * [p_daily_demand_std_dev] * Sqr([p_lead_time_days])
        + [p_average_daily_usage] * [p_lead_time_days], 6)),
    updated_at = Now()
WHERE
    manufacturing_plant_id = [p_manufacturing_plant_id]
    AND sparepart_id = [p_sparepart_id];
