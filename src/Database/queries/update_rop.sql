-- Query: update_rop
-- Type: UPDATE

PARAMETERS [p_average_daily_usage] IEEEDouble, [p_sd_demand] IEEEDouble, [p_lead_time] IEEEDouble, [p_catalogue_name] Text(255), [p_manufacturing_plant_name] Text(255);
UPDATE (manufacturing_plant_stocks AS mps INNER JOIN catalogues AS c ON mps.catalogue_id = c.id) INNER JOIN manufacturing_plants AS mp ON mps.manufacturing_plant_id = mp.id
SET mps.average_daily_usage = [p_average_daily_usage], mps.sd_demand = [p_sd_demand], mps.lead_time = [p_lead_time], mps.safety_stock = [p_sd_demand]
        * Sqr([p_lead_time])
        * 1.65, mps.reorder_point = (
            [p_sd_demand]
            * Sqr([p_lead_time])
            * 1.65
        )
        +
        (
            [p_average_daily_usage]
            * [p_lead_time]
        )
WHERE c.[name] = [p_catalogue_name]
    AND mp.[name] =
        [p_manufacturing_plant_name];
