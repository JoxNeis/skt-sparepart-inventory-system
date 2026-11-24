-- Query: update_reorder_point
-- Type: UPDATE

PARAMETERS [p_catalogue_name] Text(255), [p_manufacturing_plant_name] Text(255), [p_reorder_point] IEEEDouble, [p_sd_demand] IEEEDouble;
UPDATE (manufacturing_plant_stocks AS mps INNER JOIN catalogues AS c ON c.id = mps.catalogue_id) INNER JOIN manufacturing_plants AS mp ON mp.id = mps.manufacturing_plant_id
SET mps.reorder_point = [p_reorder_point], mps.safety_stock = [p_reorder_point], mps.average_daily_usage = [p_reorder_point], mps.sd_demand = [p_sd_demand]
WHERE c.name = [p_catalogue_name]
    AND mp.name = [p_manufacturing_plant_name];
