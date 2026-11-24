-- Query: select_plant_stocks
-- Type: SELECT

PARAMETERS [p_manufacturing_plant_name] Text(255);
SELECT c.[name] AS catalogue_name, c.unit_of_measurement AS catalogue_unit_of_measurement, mps.current_stock AS current_stock, mps.reorder_point AS reorder_point, mps.safety_stock AS safety_stock, mps.lead_time AS lead_time, mps.average_daily_usage AS average_daily_usage, mps.sd_demand AS sd_demand
FROM (manufacturing_plant_stocks AS mps INNER JOIN catalogues AS c ON mps.catalogue_id = c.id) INNER JOIN manufacturing_plants AS mp ON mps.manufacturing_plant_id = mp.id
WHERE mp.[name] = [p_manufacturing_plant_name]
ORDER BY c.[name];
