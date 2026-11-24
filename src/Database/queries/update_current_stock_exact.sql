-- Query: update_current_stock_exact
-- Type: UPDATE

PARAMETERS [p_new_stock] IEEEDouble, [p_catalogue_name] Text(255), [p_manufacturing_plant_name] Text(255);
UPDATE (manufacturing_plant_stocks AS mps INNER JOIN catalogues AS c ON c.id = mps.catalogue_id) INNER JOIN manufacturing_plants AS mp ON mp.id = mps.manufacturing_plant_id
SET mps.current_stock = [p_new_stock]
WHERE c.name = [p_catalogue_name]
    AND mp.name = [p_manufacturing_plant_name];
