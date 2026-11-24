-- =====================================================================
-- Query   : plant_update
-- Type    : UPDATE
-- Purpose : Overwrite every editable column of one plant.
--           0 rows affected = not found or archived.
-- =====================================================================
PARAMETERS
    [p_manufacturing_plant_id] Long,
    [p_code] Text(15),
    [p_name] Text(255),
    [p_address] Text(255);

UPDATE
    manufacturing_plants
SET
    code = [p_code],
    [name] = [p_name],
    address = [p_address],
    updated_at = Now()
WHERE
    id = [p_manufacturing_plant_id]
    AND deleted_at IS NULL;
