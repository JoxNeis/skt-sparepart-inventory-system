-- =====================================================================
-- Query   : manufacturing_plant_update
-- Type    : UPDATE
-- Purpose : Overwrite every editable column of one plant.
--           0 rows affected = not found, archived, or p_code differs
--           from the current code while the plant already has a
--           transaction (the code is part of every document number).
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_code] Text(15),
	[p_name] Text(255),
	[p_address] Text(255),
	[p_type] Text(50),
	[p_region] Text(100);

UPDATE
	manufacturing_plants
SET
	code = [p_code],
	[name] = [p_name],
	address = [p_address],
	[type] = [p_type],
	region = [p_region],
	updated_at = Now()
WHERE
	id = [p_manufacturing_plant_id]
	AND deleted_at IS NULL
	AND (
		code = [p_code]
		OR NOT EXISTS
			(
				SELECT
					1
				FROM
					transactions AS t
				WHERE
					t.manufacturing_plant_id = [p_manufacturing_plant_id]
			)
	);
