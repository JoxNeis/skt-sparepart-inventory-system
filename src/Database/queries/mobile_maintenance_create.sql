-- =====================================================================
-- Query   : mobile_maintenance_create
-- Type    : APPEND
-- Purpose : Add a mobile maintenance to a plant.
-- Replaces: create_mobile_post
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_code] Text(40),
	[p_name] Text(255);

INSERT INTO
	mobile_maintenances
		(
			manufacturing_plant_id,
			code,
			[name],
			created_at
		)
	VALUES
		(
			[p_manufacturing_plant_id],
			[p_code],
			[p_name],
			Now()
		);
