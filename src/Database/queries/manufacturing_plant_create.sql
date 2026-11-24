-- =====================================================================
-- Query   : manufacturing_plant_create
-- Type    : APPEND
-- Purpose : Add a plant.
-- Replaces: create_manufacturing_plant
-- =====================================================================
PARAMETERS
	[p_code] Text(15),
	[p_name] Text(255),
	[p_address] Text(255),
	[p_type] Text(50),
	[p_region] Text(100);

INSERT INTO
	manufacturing_plants
		(
			code,
			[name],
			address,
			[type],
			region,
			created_at
		)
	VALUES
		(
			[p_code],
			[p_name],
			[p_address],
			[p_type],
			[p_region],
			Now()
		);
