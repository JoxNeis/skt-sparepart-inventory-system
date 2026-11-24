-- =====================================================================
-- Query   : sparepart_create
-- Type    : APPEND
-- Purpose : Add a sparepart to the catalogue.
-- Replaces: create_catalogue
-- =====================================================================
PARAMETERS
	[p_code] Text(40),
	[p_name] Text(255),
	[p_category] Text(50),
	[p_applicable] Text(255),
	[p_unit_of_measurement] Text(15),
	[p_alias] Text(255),
	[p_status] Text(50);

INSERT INTO
	spareparts
		(
			code,
			[name],
			category,
			applicable,
			unit_of_measurement,
			alias,
			status,
			created_at
		)
	VALUES
		(
			[p_code],
			[p_name],
			[p_category],
			[p_applicable],
			[p_unit_of_measurement],
			[p_alias],
			[p_status],
			Now()
		);
