-- =====================================================================
-- Query   : user_create
-- Type    : APPEND
-- Purpose : Add a user.
-- Replaces: create_user
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_username] Text(64),
	[p_first_name] Text(100),
	[p_middle_name] Text(100),
	[p_last_name] Text(100),
	[p_jobdesk] Text(255),
	[p_tier] Long;

INSERT INTO
	users
		(
			manufacturing_plant_id,
			username,
			first_name,
			middle_name,
			last_name,
			jobdesk,
			tier,
			created_at
		)
	VALUES
		(
			[p_manufacturing_plant_id],
			[p_username],
			[p_first_name],
			[p_middle_name],
			[p_last_name],
			[p_jobdesk],
			[p_tier],
			Now()
		);
