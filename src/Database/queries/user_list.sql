-- =====================================================================
-- Query   : user_list
-- Type    : SELECT
-- Purpose : All users that are not archived, with their plant.
-- Replaces: select_users
-- =====================================================================
SELECT
	u.id,
	u.username,
	u.first_name,
	u.middle_name,
	u.last_name,
	u.jobdesk,
	u.tier,
	u.manufacturing_plant_id,
	p.[name] AS manufacturing_plant_name
FROM
	users AS u
		INNER JOIN manufacturing_plants AS p
			ON p.id = u.manufacturing_plant_id
WHERE
	u.deleted_at IS NULL
ORDER BY
	u.username;
