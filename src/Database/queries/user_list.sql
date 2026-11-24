-- =====================================================================
-- Query   : user_list
-- Type    : SELECT
-- Purpose : All users that are not archived, with their plant.
-- Replaces: select_users
-- =====================================================================
SELECT
	u.id AS id,
	u.username AS username,
	u.first_name AS first_name,
	u.middle_name AS middle_name,
	u.last_name AS last_name,
	u.jobdesk AS jobdesk,
	u.tier AS tier,
	u.manufacturing_plant_id AS manufacturing_plant_id,
	p.[name] AS manufacturing_plant_name
FROM
	users AS u
		INNER JOIN manufacturing_plants AS p
			ON p.id = u.manufacturing_plant_id
WHERE
	u.deleted_at IS NULL
ORDER BY
	u.username;
