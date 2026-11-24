-- Query: create_user
-- Type: APPEND

PARAMETERS [p_manufacturing_plant_name] Text(255), [p_username] Text(255), [p_first_name] Text(255), [p_middle_name] Text(255), [p_last_name] Text(255), [p_position] Text(255), [p_tier] Long;
INSERT INTO users (manufacturing_plant_id, username, first_name, middle_name, last_name, position, tier)
SELECT TOP 1 mp.id, [p_username], [p_first_name], [p_middle_name], [p_last_name], [p_position], [p_tier]
FROM manufacturing_plants AS mp
WHERE mp.[name] = [p_manufacturing_plant_name]
ORDER BY mp.id;
