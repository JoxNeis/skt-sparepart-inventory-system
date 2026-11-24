-- Query: create_mobile_post
-- Type: APPEND

PARAMETERS [p_manufacturing_plant_name] Text(255), [p_mobile_post_name] Text(255), [p_mobile_post_code] Text(255);
INSERT INTO mobile_posts (manufacturing_plant_id, name, code)
SELECT TOP 1 p.id, [p_mobile_post_name], [p_mobile_post_code]
FROM manufacturing_plants AS p
WHERE p.[name] = [p_manufacturing_plant_name]
ORDER BY p.id;
