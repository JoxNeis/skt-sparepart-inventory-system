-- Query: select_mobile_post_by_manufacturing_plant_name
-- Type: SELECT

PARAMETERS [p_manufacturing_plant_name] Text(255);
SELECT m.name
FROM mobile_posts AS m INNER JOIN manufacturing_plants AS p ON m.manufacturing_plant_id = p.id
WHERE p.name = [p_manufacturing_plant_name]
ORDER BY m.name;
