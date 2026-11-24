-- Query: select_users
-- Type: SELECT

SELECT u.*, mp.[name] AS manufacturing_plant_name
FROM users AS u INNER JOIN manufacturing_plants AS mp ON u.manufacturing_plant_id = mp.id;
