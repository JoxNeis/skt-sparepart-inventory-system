-- Query: append_manufacturing_plant_stock
-- Type: APPEND

PARAMETERS [p_manufacturing_plant_name] Text(255);
INSERT INTO manufacturing_plant_stocks (manufacturing_plant_id, catalogue_id)
SELECT p.id AS manufacturing_plant_id, c.id AS catalogue_id
FROM manufacturing_plants AS p, catalogues AS c
WHERE p.name = [p_manufacturing_plant_name]
    AND NOT EXISTS
    (
        SELECT
            s.catalogue_id
        FROM
            manufacturing_plant_stocks AS s
        WHERE
            s.manufacturing_plant_id = p.id
            AND s.catalogue_id = c.id
    );
