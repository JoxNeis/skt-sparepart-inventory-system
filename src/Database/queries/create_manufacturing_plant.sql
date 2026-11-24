-- Query: create_manufacturing_plant
-- Type: APPEND

PARAMETERS [p_manufacturing_plant_name] Text(255), [p_manufacturing_plant_code] Text(255), [p_manufacturing_plant_address] Text(255);
INSERT INTO manufacturing_plants (name, code, address)
VALUES ([p_manufacturing_plant_name], [p_manufacturing_plant_code], [p_manufacturing_plant_address]);
