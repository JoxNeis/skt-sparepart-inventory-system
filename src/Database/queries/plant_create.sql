-- =====================================================================
-- Query   : plant_create
-- Type    : APPEND
-- Purpose : Add a plant.
-- Replaces: create_manufacturing_plant
-- =====================================================================
PARAMETERS
    [p_code] Text(15),
    [p_name] Text(255),
    [p_address] Text(255);

INSERT INTO
    manufacturing_plants
        (
            code,
            [name],
            address,
            created_at
        )
    VALUES
        (
            [p_code],
            [p_name],
            [p_address],
            Now()
        );
