-- Query: create_catalogue
-- Type: APPEND

PARAMETERS [p_name] Text(255), [p_alias] Text(255), [p_code] Text(255), [p_category] Text(255), [p_unit_of_measurement] Text(255);
INSERT INTO catalogues (name, alias_name, code, category, unit_of_measurement)
VALUES ([p_name], [p_alias], [p_code], [p_category], [p_unit_of_measurement]);
