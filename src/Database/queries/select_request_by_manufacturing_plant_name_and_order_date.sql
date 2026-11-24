-- Query: select_request_by_manufacturing_plant_name_and_order_date
-- Type: SELECT

PARAMETERS [p_manufacturing_plant_name] Text(255), [p_from_date] DateTime, [p_until_date] DateTime;
SELECT t.uuid AS request_uuid, r.request_date AS request_date, c.name AS catalogue_name, c.unit_of_measurement AS catalogue_unit_of_measurement, r.amount AS request_amount, m.name AS mobile_post_name, r.remarks AS request_remarks, r.system_remarks AS request_system_remarks, r.person_in_charge AS request_person_in_charge
FROM (((requests AS r INNER JOIN catalogues AS c ON c.id = r.catalogue_id) INNER JOIN mobile_posts AS m ON m.id = r.mobile_post_id) INNER JOIN manufacturing_plants AS p ON p.id = m.manufacturing_plant_id) INNER JOIN transactions AS t ON r.transaction_id = t.id
WHERE p.name = [p_manufacturing_plant_name]
    AND r.request_date >= [p_from_date]
    AND r.request_date <= [p_until_date]
    AND t.actived_at IS NOT NULL
    AND t.reversed_at IS NULL
ORDER BY r.request_date, t.uuid;
