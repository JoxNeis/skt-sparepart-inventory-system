-- Query: select_request_request_update
-- Type: SELECT

PARAMETERS [p_manufacturing_plant_name] Text(255), [p_from_date] DateTime, [p_until_date] DateTime;
SELECT tut.id AS update_request_id, tut.requested_at AS update_request_date, tut.requested_by AS update_request_by, tut.original_transaction_id, ot.uuid AS old_transaction_uuid, ot.actived_at AS old_transaction_actived_at, ot.reversed_at AS old_transaction_reversed_at, orq.request_date AS old_request_date, oc.name AS old_catalogue_name, oc.unit_of_measurement AS old_unit_of_measurement, orq.amount AS old_amount, om.name AS old_mobile_post_name, orq.remarks AS old_remarks, orq.system_remarks AS old_system_remarks, orq.person_in_charge AS old_pic, omp.name AS old_manufacturing_plant_name, tut.proposed_transaction_id, pt.uuid AS new_transaction_uuid, pt.actived_at AS new_transaction_actived_at, pt.reversed_at AS new_transaction_reversed_at, prq.request_date AS new_request_date, pc.name AS new_catalogue_name, pc.unit_of_measurement AS new_unit_of_measurement, prq.amount AS new_amount, pm.name AS new_mobile_post_name, prq.remarks AS new_remarks, prq.system_remarks AS new_system_remarks, prq.person_in_charge AS new_pic, pmp.name AS new_manufacturing_plant_name
FROM (((((((((transaction_update_requests AS tut INNER JOIN transactions AS ot ON tut.original_transaction_id = ot.id) INNER JOIN transactions AS pt ON tut.proposed_transaction_id = pt.id) INNER JOIN requests AS orq ON orq.transaction_id = ot.id) INNER JOIN requests AS prq ON prq.transaction_id = pt.id) INNER JOIN catalogues AS oc ON oc.id = orq.catalogue_id) INNER JOIN catalogues AS pc ON pc.id = prq.catalogue_id) INNER JOIN mobile_posts AS om ON om.id = orq.mobile_post_id) INNER JOIN mobile_posts AS pm ON pm.id = prq.mobile_post_id) INNER JOIN manufacturing_plants AS omp ON omp.id = om.manufacturing_plant_id) INNER JOIN manufacturing_plants AS pmp ON pmp.id = pm.manufacturing_plant_id
WHERE omp.name = [p_manufacturing_plant_name]
    AND orq.request_date >= [p_from_date]
    AND orq.request_date < DateAdd("d", 1, [p_until_date])
    AND tut.approved_at IS NULL
ORDER BY tut.requested_at, ot.uuid;
