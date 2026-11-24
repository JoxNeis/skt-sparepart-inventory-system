-- Query: select_order_request_update
-- Type: SELECT

PARAMETERS [p_manufacturing_plant_name] Text(255), [p_from_date] DateTime, [p_until_date] DateTime;
SELECT tut.id AS update_request_id, tut.requested_at AS update_request_date, tut.requested_by AS update_request_by, tut.original_transaction_id, ot.uuid AS old_transaction_uuid, ot.actived_at AS old_transaction_actived_at, ot.reversed_at AS old_transaction_reversed_at, oo.po_id AS old_order_id, oo.order_date AS old_order_date, ov.name AS old_vendor_name, oc.name AS old_catalogue_name, oc.unit_of_measurement AS old_unit_of_measurement, oo.amount AS old_amount, oo.remarks AS old_remarks, oo.person_in_charge AS old_pic, omp.name AS old_manufacturing_plant_name, tut.proposed_transaction_id, pt.uuid AS new_transaction_uuid, pt.actived_at AS new_transaction_actived_at, pt.reversed_at AS new_transaction_reversed_at, po.po_id AS new_order_id, po.order_date AS new_order_date, pv.name AS new_vendor_name, pc.name AS new_catalogue_name, pc.unit_of_measurement AS new_unit_of_measurement, po.amount AS new_amount, po.remarks AS new_remarks, po.person_in_charge AS new_pic, pmp.name AS new_manufacturing_plant_name
FROM (((((((((transaction_update_requests AS tut INNER JOIN transactions AS ot ON tut.original_transaction_id = ot.id) INNER JOIN transactions AS pt ON tut.proposed_transaction_id = pt.id) INNER JOIN orders AS oo ON oo.transaction_id = ot.id) INNER JOIN orders AS po ON po.transaction_id = pt.id) INNER JOIN vendors AS ov ON ov.id = oo.vendor_id) INNER JOIN vendors AS pv ON pv.id = po.vendor_id) INNER JOIN catalogues AS oc ON oc.id = oo.catalogue_id) INNER JOIN catalogues AS pc ON pc.id = po.catalogue_id) INNER JOIN manufacturing_plants AS omp ON omp.id = oo.manufacturing_plant_id) INNER JOIN manufacturing_plants AS pmp ON pmp.id = po.manufacturing_plant_id
WHERE omp.name = [p_manufacturing_plant_name]
    AND oo.order_date >= [p_from_date]
    AND oo.order_date < DateAdd("d", 1, [p_until_date])
    AND tut.approved_at IS NULL
ORDER BY tut.requested_at, ot.uuid;
