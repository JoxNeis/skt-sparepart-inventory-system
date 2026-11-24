-- Query: update_receive_order_id
-- Type: UPDATE

PARAMETERS [p_old_order_uuid] Text(255), [p_new_order_uuid] Text(255);
UPDATE (((receives AS r INNER JOIN orders AS old_o ON r.order_id = old_o.id) INNER JOIN transactions AS old_t ON old_o.transaction_id = old_t.id) INNER JOIN orders AS new_o ON (old_o.manufacturing_plant_id =
                new_o.manufacturing_plant_id) AND (old_o.vendor_id = new_o.vendor_id) AND (old_o.catalogue_id = new_o.catalogue_id)) INNER JOIN transactions AS new_t ON new_o.transaction_id = new_t.id
SET r.order_id = new_o.id
WHERE old_t.uuid = [p_old_order_uuid]
    AND new_t.uuid = [p_new_order_uuid];
