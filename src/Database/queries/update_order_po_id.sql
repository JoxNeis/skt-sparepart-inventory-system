-- Query: update_order_po_id
-- Type: UPDATE

PARAMETERS [p_order_uuid] Text(255), [p_order_po_id] Text(255);
UPDATE orders AS o INNER JOIN transactions AS t ON o.transaction_id = t.id
SET o.po_id = [p_order_po_id]
WHERE t.uuid = [p_order_uuid];
