-- =====================================================================
-- Query   : receive_reassign_order
-- Type    : UPDATE
-- Purpose : After an order edit is approved, move EVERY receive of the
--           old order (p_old_order_uuid) to the new order
--           (p_new_order_uuid).
--           Both orders must be in the same plant (joined on the
--           header's manufacturing_plant_id); otherwise 0 rows.
-- Replaces: update_receive_order_id
-- =====================================================================
PARAMETERS
	[p_old_order_uuid] Text(36),
	[p_new_order_uuid] Text(36);

UPDATE
	(((receives AS r
		INNER JOIN orders AS old_o
			ON old_o.id = r.order_id)
		INNER JOIN transactions AS old_t
			ON old_t.id = old_o.transaction_id)
		INNER JOIN transactions AS new_t
			ON new_t.manufacturing_plant_id = old_t.manufacturing_plant_id)
		INNER JOIN orders AS new_o
			ON new_o.transaction_id = new_t.id
SET
	r.order_id = new_o.id
WHERE
	old_t.uuid = [p_old_order_uuid]
	AND new_t.uuid = [p_new_order_uuid];
