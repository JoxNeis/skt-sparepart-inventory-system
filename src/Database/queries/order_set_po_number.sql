-- =====================================================================
-- Query   : order_set_po_number
-- Type    : UPDATE
-- Purpose : Fill in or correct the SAP PO number of an order. This is
--           the one order field that may change without an update
--           request.
-- Replaces: update_order_po_id
-- =====================================================================
PARAMETERS
	[p_uuid] Text(36),
	[p_po_number] Text(50);

UPDATE
	orders AS o
		INNER JOIN transactions AS t
			ON t.id = o.transaction_id
SET
	o.po_number = [p_po_number]
WHERE
	t.uuid = [p_uuid];
