-- =====================================================================
-- Query   : transaction_update_request_list_order
-- Type    : SELECT
-- Purpose : Pending update requests on ORDERS of a plant whose original
--           order is dated p_date_from .. p_date_to (whole days), with
--           the old and new values side by side.
-- Replaces: select_order_request_update
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_date_from] DateTime,
	[p_date_to] DateTime;

SELECT
	ur.id,
	ur.reason,
	ur.requested_by,
	ur.requested_at,
	p.[name] AS manufacturing_plant_name,
	old_t.uuid AS old_transaction_uuid,
	old_t.activated_at AS old_transaction_activated_at,
	old_t.reversed_at AS old_transaction_reversed_at,
	old_o.po_number AS old_order_po_number,
	old_o.order_date AS old_order_date,
	old_o.quantity AS old_order_quantity,
	old_o.remarks AS old_order_remarks,
	old_o.person_in_charge AS old_order_person_in_charge,
	old_s.id AS old_sparepart_id,
	old_s.[name] AS old_sparepart_name,
	old_s.unit_of_measurement AS old_sparepart_unit_of_measurement,
	old_v.id AS old_vendor_id,
	old_v.[name] AS old_vendor_name,
	new_t.uuid AS new_transaction_uuid,
	new_t.activated_at AS new_transaction_activated_at,
	new_t.reversed_at AS new_transaction_reversed_at,
	new_o.po_number AS new_order_po_number,
	new_o.order_date AS new_order_date,
	new_o.quantity AS new_order_quantity,
	new_o.remarks AS new_order_remarks,
	new_o.person_in_charge AS new_order_person_in_charge,
	new_s.id AS new_sparepart_id,
	new_s.[name] AS new_sparepart_name,
	new_s.unit_of_measurement AS new_sparepart_unit_of_measurement,
	new_v.id AS new_vendor_id,
	new_v.[name] AS new_vendor_name
FROM
	((((((((transaction_update_requests AS ur
		INNER JOIN transactions AS old_t
			ON old_t.id = ur.original_transaction_id)
		INNER JOIN transactions AS new_t
			ON new_t.id = ur.proposed_transaction_id)
		INNER JOIN orders AS old_o
			ON old_o.transaction_id = old_t.id)
		INNER JOIN orders AS new_o
			ON new_o.transaction_id = new_t.id)
		INNER JOIN spareparts AS old_s
			ON old_s.id = old_o.sparepart_id)
		INNER JOIN spareparts AS new_s
			ON new_s.id = new_o.sparepart_id)
		INNER JOIN manufacturing_plants AS p
			ON p.id = old_t.manufacturing_plant_id)
		LEFT JOIN vendors AS old_v
			ON old_v.id = old_o.vendor_id)
		LEFT JOIN vendors AS new_v
			ON new_v.id = new_o.vendor_id
WHERE
	old_t.manufacturing_plant_id = [p_manufacturing_plant_id]
	AND old_o.order_date >= [p_date_from]
	AND old_o.order_date < DateAdd('d', 1, [p_date_to])
	AND ur.approved_at IS NULL
	AND ur.rejected_at IS NULL
ORDER BY
	ur.requested_at,
	old_t.uuid;
