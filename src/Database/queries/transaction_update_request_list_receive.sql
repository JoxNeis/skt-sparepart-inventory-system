-- =====================================================================
-- Query   : transaction_update_request_list_receive
-- Type    : SELECT
-- Purpose : Pending update requests on RECEIVES of a plant whose
--           original receive arrived p_date_from .. p_date_to (whole
--           days), with the old and new values side by side.
-- Replaces: select_receive_request_update
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
	old_r.id AS old_receive_id,
	old_r.arrival_date AS old_receive_arrival_date,
	old_r.verification_date AS old_receive_verification_date,
	old_r.received_quantity AS old_receive_received_quantity,
	old_r.ok_quantity AS old_receive_ok_quantity,
	old_r.not_ok_quantity AS old_receive_not_ok_quantity,
	old_r.remarks AS old_receive_remarks,
	old_r.system_remarks AS old_receive_system_remarks,
	old_r.person_in_charge AS old_receive_person_in_charge,
	old_ot.uuid AS old_order_uuid,
	old_o.id AS old_order_id,
	old_o.po_number AS old_order_po_number,
	old_o.order_date AS old_order_date,
	old_o.quantity AS old_order_quantity,
	old_s.id AS old_sparepart_id,
	old_s.[name] AS old_sparepart_name,
	old_s.unit_of_measurement AS old_sparepart_unit_of_measurement,
	old_v.id AS old_vendor_id,
	old_v.[name] AS old_vendor_name,
	new_t.uuid AS new_transaction_uuid,
	new_t.activated_at AS new_transaction_activated_at,
	new_t.reversed_at AS new_transaction_reversed_at,
	new_r.id AS new_receive_id,
	new_r.arrival_date AS new_receive_arrival_date,
	new_r.verification_date AS new_receive_verification_date,
	new_r.received_quantity AS new_receive_received_quantity,
	new_r.ok_quantity AS new_receive_ok_quantity,
	new_r.not_ok_quantity AS new_receive_not_ok_quantity,
	new_r.remarks AS new_receive_remarks,
	new_r.system_remarks AS new_receive_system_remarks,
	new_r.person_in_charge AS new_receive_person_in_charge,
	new_ot.uuid AS new_order_uuid,
	new_o.id AS new_order_id,
	new_o.po_number AS new_order_po_number,
	new_o.order_date AS new_order_date,
	new_o.quantity AS new_order_quantity,
	new_s.id AS new_sparepart_id,
	new_s.[name] AS new_sparepart_name,
	new_s.unit_of_measurement AS new_sparepart_unit_of_measurement,
	new_v.id AS new_vendor_id,
	new_v.[name] AS new_vendor_name
FROM
	((((((((((((transaction_update_requests AS ur
		INNER JOIN transactions AS old_t
			ON old_t.id = ur.original_transaction_id)
		INNER JOIN transactions AS new_t
			ON new_t.id = ur.proposed_transaction_id)
		INNER JOIN receives AS old_r
			ON old_r.transaction_id = old_t.id)
		INNER JOIN receives AS new_r
			ON new_r.transaction_id = new_t.id)
		INNER JOIN orders AS old_o
			ON old_o.id = old_r.order_id)
		INNER JOIN orders AS new_o
			ON new_o.id = new_r.order_id)
		INNER JOIN transactions AS old_ot
			ON old_ot.id = old_o.transaction_id)
		INNER JOIN transactions AS new_ot
			ON new_ot.id = new_o.transaction_id)
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
	AND old_r.arrival_date >= [p_date_from]
	AND old_r.arrival_date < DateAdd('d', 1, [p_date_to])
	AND ur.approved_at IS NULL
	AND ur.rejected_at IS NULL
ORDER BY
	ur.requested_at,
	old_t.uuid;
