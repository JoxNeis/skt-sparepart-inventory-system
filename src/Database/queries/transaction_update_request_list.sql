-- =====================================================================
-- Query   : transaction_update_request_list
-- Type    : SELECT
-- Purpose : Update requests of a plant REQUESTED p_date_from ..
--           p_date_to (whole days, both inclusive), for every document
--           type, with the old and new values side by side.
--           Pending requests have verdict IS NULL.
--           One row shape for all types (UNION ALL); a column that does
--           not apply to a type is NULL:
--             document_date    ORDER: order_date
--                              RECEIVE: arrival_date
--                              REQUEST: request_date
--             quantity         ORDER, REQUEST: quantity
--                              RECEIVE: received_quantity
--             order_uuid, po_number, vendor_name
--                              ORDER: its own
--                              RECEIVE: those of the received order
--             ok_quantity, not_ok_quantity   RECEIVE only
--             mobile_maintenance_name        REQUEST only
-- Replaces: transaction_update_request_list_order,
--           transaction_update_request_list_receive,
--           transaction_update_request_list_request
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_date_from] DateTime,
	[p_date_to] DateTime;

SELECT
	ur.id AS id,
	old_t.transaction_type AS transaction_type,
	ur.reason AS reason,
	ur.requested_by AS requested_by,
	ur.requested_at AS requested_at,
	ur.verdict AS verdict,
	ur.reviewer AS reviewer,
	ur.reviewed_at AS reviewed_at,
	ur.decision_note AS decision_note,
	p.[name] AS manufacturing_plant_name,
	old_t.uuid AS old_transaction_uuid,
	old_r.arrival_date AS old_document_date,
	old_r.received_quantity AS old_quantity,
	old_r.ok_quantity AS old_ok_quantity,
	old_r.not_ok_quantity AS old_not_ok_quantity,
	old_r.remarks AS old_remarks,
	old_r.person_in_charge AS old_person_in_charge,
	old_ot.uuid AS old_order_uuid,
	old_o.po_number AS old_po_number,
	old_s.id AS old_sparepart_id,
	old_s.[name] AS old_sparepart_name,
	old_s.unit_of_measurement AS old_sparepart_unit_of_measurement,
	old_v.[name] AS old_vendor_name,
	IIf(False, '', Null) AS old_mobile_maintenance_name,
	new_t.uuid AS new_transaction_uuid,
	new_r.arrival_date AS new_document_date,
	new_r.received_quantity AS new_quantity,
	new_r.ok_quantity AS new_ok_quantity,
	new_r.not_ok_quantity AS new_not_ok_quantity,
	new_r.remarks AS new_remarks,
	new_r.person_in_charge AS new_person_in_charge,
	new_ot.uuid AS new_order_uuid,
	new_o.po_number AS new_po_number,
	new_s.id AS new_sparepart_id,
	new_s.[name] AS new_sparepart_name,
	new_s.unit_of_measurement AS new_sparepart_unit_of_measurement,
	new_v.[name] AS new_vendor_name,
	IIf(False, '', Null) AS new_mobile_maintenance_name
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
	AND ur.requested_at >= [p_date_from]
	AND ur.requested_at < DateAdd('d', 1, [p_date_to])

UNION ALL

SELECT
	ur.id,
	old_t.transaction_type,
	ur.reason,
	ur.requested_by,
	ur.requested_at,
	ur.verdict,
	ur.reviewer,
	ur.reviewed_at,
	ur.decision_note,
	p.[name],
	old_t.uuid,
	old_o.order_date,
	old_o.quantity,
	Null,
	Null,
	old_o.remarks,
	old_o.person_in_charge,
	old_t.uuid,
	old_o.po_number,
	old_s.id,
	old_s.[name],
	old_s.unit_of_measurement,
	old_v.[name],
	Null,
	new_t.uuid,
	new_o.order_date,
	new_o.quantity,
	Null,
	Null,
	new_o.remarks,
	new_o.person_in_charge,
	new_t.uuid,
	new_o.po_number,
	new_s.id,
	new_s.[name],
	new_s.unit_of_measurement,
	new_v.[name],
	Null
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
	AND ur.requested_at >= [p_date_from]
	AND ur.requested_at < DateAdd('d', 1, [p_date_to])

UNION ALL

SELECT
	ur.id,
	old_t.transaction_type,
	ur.reason,
	ur.requested_by,
	ur.requested_at,
	ur.verdict,
	ur.reviewer,
	ur.reviewed_at,
	ur.decision_note,
	p.[name],
	old_t.uuid,
	old_rq.request_date,
	old_rq.quantity,
	Null,
	Null,
	old_rq.remarks,
	old_rq.person_in_charge,
	Null,
	Null,
	old_s.id,
	old_s.[name],
	old_s.unit_of_measurement,
	Null,
	old_mm.[name],
	new_t.uuid,
	new_rq.request_date,
	new_rq.quantity,
	Null,
	Null,
	new_rq.remarks,
	new_rq.person_in_charge,
	Null,
	Null,
	new_s.id,
	new_s.[name],
	new_s.unit_of_measurement,
	Null,
	new_mm.[name]
FROM
	((((((((transaction_update_requests AS ur
		INNER JOIN transactions AS old_t
			ON old_t.id = ur.original_transaction_id)
		INNER JOIN transactions AS new_t
			ON new_t.id = ur.proposed_transaction_id)
		INNER JOIN requests AS old_rq
			ON old_rq.transaction_id = old_t.id)
		INNER JOIN requests AS new_rq
			ON new_rq.transaction_id = new_t.id)
		INNER JOIN spareparts AS old_s
			ON old_s.id = old_rq.sparepart_id)
		INNER JOIN spareparts AS new_s
			ON new_s.id = new_rq.sparepart_id)
		INNER JOIN mobile_maintenances AS old_mm
			ON old_mm.id = old_rq.mobile_maintenance_id)
		INNER JOIN mobile_maintenances AS new_mm
			ON new_mm.id = new_rq.mobile_maintenance_id)
		INNER JOIN manufacturing_plants AS p
			ON p.id = old_t.manufacturing_plant_id
WHERE
	old_t.manufacturing_plant_id = [p_manufacturing_plant_id]
	AND ur.requested_at >= [p_date_from]
	AND ur.requested_at < DateAdd('d', 1, [p_date_to])

ORDER BY
	requested_at,
	old_transaction_uuid;
