-- =====================================================================
-- Query   : vw_receive_active
-- Type    : SELECT (view - no parameters, used by other queries)
-- Purpose : Receives whose transaction is active, with the header's
--           document number, plant and activation time.
-- =====================================================================
SELECT
	r.id AS id,
	r.transaction_id AS transaction_id,
	r.order_id AS order_id,
	t.uuid AS uuid,
	t.manufacturing_plant_id AS manufacturing_plant_id,
	t.activated_at AS activated_at,
	r.arrival_date AS arrival_date,
	r.verification_date AS verification_date,
	r.received_quantity AS received_quantity,
	r.ok_quantity AS ok_quantity,
	r.not_ok_quantity AS not_ok_quantity,
	r.remarks AS remarks,
	r.system_remarks AS system_remarks,
	r.person_in_charge AS person_in_charge
FROM
	receives AS r
		INNER JOIN vw_transaction_active AS t
			ON t.id = r.transaction_id;
