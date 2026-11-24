-- =====================================================================
-- Query   : vw_stock_ledger
-- Type    : UNION (view - no parameters, used by other queries)
-- Purpose : Every stock movement of every plant, one row each.
--             IN  = verified OK quantity of an active receive,
--                   dated verification_date
--             OUT = quantity of an active request, dated request_date
--           Source of stock_ledger_get_balance and stock_ledger_list_by_month.
-- Replaces: select_stock_movement
-- =====================================================================
SELECT
	r.manufacturing_plant_id,
	o.sparepart_id,
	r.verification_date AS movement_date,
	'IN' AS direction,
	r.ok_quantity AS quantity_in,
	0 AS quantity_out,
	r.uuid AS reference_uuid,
	v.[name] AS counterparty_name,
	r.remarks,
	r.person_in_charge
FROM
	(vw_receive_active AS r
		INNER JOIN orders AS o
			ON o.id = r.order_id)
		LEFT JOIN vendors AS v
			ON v.id = o.vendor_id
WHERE
	r.verification_date IS NOT NULL

UNION ALL

SELECT
	t.manufacturing_plant_id,
	rq.sparepart_id,
	rq.request_date,
	'OUT',
	0,
	rq.quantity,
	t.uuid,
	mm.[name],
	rq.remarks,
	rq.person_in_charge
FROM
	(requests AS rq
		INNER JOIN vw_transaction_active AS t
			ON t.id = rq.transaction_id)
		INNER JOIN mobile_maintenances AS mm
			ON mm.id = rq.mobile_maintenance_id;
