-- =====================================================================
-- Query   : vw_order_receive_total
-- Type    : SELECT (view - no parameters, used by other queries)
-- Purpose : Per order, the total of its active receives (0 when none).
-- =====================================================================
SELECT
	o.id AS order_id,
	Sum(
		IIf(r.ok_quantity IS NULL, 0, r.ok_quantity)
	) AS ok_quantity_total,
	Sum(
		IIf(r.received_quantity IS NULL, 0, r.received_quantity)
	) AS received_quantity_total
FROM
	orders AS o
		LEFT JOIN vw_receive_active AS r
			ON r.order_id = o.id
GROUP BY
	o.id;
