-- =====================================================================
-- Query   : stock_ledger_list_by_month
-- Type    : SELECT (one page)
-- Purpose : Stock card (kartu stok) of one sparepart in one plant for
--           one calendar month: every IN and OUT movement, receipts
--           before issues on the same day.
--           The date range (instead of Month()/Year()) lets Access use
--           the date indexes.
-- Paging  : keyset, sorted by (movement_date, direction, reference).
--           p_direction 'NEXT' = rows after  the key
--           p_direction 'PREV' = rows before the key
--           First page: 'NEXT' from (1900-01-01, '', '').
--           Total rows: stock_ledger_count_by_month. STOK AWAL of the
--           first row of a page: stock_ledger_get_balance_before.
-- Replaces: select_stock_movement_by_manufacturing_plant_name_and_order_date
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_sparepart_id] Long,
	[p_year] Short,
	[p_month] Short,
	[p_direction] Text(4),
	[p_key_date] DateTime,
	[p_key_direction] Text(3),
	[p_key_reference] Text(36);
SELECT
	l.movement_date AS movement_date,
	l.direction AS direction,
	l.reference_uuid AS reference_uuid,
	IIf(l.direction = 'IN', l.reference_uuid, '') AS incoming_reference,
	IIf(l.direction = 'IN', l.counterparty_name, '') AS incoming_from,
	IIf(l.direction = 'IN', l.quantity_in, NULL) AS incoming_quantity,
	IIf(l.direction = 'OUT', l.reference_uuid, '') AS outgoing_reference,
	IIf(l.direction = 'OUT', l.counterparty_name, '') AS outgoing_to,
	IIf(l.direction = 'OUT', l.quantity_out, NULL) AS outgoing_quantity,
	l.remarks AS remarks,
	l.person_in_charge AS person_in_charge
FROM
	vw_stock_ledger AS l
WHERE
	l.manufacturing_plant_id = [p_manufacturing_plant_id]
	AND l.sparepart_id = [p_sparepart_id]
	AND l.movement_date >= DateSerial([p_year], [p_month], 1)
	AND l.movement_date < DateSerial([p_year], [p_month] + 1, 1)
	AND IIf(
		[p_direction] = 'PREV',
		l.movement_date < [p_key_date]
			OR (l.movement_date = [p_key_date]
				AND (l.direction < [p_key_direction]
					OR (l.direction = [p_key_direction]
						AND l.reference_uuid < [p_key_reference]))),
		l.movement_date > [p_key_date]
			OR (l.movement_date = [p_key_date]
				AND (l.direction > [p_key_direction]
					OR (l.direction = [p_key_direction]
						AND l.reference_uuid > [p_key_reference])))
	)
ORDER BY
	l.movement_date,
	l.direction,
	l.reference_uuid;
