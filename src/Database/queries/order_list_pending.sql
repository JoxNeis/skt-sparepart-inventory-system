-- =====================================================================
-- Query   : order_list_pending
-- Type    : SELECT (one page)
-- Purpose : Same as order_list, only orders with a remaining quantity
--           (REKAP PENDING ORDER, INPUT PENERIMAAN).
-- Filter  : p_filter_value is searched (contains, ignoring case) in
--           the column number p_filter_index (list below); 0 = every
--           column (search_text); blank value = no filter.
-- Paging  : keyset. p_direction 'NEXT' = rows after (p_key_date,
--           p_key_id), 'PREV' = rows before it; first page = 'NEXT'
--           from (1900-01-01, 0). The DAO reads the first page-size
--           rows (NEXT) or the last ones (PREV) of the result.
--           Columns for p_filter_index:
--           1 uuid, 2 po_number, 3 sparepart (code - name),
--           4 quantity, 5 unit, 6 vendor (code - name), 7 order date,
--           8 estimated arrival, 9 status, 10 remarks,
--           11 person in charge, 12 OK received, 13 remaining.
--           Sorted by (order_date, id). Total rows: order_list_pending_count.
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_date_from] DateTime,
	[p_date_to] DateTime,
	[p_filter_index] Long,
	[p_filter_value] Text(255),
	[p_direction] Text(4),
	[p_key_date] DateTime,
	[p_key_id] Long;
SELECT
	l.id AS id,
	l.uuid AS uuid,
	l.po_number AS po_number,
	l.order_date AS order_date,
	l.quantity AS quantity,
	l.ok_quantity_total AS ok_quantity_total,
	l.received_quantity_total AS received_quantity_total,
	l.remaining_quantity AS remaining_quantity,
	l.status AS [status],
	l.remarks AS remarks,
	l.system_remarks AS system_remarks,
	l.person_in_charge AS person_in_charge,
	l.activated_at AS activated_at,
	l.estimated_arrival AS estimated_arrival,
	l.sparepart_id AS sparepart_id,
	l.sparepart_code AS sparepart_code,
	l.sparepart_name AS sparepart_name,
	l.sparepart_unit_of_measurement AS sparepart_unit_of_measurement,
	l.vendor_id AS vendor_id,
	l.vendor_code AS vendor_code,
	l.vendor_name AS vendor_name,
	l.manufacturing_plant_name AS manufacturing_plant_name
FROM
	vw_order_list AS l
WHERE
	l.manufacturing_plant_id = [p_manufacturing_plant_id]
	AND l.order_date >= [p_date_from]
	AND l.order_date < DateAdd('d', 1, [p_date_to])
	AND l.quantity > 0
	AND l.remaining_quantity > 0
	AND (
		[p_filter_value] IS NULL
		OR InStr(1,
			IIf([p_filter_index] = 0,
				l.search_text,
				Choose([p_filter_index],
					l.uuid,
					l.po_number,
					l.sparepart_display,
					l.quantity_text,
					l.sparepart_unit_of_measurement,
					l.vendor_display,
					l.order_date_text,
					l.estimated_arrival_text,
					l.status,
					l.remarks,
					l.person_in_charge,
					l.ok_quantity_text,
					l.remaining_quantity_text)),
			[p_filter_value]) > 0
	)
	AND IIf(
		[p_direction] = 'PREV',
		l.order_date < [p_key_date]
			OR (l.order_date = [p_key_date] AND l.id < [p_key_id]),
		l.order_date > [p_key_date]
			OR (l.order_date = [p_key_date] AND l.id > [p_key_id])
	)
ORDER BY
	l.order_date,
	l.id;
