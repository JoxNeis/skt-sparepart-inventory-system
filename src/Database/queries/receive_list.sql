-- =====================================================================
-- Query   : receive_list
-- Type    : SELECT (one page)
-- Purpose : Active, not deleted (received_quantity > 0) receives of a
--           plant that ARRIVED p_date_from .. p_date_to (whole days,
--           both inclusive), with their order (REKAP PENERIMAAN).
-- Filter  : p_filter_value is searched (contains, ignoring case) in
--           the column number p_filter_index (list below); 0 = every
--           column (search_text); blank value = no filter.
-- Paging  : keyset. p_direction 'NEXT' = rows after (p_key_date,
--           p_key_id), 'PREV' = rows before it; first page = 'NEXT'
--           from (1900-01-01, 0). The DAO reads the first page-size
--           rows (NEXT) or the last ones (PREV) of the result.
--           Columns for p_filter_index:
--           1 order uuid, 2 receive uuid, 3 po_number,
--           4 sparepart (code - name), 5 order date, 6 arrival date,
--           7 verification date, 8 remaining before, 9 received,
--           10 OK, 11 NOT OK, 12 unit, 13 remarks,
--           14 person in charge.
--           Sorted by (arrival_date, id). Total rows: receive_list_count.
-- Replaces: select_receive_by_manufacturing_plant_name_and_arrival_date
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
	l.arrival_date AS arrival_date,
	l.verification_date AS verification_date,
	l.received_quantity AS received_quantity,
	l.ok_quantity AS ok_quantity,
	l.not_ok_quantity AS not_ok_quantity,
	l.remarks AS remarks,
	l.system_remarks AS system_remarks,
	l.person_in_charge AS person_in_charge,
	l.activated_at AS activated_at,
	l.remaining_quantity_before AS remaining_quantity_before,
	l.order_id AS order_id,
	l.order_uuid AS order_uuid,
	l.order_po_number AS order_po_number,
	l.order_date AS order_date,
	l.order_quantity AS order_quantity,
	l.order_remarks AS order_remarks,
	l.sparepart_id AS sparepart_id,
	l.sparepart_code AS sparepart_code,
	l.sparepart_name AS sparepart_name,
	l.sparepart_unit_of_measurement AS sparepart_unit_of_measurement,
	l.vendor_id AS vendor_id,
	l.vendor_code AS vendor_code,
	l.vendor_name AS vendor_name,
	l.manufacturing_plant_name AS manufacturing_plant_name
FROM
	vw_receive_list AS l
WHERE
	l.manufacturing_plant_id = [p_manufacturing_plant_id]
	AND l.arrival_date >= [p_date_from]
	AND l.arrival_date < DateAdd('d', 1, [p_date_to])
	AND l.received_quantity > 0
	AND (
		[p_filter_value] IS NULL
		OR InStr(1,
			IIf([p_filter_index] = 0,
				l.search_text,
				Choose([p_filter_index],
					l.order_uuid,
					l.uuid,
					l.order_po_number,
					l.sparepart_display,
					l.order_date_text,
					l.arrival_date_text,
					l.verification_date_text,
					l.remaining_quantity_before_text,
					l.received_quantity_text,
					l.ok_quantity_text,
					l.not_ok_quantity_text,
					l.sparepart_unit_of_measurement,
					l.remarks,
					l.person_in_charge)),
			[p_filter_value]) > 0
	)
	AND IIf(
		[p_direction] = 'PREV',
		l.arrival_date < [p_key_date]
			OR (l.arrival_date = [p_key_date] AND l.id < [p_key_id]),
		l.arrival_date > [p_key_date]
			OR (l.arrival_date = [p_key_date] AND l.id > [p_key_id])
	)
ORDER BY
	l.arrival_date,
	l.id;
