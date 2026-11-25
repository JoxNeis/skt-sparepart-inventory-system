-- =====================================================================
-- Query   : request_list
-- Type    : SELECT (one page)
-- Purpose : Active, not deleted (quantity > 0) requests of a plant
--           dated p_date_from .. p_date_to (whole days, both
--           inclusive) (REKAP REQUEST).
-- Filter  : p_filter_value is searched (contains, ignoring case) in
--           the column number p_filter_index (list below); 0 = every
--           column (search_text); blank value = no filter.
-- Paging  : keyset. p_direction 'NEXT' = rows after (p_key_date,
--           p_key_id), 'PREV' = rows before it; first page = 'NEXT'
--           from (1900-01-01, 0). The DAO reads the first page-size
--           rows (NEXT) or the last ones (PREV) of the result.
--           Columns for p_filter_index:
--           1 uuid, 2 sparepart (code - name), 3 quantity, 4 unit,
--           5 mobile post (code - name), 6 request date, 7 remarks,
--           8 person in charge.
--           Sorted by (request_date, id). Total rows: request_list_count.
-- Replaces: select_request_by_manufacturing_plant_name_and_order_date
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
	l.request_date AS request_date,
	l.quantity AS quantity,
	l.remarks AS remarks,
	l.system_remarks AS system_remarks,
	l.person_in_charge AS person_in_charge,
	l.activated_at AS activated_at,
	l.sparepart_id AS sparepart_id,
	l.sparepart_code AS sparepart_code,
	l.sparepart_name AS sparepart_name,
	l.sparepart_unit_of_measurement AS sparepart_unit_of_measurement,
	l.mobile_maintenance_id AS mobile_maintenance_id,
	l.mobile_maintenance_code AS mobile_maintenance_code,
	l.mobile_maintenance_name AS mobile_maintenance_name
FROM
	vw_request_list AS l
WHERE
	l.manufacturing_plant_id = [p_manufacturing_plant_id]
	AND l.request_date >= [p_date_from]
	AND l.request_date < DateAdd('d', 1, [p_date_to])
	AND l.quantity > 0
	AND (
		[p_filter_value] IS NULL
		OR InStr(1,
			IIf([p_filter_index] = 0,
				l.search_text,
				Choose([p_filter_index],
					l.uuid,
					l.sparepart_display,
					l.quantity_text,
					l.sparepart_unit_of_measurement,
					l.mobile_maintenance_display,
					l.request_date_text,
					l.remarks,
					l.person_in_charge)),
			[p_filter_value]) > 0
	)
	AND IIf(
		[p_direction] = 'PREV',
		l.request_date < [p_key_date]
			OR (l.request_date = [p_key_date] AND l.id < [p_key_id]),
		l.request_date > [p_key_date]
			OR (l.request_date = [p_key_date] AND l.id > [p_key_id])
	)
ORDER BY
	l.request_date,
	l.id;
