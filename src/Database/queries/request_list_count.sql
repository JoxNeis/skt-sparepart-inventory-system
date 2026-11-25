-- =====================================================================
-- Query   : request_list_count
-- Type    : SELECT (one row)
-- Purpose : Number of rows request_list returns over all its pages.
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_date_from] DateTime,
	[p_date_to] DateTime,
	[p_filter_index] Long,
	[p_filter_value] Text(255);
SELECT
	Count(*) AS total_rows
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
	);
