-- =====================================================================
-- Query   : receive_list_count
-- Type    : SELECT (one row)
-- Purpose : Number of rows receive_list returns over all its pages.
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
	);
