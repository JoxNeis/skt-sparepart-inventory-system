-- =====================================================================
-- Query   : order_list_pending_count
-- Type    : SELECT (one row)
-- Purpose : Number of rows order_list_pending returns over all its pages (same
--           plant, date range and filter): BARIS DATA / TOTAL HALAMAN.
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
	);
