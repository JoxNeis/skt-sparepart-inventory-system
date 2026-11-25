-- =====================================================================
-- View    : vw_order_list
-- Purpose : Active orders with everything the order recaps show and
--           search: sparepart / vendor as "code - name", dates and
--           numbers as text, and search_text (all searchable columns
--           in one text) for the list filters.
--           estimated_arrival = order_date + lead_time_days of the
--           plant stock row (Null when the lead time is 0).
-- Used by : order_list, order_list_count, order_list_pending,
--           order_list_pending_count
-- =====================================================================
SELECT
	o.id AS id,
	o.uuid AS uuid,
	o.manufacturing_plant_id AS manufacturing_plant_id,
	o.po_number AS po_number,
	o.order_date AS order_date,
	o.quantity AS quantity,
	o.ok_quantity_total AS ok_quantity_total,
	o.received_quantity_total AS received_quantity_total,
	o.remaining_quantity AS remaining_quantity,
	o.status AS [status],
	o.remarks AS remarks,
	o.system_remarks AS system_remarks,
	o.person_in_charge AS person_in_charge,
	o.activated_at AS activated_at,
	IIf(
		ps.lead_time_days > 0,
		DateAdd('d', ps.lead_time_days, o.order_date),
		Null
	) AS estimated_arrival,
	s.id AS sparepart_id,
	s.code AS sparepart_code,
	s.[name] AS sparepart_name,
	s.unit_of_measurement AS sparepart_unit_of_measurement,
	v.id AS vendor_id,
	v.code AS vendor_code,
	v.[name] AS vendor_name,
	p.[name] AS manufacturing_plant_name,
	s.code & ' - ' & s.[name] AS sparepart_display,
	IIf(v.id IS NULL, Null, v.code & ' - ' & v.[name]) AS vendor_display,
	CStr(o.quantity) AS quantity_text,
	CStr(o.ok_quantity_total) AS ok_quantity_text,
	CStr(o.remaining_quantity) AS remaining_quantity_text,
	Format(o.order_date, 'dd-mmm-yyyy') AS order_date_text,
	IIf(
		ps.lead_time_days > 0,
		Format(DateAdd('d', ps.lead_time_days, o.order_date), 'dd-mmm-yyyy'),
		Null
	) AS estimated_arrival_text,
	o.uuid & '|' & o.po_number & '|' & s.code & ' - ' & s.[name]
		& '|' & v.code & ' - ' & v.[name] & '|' & o.status
		& '|' & o.remarks & '|' & o.person_in_charge AS search_text
FROM
	(((vw_order_progress AS o
		INNER JOIN spareparts AS s
			ON s.id = o.sparepart_id)
		INNER JOIN manufacturing_plants AS p
			ON p.id = o.manufacturing_plant_id)
		LEFT JOIN vendors AS v
			ON v.id = o.vendor_id)
		LEFT JOIN manufacturing_plant_stocks AS ps
			ON (ps.manufacturing_plant_id = o.manufacturing_plant_id
				AND ps.sparepart_id = o.sparepart_id);
