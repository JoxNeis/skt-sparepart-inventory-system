-- =====================================================================
-- View    : vw_receive_list
-- Purpose : Active receives with their order and everything the
--           receive recap shows and searches (texts for the filter,
--           search_text). remaining_quantity_before = order quantity
--           minus the OK quantity of the receives before this one
--           (never below 0).
-- Used by : receive_list, receive_list_count
-- =====================================================================
SELECT
	r.id AS id,
	r.uuid AS uuid,
	r.manufacturing_plant_id AS manufacturing_plant_id,
	r.arrival_date AS arrival_date,
	r.verification_date AS verification_date,
	r.received_quantity AS received_quantity,
	r.ok_quantity AS ok_quantity,
	r.not_ok_quantity AS not_ok_quantity,
	r.remarks AS remarks,
	r.system_remarks AS system_remarks,
	r.person_in_charge AS person_in_charge,
	r.activated_at AS activated_at,
	IIf(
		o.quantity - rr.earlier_ok_quantity < 0,
		0,
		o.quantity - rr.earlier_ok_quantity
	) AS remaining_quantity_before,
	o.id AS order_id,
	o.uuid AS order_uuid,
	o.po_number AS order_po_number,
	o.order_date AS order_date,
	o.quantity AS order_quantity,
	o.remarks AS order_remarks,
	s.id AS sparepart_id,
	s.code AS sparepart_code,
	s.[name] AS sparepart_name,
	s.unit_of_measurement AS sparepart_unit_of_measurement,
	v.id AS vendor_id,
	v.code AS vendor_code,
	v.[name] AS vendor_name,
	p.[name] AS manufacturing_plant_name,
	s.code & ' - ' & s.[name] AS sparepart_display,
	Format(o.order_date, 'dd-mmm-yyyy') AS order_date_text,
	Format(r.arrival_date, 'dd-mmm-yyyy') AS arrival_date_text,
	Format(r.verification_date, 'dd-mmm-yyyy') AS verification_date_text,
	CStr(IIf(
		o.quantity - rr.earlier_ok_quantity < 0,
		0,
		o.quantity - rr.earlier_ok_quantity
	)) AS remaining_quantity_before_text,
	CStr(r.received_quantity) AS received_quantity_text,
	CStr(r.ok_quantity) AS ok_quantity_text,
	CStr(r.not_ok_quantity) AS not_ok_quantity_text,
	o.uuid & '|' & r.uuid & '|' & o.po_number & '|' & s.code & ' - ' & s.[name]
		& '|' & r.remarks & '|' & r.person_in_charge AS search_text
FROM
	((((vw_receive_active AS r
		INNER JOIN vw_receive_running_total AS rr
			ON rr.id = r.id)
		INNER JOIN vw_order_progress AS o
			ON o.id = r.order_id)
		INNER JOIN spareparts AS s
			ON s.id = o.sparepart_id)
		INNER JOIN manufacturing_plants AS p
			ON p.id = r.manufacturing_plant_id)
		LEFT JOIN vendors AS v
			ON v.id = o.vendor_id;
