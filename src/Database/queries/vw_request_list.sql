-- =====================================================================
-- View    : vw_request_list
-- Purpose : Active requests with everything the request recap shows
--           and searches (texts for the filter, search_text).
-- Used by : request_list, request_list_count
-- =====================================================================
SELECT
	rq.id AS id,
	t.uuid AS uuid,
	t.manufacturing_plant_id AS manufacturing_plant_id,
	rq.request_date AS request_date,
	rq.quantity AS quantity,
	rq.remarks AS remarks,
	rq.system_remarks AS system_remarks,
	rq.person_in_charge AS person_in_charge,
	t.activated_at AS activated_at,
	s.id AS sparepart_id,
	s.code AS sparepart_code,
	s.[name] AS sparepart_name,
	s.unit_of_measurement AS sparepart_unit_of_measurement,
	mm.id AS mobile_maintenance_id,
	mm.code AS mobile_maintenance_code,
	mm.[name] AS mobile_maintenance_name,
	s.code & ' - ' & s.[name] AS sparepart_display,
	mm.code & ' - ' & mm.[name] AS mobile_maintenance_display,
	CStr(rq.quantity) AS quantity_text,
	Format(rq.request_date, 'dd-mmm-yyyy') AS request_date_text,
	t.uuid & '|' & s.code & ' - ' & s.[name] & '|' & mm.code & ' - ' & mm.[name]
		& '|' & rq.remarks & '|' & rq.person_in_charge AS search_text
FROM
	((requests AS rq
		INNER JOIN vw_transaction_active AS t
			ON t.id = rq.transaction_id)
		INNER JOIN spareparts AS s
			ON s.id = rq.sparepart_id)
		INNER JOIN mobile_maintenances AS mm
			ON mm.id = rq.mobile_maintenance_id;
