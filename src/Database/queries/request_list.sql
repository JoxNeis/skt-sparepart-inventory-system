-- =====================================================================
-- Query   : request_list
-- Type    : SELECT
-- Purpose : Active requests of a plant dated p_date_from .. p_date_to
--           (whole days, both inclusive).
-- Replaces: select_request_by_manufacturing_plant_name_and_order_date
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_date_from] DateTime,
	[p_date_to] DateTime;

SELECT
	rq.id AS id,
	t.uuid AS uuid,
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
	mm.[name] AS mobile_maintenance_name
FROM
	((requests AS rq
		INNER JOIN vw_transaction_active AS t
			ON t.id = rq.transaction_id)
		INNER JOIN spareparts AS s
			ON s.id = rq.sparepart_id)
		INNER JOIN mobile_maintenances AS mm
			ON mm.id = rq.mobile_maintenance_id
WHERE
	t.manufacturing_plant_id = [p_manufacturing_plant_id]
	AND rq.request_date >= [p_date_from]
	AND rq.request_date < DateAdd('d', 1, [p_date_to])
ORDER BY
	rq.request_date,
	t.uuid;
