-- =====================================================================
-- Query   : request_create
-- Type    : APPEND
-- Purpose : Add the request detail to a REQUEST transaction header.
--           Inserts nothing (0 rows) when the header is missing or not
--           a REQUEST, the mobile maintenance belongs to another plant
--           or is archived, or the sparepart is archived.
-- Replaces: create_request
-- =====================================================================
PARAMETERS
	[p_uuid] Text(36),
	[p_mobile_maintenance_id] Long,
	[p_sparepart_id] Long,
	[p_request_date] DateTime,
	[p_quantity] Long,
	[p_remarks] LongText,
	[p_system_remarks] LongText,
	[p_person_in_charge] Text(255);

INSERT INTO
	requests
		(
			transaction_id,
			mobile_maintenance_id,
			sparepart_id,
			request_date,
			quantity,
			remarks,
			system_remarks,
			person_in_charge
		)
	SELECT
		t.id,
		mm.id,
		s.id,
		[p_request_date],
		[p_quantity],
		[p_remarks],
		[p_system_remarks],
		[p_person_in_charge]
	FROM
		transactions AS t,
		mobile_maintenances AS mm,
		spareparts AS s
	WHERE
		t.uuid = [p_uuid]
		AND t.transaction_type = 'REQUEST'
		AND mm.id = [p_mobile_maintenance_id]
		AND mm.manufacturing_plant_id = t.manufacturing_plant_id
		AND mm.deleted_at IS NULL
		AND s.id = [p_sparepart_id]
		AND s.deleted_at IS NULL;
