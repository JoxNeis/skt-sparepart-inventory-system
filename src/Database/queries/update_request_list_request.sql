-- =====================================================================
-- Query   : update_request_list_request
-- Type    : SELECT
-- Purpose : Pending update requests on REQUESTS of a plant whose
--           original request is dated p_date_from .. p_date_to (whole
--           days), with the old and new values side by side.
-- Replaces: select_request_request_update
-- =====================================================================
PARAMETERS
    [p_manufacturing_plant_id] Long,
    [p_date_from] DateTime,
    [p_date_to] DateTime;

SELECT
    ur.id,
    ur.reason,
    ur.requested_by,
    ur.requested_at,
    p.[name] AS manufacturing_plant_name,
    old_t.uuid AS old_transaction_uuid,
    old_t.activated_at AS old_transaction_activated_at,
    old_t.reversed_at AS old_transaction_reversed_at,
    old_rq.request_date AS old_request_date,
    old_rq.quantity AS old_request_quantity,
    old_rq.remarks AS old_request_remarks,
    old_rq.system_remarks AS old_request_system_remarks,
    old_rq.person_in_charge AS old_request_person_in_charge,
    old_s.id AS old_sparepart_id,
    old_s.[name] AS old_sparepart_name,
    old_s.unit_of_measurement AS old_sparepart_unit_of_measurement,
    old_mm.id AS old_mobile_maintenance_id,
    old_mm.[name] AS old_mobile_maintenance_name,
    new_t.uuid AS new_transaction_uuid,
    new_t.activated_at AS new_transaction_activated_at,
    new_t.reversed_at AS new_transaction_reversed_at,
    new_rq.request_date AS new_request_date,
    new_rq.quantity AS new_request_quantity,
    new_rq.remarks AS new_request_remarks,
    new_rq.system_remarks AS new_request_system_remarks,
    new_rq.person_in_charge AS new_request_person_in_charge,
    new_s.id AS new_sparepart_id,
    new_s.[name] AS new_sparepart_name,
    new_s.unit_of_measurement AS new_sparepart_unit_of_measurement,
    new_mm.id AS new_mobile_maintenance_id,
    new_mm.[name] AS new_mobile_maintenance_name
FROM
    ((((((((transaction_update_requests AS ur
        INNER JOIN transactions AS old_t
            ON old_t.id = ur.original_transaction_id)
        INNER JOIN transactions AS new_t
            ON new_t.id = ur.proposed_transaction_id)
        INNER JOIN requests AS old_rq
            ON old_rq.transaction_id = old_t.id)
        INNER JOIN requests AS new_rq
            ON new_rq.transaction_id = new_t.id)
        INNER JOIN spareparts AS old_s
            ON old_s.id = old_rq.sparepart_id)
        INNER JOIN spareparts AS new_s
            ON new_s.id = new_rq.sparepart_id)
        INNER JOIN mobile_maintenances AS old_mm
            ON old_mm.id = old_rq.mobile_maintenance_id)
        INNER JOIN mobile_maintenances AS new_mm
            ON new_mm.id = new_rq.mobile_maintenance_id)
        INNER JOIN manufacturing_plants AS p
            ON p.id = old_t.manufacturing_plant_id
WHERE
    old_t.manufacturing_plant_id = [p_manufacturing_plant_id]
    AND old_rq.request_date >= [p_date_from]
    AND old_rq.request_date < DateAdd('d', 1, [p_date_to])
    AND ur.approved_at IS NULL
    AND ur.rejected_at IS NULL
ORDER BY
    ur.requested_at,
    old_t.uuid;
