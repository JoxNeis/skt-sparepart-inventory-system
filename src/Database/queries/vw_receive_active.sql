-- =====================================================================
-- Query   : vw_receive_active
-- Type    : SELECT (view - no parameters, used by other queries)
-- Purpose : Receives whose transaction is active, with the header's
--           document number, plant and activation time.
-- =====================================================================
SELECT
    r.id,
    r.transaction_id,
    r.order_id,
    t.uuid,
    t.manufacturing_plant_id,
    t.activated_at,
    r.arrival_date,
    r.verification_date,
    r.received_quantity,
    r.ok_quantity,
    r.not_ok_quantity,
    r.remarks,
    r.system_remarks,
    r.person_in_charge
FROM
    receives AS r
        INNER JOIN vw_transaction_active AS t
            ON t.id = r.transaction_id;
