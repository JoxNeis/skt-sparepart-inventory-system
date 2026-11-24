-- =====================================================================
-- Query   : receive_list
-- Type    : SELECT
-- Purpose : Active receives of a plant that ARRIVED p_date_from ..
--           p_date_to (whole days, both inclusive), with their order.
--           remaining_quantity_before = order quantity minus the OK
--           quantity of the receives before this one (never below 0).
-- Replaces: select_receive_by_manufacturing_plant_name_and_arrival_date
--           (which despite its name filtered on the ORDER date)
-- =====================================================================
PARAMETERS
    [p_manufacturing_plant_id] Long,
    [p_date_from] DateTime,
    [p_date_to] DateTime;

SELECT
    r.id,
    r.uuid,
    r.arrival_date,
    r.verification_date,
    r.received_quantity,
    r.ok_quantity,
    r.not_ok_quantity,
    r.remarks,
    r.system_remarks,
    r.person_in_charge,
    r.activated_at,
    IIf(o.quantity - rr.earlier_ok_quantity < 0, 0,
        o.quantity - rr.earlier_ok_quantity) AS remaining_quantity_before,
    o.id AS order_id,
    o.uuid AS order_uuid,
    o.po_number AS order_po_number,
    o.order_date,
    o.quantity AS order_quantity,
    o.remarks AS order_remarks,
    s.id AS sparepart_id,
    s.code AS sparepart_code,
    s.[name] AS sparepart_name,
    s.unit_of_measurement AS sparepart_unit_of_measurement,
    v.id AS vendor_id,
    v.[name] AS vendor_name,
    p.[name] AS manufacturing_plant_name
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
            ON v.id = o.vendor_id
WHERE
    r.manufacturing_plant_id = [p_manufacturing_plant_id]
    AND r.arrival_date >= [p_date_from]
    AND r.arrival_date < DateAdd('d', 1, [p_date_to])
ORDER BY
    o.order_date,
    o.id,
    r.activated_at,
    r.id;
