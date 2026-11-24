-- =====================================================================
-- Query   : vw_order_progress
-- Type    : SELECT (view - no parameters, used by other queries)
-- Purpose : Active orders with what has been received so far, the
--           remaining quantity and the delivery status.
--             BELUM DATANG    nothing received yet
--             DATANG SEBAGIAN received, OK quantity < ordered
--             DATANG          OK quantity >= ordered
-- =====================================================================
SELECT
    o.id,
    o.transaction_id,
    t.uuid,
    t.manufacturing_plant_id,
    t.activated_at,
    o.sparepart_id,
    o.vendor_id,
    o.po_number,
    o.order_date,
    o.quantity,
    o.remarks,
    o.system_remarks,
    o.person_in_charge,
    rt.ok_quantity_total,
    rt.received_quantity_total,
    IIf(o.quantity - rt.ok_quantity_total < 0, 0,
        o.quantity - rt.ok_quantity_total) AS remaining_quantity,
    IIf(rt.received_quantity_total = 0, 'BELUM DATANG',
        IIf(rt.ok_quantity_total >= o.quantity, 'DATANG', 'DATANG SEBAGIAN')) AS status
FROM
    (orders AS o
        INNER JOIN vw_transaction_active AS t
            ON t.id = o.transaction_id)
        INNER JOIN vw_order_receipt_total AS rt
            ON rt.order_id = o.id;
