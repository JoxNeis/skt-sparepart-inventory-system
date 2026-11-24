-- =====================================================================
-- Query   : receive_reassign_order
-- Type    : UPDATE
-- Purpose : After an order edit is approved, move the receives of the
--           old order to the new order - only when both orders are for
--           the same sparepart, vendor and plant.
--           orders.vendor_id is optional: two orders without a vendor
--           count as the same vendor.
-- Replaces: update_receive_order_id
-- =====================================================================
PARAMETERS
    [p_old_order_uuid] Text(36),
    [p_new_order_uuid] Text(36);

UPDATE
    (((receives AS r
        INNER JOIN orders AS old_o
            ON old_o.id = r.order_id)
        INNER JOIN transactions AS old_t
            ON old_t.id = old_o.transaction_id)
        INNER JOIN orders AS new_o
            ON new_o.sparepart_id = old_o.sparepart_id)
        INNER JOIN transactions AS new_t
            ON new_t.id = new_o.transaction_id
SET
    r.order_id = new_o.id
WHERE
    old_t.uuid = [p_old_order_uuid]
    AND new_t.uuid = [p_new_order_uuid]
    AND new_t.manufacturing_plant_id = old_t.manufacturing_plant_id
    AND (
            new_o.vendor_id = old_o.vendor_id
            OR (new_o.vendor_id IS NULL AND old_o.vendor_id IS NULL)
        );
