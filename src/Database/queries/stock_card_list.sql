-- =====================================================================
-- Query   : stock_card_list
-- Type    : SELECT
-- Purpose : Stock card (kartu stok) of one sparepart in one plant for
--           one calendar month: every IN and OUT movement, receipts
--           before issues on the same day.
--           The date range (instead of Month()/Year()) lets Access use
--           the date indexes.
-- Replaces: select_stock_movement_by_manufacturing_plant_name_and_order_date
-- =====================================================================
PARAMETERS
    [p_manufacturing_plant_id] Long,
    [p_sparepart_id] Long,
    [p_year] Short,
    [p_month] Short;

SELECT
    l.movement_date,
    IIf(l.direction = 'IN', l.reference_uuid, '') AS incoming_reference,
    IIf(l.direction = 'IN', l.counterparty_name, '') AS incoming_from,
    IIf(l.direction = 'IN', l.quantity_in, NULL) AS incoming_quantity,
    IIf(l.direction = 'OUT', l.reference_uuid, '') AS outgoing_reference,
    IIf(l.direction = 'OUT', l.counterparty_name, '') AS outgoing_to,
    IIf(l.direction = 'OUT', l.quantity_out, NULL) AS outgoing_quantity,
    l.remarks,
    l.person_in_charge
FROM
    vw_stock_ledger AS l
WHERE
    l.manufacturing_plant_id = [p_manufacturing_plant_id]
    AND l.sparepart_id = [p_sparepart_id]
    AND l.movement_date >= DateSerial([p_year], [p_month], 1)
    AND l.movement_date < DateSerial([p_year], [p_month] + 1, 1)
ORDER BY
    l.movement_date,
    l.direction;
