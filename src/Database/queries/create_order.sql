-- Query: create_order
-- Type: APPEND

PARAMETERS [p_transaction_uuid] Text(255), [p_order_date] DateTime, [p_catalogue_name] Text(255), [p_amount] IEEEDouble, [p_manufacturing_name] Text(255), [p_vendor_name] Text(255), [p_remarks] Text(255), [person_in_charge] Text(255);
INSERT INTO orders (transaction_id, order_date, catalogue_id, amount, manufacturing_plant_id, vendor_id, remarks, person_in_charge)
SELECT TOP 1 t.id, [p_order_date], c.id, [p_amount], p.id, v.id, [p_remarks], [person_in_charge]
FROM transactions AS t, catalogues AS c, manufacturing_plants AS p, vendors AS v
WHERE t.uuid = [p_transaction_uuid]
    AND c.name = [p_catalogue_name]
    AND p.name = [p_manufacturing_name]
    AND v.name = [p_vendor_name]
    AND c.deleted_at IS NULL;
