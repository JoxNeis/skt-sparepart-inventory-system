-- Query: create_receive
-- Type: APPEND

PARAMETERS [p_receive_transaction_uuid] Text(255), [p_order_transaction_uuid] Text(255), [p_arrival_date] DateTime, [p_verification_date] DateTime, [p_receive_total_amount] IEEEDouble, [p_receive_ok_amount] IEEEDouble, [p_receive_not_ok_amount] IEEEDouble, [p_receive_remarks] Text(255), [p_receive_system_remarks] Text(255), [p_receive_person_in_charge] Text(255);
INSERT INTO receives (transaction_id, order_id, arrival_date, verification_date, total_receive_amount, ok_amount, not_ok_amount, remarks, system_remarks, person_in_charge)
SELECT TOP 1 rt.id, o.id, [p_arrival_date], [p_verification_date], [p_receive_total_amount], [p_receive_ok_amount], [p_receive_not_ok_amount], [p_receive_remarks], [p_receive_system_remarks], [p_receive_person_in_charge]
FROM transactions AS rt, transactions AS ot INNER JOIN orders AS o ON ot.id = o.transaction_id
WHERE rt.uuid = [p_receive_transaction_uuid]
    AND ot.uuid = [p_order_transaction_uuid];
