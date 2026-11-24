-- Query: create_request
-- Type: APPEND

PARAMETERS [p_transaction_uuid] Text(255), [p_mobile_post_name] Text(255), [p_request_date] DateTime, [p_catalogue_name] Text(255), [p_request_amount] IEEEDouble, [p_request_remarks] LongText, [p_request_system_remarks] LongText, [p_request_person_in_charge] Text(255);
INSERT INTO requests (transaction_id, mobile_post_id, request_date, catalogue_id, amount, remarks, system_remarks, person_in_charge)
SELECT TOP 1 t.id, m.id, [p_request_date], c.id, [p_request_amount], [p_request_remarks], [p_request_system_remarks], [p_request_person_in_charge]
FROM transactions AS t, mobile_posts AS m, catalogues AS c
WHERE t.uuid = [p_transaction_uuid]
    AND m.[name] = [p_mobile_post_name]
    AND c.[name] = [p_catalogue_name]
ORDER BY t.id, m.id, c.id;
