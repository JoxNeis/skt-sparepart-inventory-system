-- Query: select_receive_by_manufacturing_plant_name_and_arrival_date
-- Type: SELECT

PARAMETERS [p_manufacturing_plant_name] Text(255), [p_from_date] DateTime, [p_until_date] DateTime;
SELECT r.id AS receive_id, t.uuid AS transaction_uuid, o.id AS order_id, ot.uuid AS order_uuid, o.po_id AS order_poid, p.name AS manufacturing_plant_name, v.name AS vendor_name, c.name AS catalogue_name, c.unit_of_measurement AS catalogue_unit_of_measurement, o.amount AS order_amount, o.order_date AS order_date, o.remarks AS order_remarks, r.ok_amount AS total_ok_amount, r.not_ok_amount AS total_not_ok_amount, r.total_receive_amount AS total_receive_amount, IIf(
        o.amount
        - Sum(
            IIf(
                other_t.id IS NOT NULL
                AND other_t.actived_at IS NOT NULL
                AND other_t.reversed_at IS NULL
                AND (
                    other_t.actived_at < t.actived_at
                    OR (
                        other_t.actived_at = t.actived_at
                        AND other_r.id < r.id
                    )
                ),
                IIf(
                    other_r.ok_amount IS NULL,
                    0,
                    other_r.ok_amount
                ),
                0
            )
        ) < 0,
        0,
        o.amount
        - Sum(
            IIf(
                other_t.id IS NOT NULL
                AND other_t.actived_at IS NOT NULL
                AND other_t.reversed_at IS NULL
                AND (
                    other_t.actived_at < t.actived_at
                    OR (
                        other_t.actived_at = t.actived_at
                        AND other_r.id < r.id
                    )
                ),
                IIf(
                    other_r.ok_amount IS NULL,
                    0,
                    other_r.ok_amount
                ),
                0
            )
        )
    ) AS total_remaining_order, r.arrival_date AS receive_arrival_date, r.verification_date AS receive_verification_date, r.remarks AS receive_remarks, r.system_remarks AS receive_system_remarks, r.person_in_charge AS receive_person_in_charge
FROM (((((((transactions AS t INNER JOIN receives AS r ON t.id = r.transaction_id) INNER JOIN orders AS o ON r.order_id = o.id) INNER JOIN transactions AS ot ON o.transaction_id = ot.id) INNER JOIN catalogues AS c ON o.catalogue_id = c.id) INNER JOIN vendors AS v ON o.vendor_id = v.id) INNER JOIN manufacturing_plants AS p ON o.manufacturing_plant_id = p.id) LEFT JOIN receives AS other_r ON o.id = other_r.order_id) LEFT JOIN transactions AS other_t ON other_r.transaction_id = other_t.id
WHERE p.name = [p_manufacturing_plant_name]
    AND o.order_date >= [p_from_date]
    AND o.order_date <= [p_until_date]
    AND t.actived_at IS NOT NULL
    AND t.reversed_at IS NULL
    AND ot.actived_at IS NOT NULL
    AND ot.reversed_at IS NULL
GROUP BY r.id, r.ok_amount, r.not_ok_amount, r.total_receive_amount, t.uuid, t.actived_at, o.id, ot.uuid, o.po_id, p.name, v.name, c.name, c.unit_of_measurement, o.amount, o.order_date, o.remarks, r.arrival_date, r.verification_date, r.remarks, r.system_remarks, r.person_in_charge
ORDER BY o.order_date, o.id, t.actived_at, r.id;
