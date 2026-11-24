-- Query: select_stock_movement_by_manufacturing_plant_name_and_order_date
-- Type: UNION

PARAMETERS [p_manufacturing_plant_name] Text(255), [p_catalogue_name] Text(255), [p_month] Short, [p_year] Long;
SELECT
    r.verification_date AS movement_date,
    t.uuid AS incoming_reference_number,
    v.name AS incoming_from,
    r.ok_amount AS incoming_amount,
    '' AS outgoing_reference_number,
    '' AS outgoing_to,
    NULL AS outgoing_amount,
    r.remarks AS remarks,
    r.person_in_charge AS person_in_charge
FROM
    (
        (
            (
                (
                    (
                        receives AS r
                        INNER JOIN transactions AS t ON t.id = r.transaction_id
                    )
                    INNER JOIN orders AS o ON o.id = r.order_id
                )
                INNER JOIN catalogues AS c ON c.id = o.catalogue_id
            )
            INNER JOIN manufacturing_plants AS p ON p.id = o.manufacturing_plant_id
        )
        INNER JOIN vendors AS v ON v.id = o.vendor_id
    )
WHERE
    p.name = [p_manufacturing_plant_name]
    AND c.name = [p_catalogue_name]
    AND Month(r.verification_date) = [p_month]
    AND Year(r.verification_date) = [p_year]
    AND t.reversed_at IS NULL
    AND t.actived_at IS NOT NULL
UNION ALL SELECT
    r.request_date AS movement_date,
    '' AS incoming_reference_number,
    '' AS incoming_from,
    NULL AS incoming_amount,
    t.uuid AS outgoing_reference_number,
    mp.name AS outgoing_to,
    r.amount AS outgoing_amount,
    r.remarks AS remarks,
    r.person_in_charge AS person_in_charge
FROM
    (
        (
            (
                (
                    requests AS r
                    INNER JOIN transactions AS t ON t.id = r.transaction_id
                )
                INNER JOIN catalogues AS c ON c.id = r.catalogue_id
            )
            INNER JOIN mobile_posts AS mp ON mp.id = r.mobile_post_id
        )
        INNER JOIN manufacturing_plants AS p ON p.id = mp.manufacturing_plant_id
    )
WHERE
    p.name = [p_manufacturing_plant_name]
    AND c.name = [p_catalogue_name]
    AND Month(r.request_date) = [p_month]
    AND Year(r.request_date) = [p_year]
    AND t.reversed_at IS NULL
    AND t.actived_at IS NOT NULL
ORDER BY movement_date;
