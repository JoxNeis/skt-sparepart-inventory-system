-- Query: select_pending_order_by_manufacturing_plant_name_and_order_date
-- Type: SELECT

PARAMETERS [p_manufacturing_plant_name] Text(255), [p_from_date] DateTime, [p_until_date] DateTime;
SELECT o.id AS order_id, t.uuid AS order_uuid, o.po_id AS order_poid, p.name AS manufacturing_plant_name, v.name AS vendor_name, c.name AS catalogue_name, c.unit_of_measurement AS catalogue_unit_of_measurement, o.amount AS order_amount, Sum(
        IIf(
            rt.actived_at IS NOT NULL
            AND rt.reversed_at IS NULL,
            IIf(r.ok_amount IS NULL, 0, r.ok_amount),
            0
        )
    ) AS total_ok_amount, IIf(
        o.amount - Sum(
            IIf(
                rt.actived_at IS NOT NULL
                AND rt.reversed_at IS NULL,
                IIf(r.ok_amount IS NULL, 0, r.ok_amount),
                0
            )
        ) < 0,
        0,
        o.amount - Sum(
            IIf(
                rt.actived_at IS NOT NULL
                AND rt.reversed_at IS NULL,
                IIf(r.ok_amount IS NULL, 0, r.ok_amount),
                0
            )
        )
    ) AS remaining_order_amount, o.order_date AS order_date, IIf(
        Sum(
            IIf(
                rt.actived_at IS NOT NULL
                AND rt.reversed_at IS NULL,
                IIf(
                    r.total_receive_amount IS NULL,
                    0,
                    r.total_receive_amount
                ),
                0
            )
        ) = 0,
        "BELUM DATANG",
        IIf(
            Sum(
                IIf(
                    rt.actived_at IS NOT NULL
                    AND rt.reversed_at IS NULL,
                    IIf(r.ok_amount IS NULL, 0, r.ok_amount),
                    0
                )
            ) >= o.amount,
            "DATANG",
            "DATANG SEBAGIAN"
        )
    ) AS order_status, o.remarks AS order_remarks, o.person_in_charge AS order_person_in_charge, t.actived_at AS actived_at
FROM (((((transactions AS t INNER JOIN orders AS o ON t.id = o.transaction_id) INNER JOIN catalogues AS c ON o.catalogue_id = c.id) INNER JOIN vendors AS v ON o.vendor_id = v.id) INNER JOIN manufacturing_plants AS p ON o.manufacturing_plant_id = p.id) LEFT JOIN receives AS r ON o.id = r.order_id) LEFT JOIN transactions AS rt ON r.transaction_id = rt.id
WHERE p.name = [p_manufacturing_plant_name]
    AND o.order_date >= [p_from_date]
    AND o.order_date < DateAdd("d", 1, [p_until_date])
    AND t.actived_at IS NOT NULL
    AND t.reversed_at IS NULL
GROUP BY o.id, t.uuid, o.po_id, p.name, v.name, c.name, c.unit_of_measurement, o.amount, o.order_date, o.remarks, o.person_in_charge, t.actived_at
HAVING Sum(
        IIf(
            rt.actived_at IS NOT NULL
            AND rt.reversed_at IS NULL,
            IIf(r.ok_amount IS NULL, 0, r.ok_amount),
            0
        )
    ) < o.amount
ORDER BY o.order_date, o.id;
