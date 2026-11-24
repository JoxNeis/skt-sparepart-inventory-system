-- Query: select_beginning_stock
-- Type: SELECT

PARAMETERS [p_date] DateTime, [p_manufacturing_plant_name] Text(255), [p_catalogue_name] Text(255);
SELECT c.id AS catalogue_id, c.name AS catalogue_name, Sum(m.total_in) AS total_in, Sum(m.total_out) AS total_out, Sum(m.total_in) - Sum(m.total_out) AS total_stock
FROM catalogues AS c INNER JOIN [SELECT
            o.catalogue_id,
            rc.ok_amount AS total_in,
            0 AS total_out
        FROM
            (
                (
                    receives AS rc
                    INNER JOIN transactions AS rt
                        ON rt.id = rc.transaction_id
                )
                INNER JOIN orders AS o
                    ON o.id = rc.order_id
            )
            INNER JOIN manufacturing_plants AS rp
                ON rp.id = o.manufacturing_plant_id
        WHERE
            rc.verification_date < [p_date]
            AND rp.name = [p_manufacturing_plant_name]
            AND rt.reversed_at IS NULL
            AND rt.actived_at IS NOT NULL

        UNION ALL

        SELECT
            rq.catalogue_id,
            0 AS total_in,
            rq.amount AS total_out
        FROM
            (
                (
                    requests AS rq
                    INNER JOIN transactions AS qt
                        ON qt.id = rq.transaction_id
                )
                INNER JOIN mobile_posts AS mp
                    ON mp.id = rq.mobile_post_id
            )
            INNER JOIN manufacturing_plants AS qp
                ON qp.id = mp.manufacturing_plant_id
        WHERE
            rq.request_date < [p_date]
            AND qp.name = [p_manufacturing_plant_name]
            AND qt.reversed_at IS NULL
            AND qt.actived_at IS NOT NULL
    ]. AS m ON c.id = m.catalogue_id
WHERE c.name = [p_catalogue_name]
GROUP BY c.id, c.name
ORDER BY c.name;
