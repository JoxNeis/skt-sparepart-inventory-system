-- Query: select_stock_movement
-- Type: UNION

PARAMETERS [p_date] DateTime, [p_manufacturing_plant_name] Text(255);
SELECT
    o.catalogue_id,
    rc.ok_amount AS total_in,
    0 AS total_out
FROM
    (((receives AS rc
    INNER JOIN transactions AS t
        ON t.id = rc.transaction_id)
    INNER JOIN orders AS o
        ON o.id = rc.order_id)
    INNER JOIN manufacturing_plants AS p
        ON p.id = o.manufacturing_plant_id)
WHERE
    rc.verification_date < [p_date]
    AND p.name = [p_manufacturing_plant_name]
    AND t.reversed_at IS NULL
    AND t.actived_at IS NOT NULL
UNION ALL SELECT
    rq.catalogue_id,
    0 AS total_in,
    rq.amount AS total_out
FROM
    (((requests AS rq
    INNER JOIN transactions AS t
        ON t.id = rq.transaction_id)
    INNER JOIN mobile_posts AS mp
        ON mp.id = rq.mobile_post_id)
    INNER JOIN manufacturing_plants AS p
        ON p.id = mp.manufacturing_plant_id)
WHERE
    rq.request_date < [p_date]
    AND p.name = [p_manufacturing_plant_name]
    AND t.reversed_at IS NULL
    AND t.actived_at IS NOT NULL;
