-- =====================================================================
-- Query   : vw_receive_running_total
-- Type    : SELECT (view - no parameters, used by other queries)
-- Purpose : For each active receive, the OK quantity already received
--           on the same order BEFORE it (by activation time, then id).
--           0 for the first receive of an order.
-- Notes   : Written as a self-join, not a correlated subquery: Access
--           cannot correlate a subquery on a saved query with itself.
-- =====================================================================
SELECT
    r.id,
    Sum(IIf(
        e.activated_at < r.activated_at
        OR (e.activated_at = r.activated_at AND e.id < r.id),
        e.ok_quantity, 0)) AS earlier_ok_quantity
FROM
    vw_receive_active AS r
        INNER JOIN vw_receive_active AS e
            ON e.order_id = r.order_id
GROUP BY
    r.id;
