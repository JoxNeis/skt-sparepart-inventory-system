-- =====================================================================
-- Query   : stock_ledger_get_available
-- Type    : SELECT
-- Purpose : The most that can be issued from one sparepart in one plant
--           on p_date without the stock card going negative on p_date
--           or on any later day. That is the lowest end-of-day balance
--           of p_date and of every later day with a movement.
--           current_stock only covers today, so a backdated request
--           must be checked with this query.
--           Always returns one row; available_quantity is Null when
--           the plant does not exist.
--           Replacing a request (update): add the original request's
--           quantity back only when its request_date is on or before
--           p_date (otherwise it did not lower these balances).
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_sparepart_id] Long,
	[p_date] DateTime;

SELECT
	Min(
		IIf(b.balance IS NULL, 0, b.balance)
	) AS available_quantity
FROM
	(
		SELECT
			d.movement_day AS movement_day,
			(
				SELECT
					Sum(l.quantity_in) - Sum(l.quantity_out)
				FROM
					vw_stock_ledger AS l
				WHERE
					l.manufacturing_plant_id = [p_manufacturing_plant_id]
					AND l.sparepart_id = [p_sparepart_id]
					AND l.movement_date < DateAdd('d', 1, d.movement_day)
			) AS balance
		FROM
			(
				SELECT
					DateValue([p_date]) AS movement_day
				FROM
					manufacturing_plants AS p
				WHERE
					p.id = [p_manufacturing_plant_id]

				UNION

				SELECT
					DateValue(x.movement_date)
				FROM
					vw_stock_ledger AS x
				WHERE
					x.manufacturing_plant_id = [p_manufacturing_plant_id]
					AND x.sparepart_id = [p_sparepart_id]
					AND x.movement_date >= DateValue([p_date])
			) AS d
	) AS b;
