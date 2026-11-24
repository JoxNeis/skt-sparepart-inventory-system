-- =====================================================================
-- Query   : manufacturing_plant_stock_create_missing
-- Type    : APPEND
-- Purpose : Give the plant a stock row (all zero) for every catalogue
--           sparepart it does not have yet. Safe to run repeatedly.
-- Replaces: append_manufacturing_plant_stock
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long;

INSERT INTO
	manufacturing_plant_stocks
		(
			manufacturing_plant_id,
			sparepart_id,
			created_at
		)
	SELECT
		p.id AS manufacturing_plant_id,
		s.id AS sparepart_id,
		Now() AS created_at
	FROM
		manufacturing_plants AS p,
		spareparts AS s
	WHERE
		p.id = [p_manufacturing_plant_id]
		AND s.deleted_at IS NULL
		AND NOT EXISTS
			(
				SELECT
					1
				FROM
					manufacturing_plant_stocks AS ps
				WHERE
					ps.manufacturing_plant_id = p.id
					AND ps.sparepart_id = s.id
			);
