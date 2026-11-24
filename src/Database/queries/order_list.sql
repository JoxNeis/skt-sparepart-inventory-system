-- =====================================================================
-- Query   : order_list
-- Type    : SELECT
-- Purpose : Active orders of a plant dated p_date_from .. p_date_to
--           (whole days, both inclusive), with receive progress.
-- Replaces: select_order_receive_by_manufacturing_plant_name_and_order_date
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long,
	[p_date_from] DateTime,
	[p_date_to] DateTime;

SELECT
	o.id,
	o.uuid,
	o.po_number,
	o.order_date,
	o.quantity,
	o.ok_quantity_total,
	o.received_quantity_total,
	o.remaining_quantity,
	o.status,
	o.remarks,
	o.system_remarks,
	o.person_in_charge,
	o.activated_at,
	s.id AS sparepart_id,
	s.code AS sparepart_code,
	s.[name] AS sparepart_name,
	s.unit_of_measurement AS sparepart_unit_of_measurement,
	v.id AS vendor_id,
	v.[name] AS vendor_name,
	p.[name] AS manufacturing_plant_name
FROM
	((vw_order_progress AS o
		INNER JOIN spareparts AS s
			ON s.id = o.sparepart_id)
		INNER JOIN manufacturing_plants AS p
			ON p.id = o.manufacturing_plant_id)
		LEFT JOIN vendors AS v
			ON v.id = o.vendor_id
WHERE
	o.manufacturing_plant_id = [p_manufacturing_plant_id]
	AND o.order_date >= [p_date_from]
	AND o.order_date < DateAdd('d', 1, [p_date_to])
ORDER BY
	o.order_date,
	o.id;
