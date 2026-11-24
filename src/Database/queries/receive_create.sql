-- =====================================================================
-- Query   : receive_create
-- Type    : APPEND
-- Purpose : Add the receive detail to a RECEIVE transaction header,
--           against the order with document number p_order_uuid.
--           Inserts nothing (0 rows) when either header is missing,
--           the types are wrong, or the two belong to different plants.
-- Replaces: create_receive
-- =====================================================================
PARAMETERS
	[p_uuid] Text(36),
	[p_order_uuid] Text(36),
	[p_arrival_date] DateTime,
	[p_verification_date] DateTime,
	[p_received_quantity] Long,
	[p_ok_quantity] Long,
	[p_not_ok_quantity] Long,
	[p_remarks] LongText,
	[p_system_remarks] LongText,
	[p_person_in_charge] Text(255);

INSERT INTO
	receives
		(
			transaction_id,
			order_id,
			arrival_date,
			verification_date,
			received_quantity,
			ok_quantity,
			not_ok_quantity,
			remarks,
			system_remarks,
			person_in_charge
		)
	SELECT
		rt.id,
		o.id,
		[p_arrival_date],
		[p_verification_date],
		[p_received_quantity],
		[p_ok_quantity],
		[p_not_ok_quantity],
		[p_remarks],
		[p_system_remarks],
		[p_person_in_charge]
	FROM
		transactions AS rt,
		transactions AS ot
			INNER JOIN orders AS o
				ON o.transaction_id = ot.id
	WHERE
		rt.uuid = [p_uuid]
		AND rt.transaction_type = 'RECEIVE'
		AND ot.uuid = [p_order_uuid]
		AND ot.manufacturing_plant_id = rt.manufacturing_plant_id;
