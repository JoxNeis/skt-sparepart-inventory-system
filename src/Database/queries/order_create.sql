-- =====================================================================
-- Query   : order_create
-- Type    : APPEND
-- Purpose : Add the order detail to an ORDER transaction header.
--           Inserts nothing (0 rows) when the header is missing, is not
--           an ORDER, or the sparepart is archived.
-- Replaces: create_order
-- =====================================================================
PARAMETERS
	[p_uuid] Text(36),
	[p_sparepart_id] Long,
	[p_vendor_id] Long,
	[p_po_number] Text(50),
	[p_order_date] DateTime,
	[p_quantity] Long,
	[p_remarks] LongText,
	[p_system_remarks] LongText,
	[p_person_in_charge] Text(255);

INSERT INTO
	orders
		(
			transaction_id,
			sparepart_id,
			vendor_id,
			po_number,
			order_date,
			quantity,
			remarks,
			system_remarks,
			person_in_charge
		)
	SELECT
		t.id AS transaction_id,
		s.id AS sparepart_id,
		[p_vendor_id] AS vendor_id,
		[p_po_number] AS po_number,
		[p_order_date] AS order_date,
		[p_quantity] AS quantity,
		[p_remarks] AS remarks,
		[p_system_remarks] AS system_remarks,
		[p_person_in_charge] AS person_in_charge
	FROM
		transactions AS t,
		spareparts AS s
	WHERE
		t.uuid = [p_uuid]
		AND t.transaction_type = 'ORDER'
		AND s.id = [p_sparepart_id]
		AND s.deleted_at IS NULL;
