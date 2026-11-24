-- =====================================================================
-- Query   : vendor_create
-- Type    : APPEND
-- Purpose : Add a vendor.
-- Replaces: create_vendor
-- =====================================================================
PARAMETERS
	[p_code] Text(40),
	[p_name] Text(255),
	[p_company_name] Text(255),
	[p_email] Text(255);

INSERT INTO
	vendors
		(
			code,
			[name],
			company_name,
			email,
			created_at
		)
	VALUES
		(
			[p_code],
			[p_name],
			[p_company_name],
			[p_email],
			Now()
		);
