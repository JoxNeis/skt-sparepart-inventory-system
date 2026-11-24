-- Query: create_vendor
-- Type: APPEND

PARAMETERS [p_vendor_name] Text(255), [p_vendor_company_name] Text(255), [p_vendor_code] Text(255);
INSERT INTO vendors (name, company_name, code)
VALUES ([p_vendor_name], [p_vendor_company_name], [p_vendor_code]);
