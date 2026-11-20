-- Order Status query for the NAV SQL Server database.
-- {company} is replaced with [database] company from config.ini (e.g. "JAYTEA CBE").
--
-- MUST return these column names (one row per order line, or per order line per planned week):
--   ord_id, buyer_id, buyer_name, nav_doc_no, buyer_contact_no, shipment_date,
--   line_key, prod_name, pending_qty, stock_in_process, ready_stock, pending_prod,
--   can_pack_cfc, planned_week, planned_qty
--
-- THIS IS A STARTING POINT built on standard NAV tables. Fields marked TODO are custom in your
-- NAV (they come from the Order Status report's own logic) - ask your NAV partner which table/field
-- the report uses, or run:  python -m order_status_monthwise.nav_db --discover
-- Then check with:          python -m order_status_monthwise.nav_db --compare <real ERP export>.xlsx

SELECT
    h.[External Document No_]                     AS ord_id,            -- TODO: field behind "Ord Id"
    h.[Sell-to Customer No_]                      AS buyer_id,
    h.[Sell-to Customer Name]                     AS buyer_name,
    h.[No_]                                       AS nav_doc_no,
    c.[Phone No_]                                 AS buyer_contact_no,
    h.[Requested Delivery Date]                   AS shipment_date,     -- TODO: field behind "Buyer Requested Shipment Date"
    CONCAT(l.[Document No_], '|', l.[Line No_])   AS line_key,
    l.[Description]                               AS prod_name,
    l.[Outstanding Quantity]                      AS pending_qty,
    0                                             AS stock_in_process,  -- TODO
    0                                             AS ready_stock,       -- TODO
    l.[Outstanding Quantity]                      AS pending_prod,      -- TODO: pending qty minus stock?
    0                                             AS can_pack_cfc,      -- TODO
    NULL                                          AS planned_week,      -- TODO: ISO week of planned production
    0                                             AS planned_qty        -- TODO
FROM [{company}$Sales Header] h WITH (NOLOCK)
JOIN [{company}$Sales Line]   l WITH (NOLOCK)
      ON l.[Document Type] = h.[Document Type] AND l.[Document No_] = h.[No_]
LEFT JOIN [{company}$Customer] c WITH (NOLOCK)
      ON c.[No_] = h.[Sell-to Customer No_]
WHERE h.[Document Type] = 1            -- 1 = Order
  AND l.[Type] = 2                     -- 2 = Item
  AND l.[Outstanding Quantity] > 0
ORDER BY h.[Requested Delivery Date], h.[No_], l.[Line No_];

-- Example for planned weeks, if planning lives in production order lines (replace the NULL/0 above
-- and add this join; one row per line per week is fine - the code adds them up):
--   LEFT JOIN [{company}$Prod_ Order Line] p WITH (NOLOCK)
--          ON p.[Item No_] = l.[No_] AND p.[Status] IN (2,3)      -- firm planned / released
--   planned_week = DATEPART(ISO_WEEK, p.[Due Date]), planned_qty = p.[Remaining Quantity]
