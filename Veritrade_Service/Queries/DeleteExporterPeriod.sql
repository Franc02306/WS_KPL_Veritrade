DELETE FROM "KPL_VERITRADE_EXPORTERS"
WHERE "Year" = ?
  AND "Month" = ?
  AND "Operation" = ?
  AND "Product" = ?
  AND "TariffCode" = ?;