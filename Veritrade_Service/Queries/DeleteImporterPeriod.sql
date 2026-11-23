DELETE FROM "KPL_VERITRADE_IMPORTERS"
WHERE "Year" = ?
  AND "Month" = ?
  AND "Operation" = ?
  AND "Product" = ?
  AND "TariffCode" = ?;