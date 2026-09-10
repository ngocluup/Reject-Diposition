-- EIMS lot history query (rebuilt from "EIMS lot.VG2" query grid)
-- Source: MARS Oracle, table F_LOT_HISTORY_V3 (f_lothist joined with f_lot)
-- Columns: lot, create_data1..4 (create_data2 = owner NT userid), inqty
-- Filter : lot IN (:lots)  AND  prevout_date >= last 360 days
--
-- The token {{LOT_LIST}} is replaced by build_lot_sql.py with the lots from a CSV
-- (LotNumber column), e.g.  'X629Q396-!1','U550C531'
SELECT
    a0.lot            AS lot,
    a0.create_data1   AS create_data1,
    a0.create_data2   AS create_data2,   -- owner NT userid (e.g. AMR\johndoe)
    a0.create_data3   AS create_data3,   -- Cu contaminant flag
    a0.create_data4   AS create_data4,
    a0.inqty          AS inqty
FROM f_lot_history_v3 a0
WHERE a0.lot IN ({{LOT_LIST}})
  AND a0.prevout_date >= TRUNC(SYSDATE) - 360
ORDER BY a0.lot, a0.prevout_date;
