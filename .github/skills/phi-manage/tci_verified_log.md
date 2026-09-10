# TCI PHI Tracking — Verified Examples Log

Regression/proof records extracted from SKILL.md to keep the skill lean. **Read on demand** when you need a known-good reference count to validate output. These are evidence, not rules — the rules live in SKILL.md.

---

## Filter row counts (Class / PPV / Burn-In / yield)

- ARL Refresh S 8C+16A+GT1 SDA Monthly + Class → 148 TEST_* + 3 monitors = **151 rows**.
- ARL Refresh HX 8C+16A+GT1 SDA Monthly + PPV → 32 PPV_* + 3 `PPV-M SAMPLE SIZE` = **35 rows**.
- PTL U404 SDA Monthly + PPV → 32 PPV_* + 3 `PPV-M SAMPLE SIZE` = **35 rows** (user-confirmed 2026-05-29).
- PTL U404 SDA Monthly + Burn-In → **45 rows**.
- ARL Refresh S → 4 yield rows.
- PTL H 12Xe → 12 yield rows.
- PTL U404 → **0 yield rows** (only Reconfig Yield exists; mobile-class U often lacks U/D, R/D, FINISH YIELD).
- ADL P 2+8+2 SDA Class=test → 4 rows (TEST_MPS only).
- Twin Lake N Refresh SDA Class=test → 141 rows.
- Raptor Lake S Catlow SDA Class=test → 88 rows.

## SDA Monthly WW→fiscal-month anchors (verified 2026-05-28, PTL 4Xe)

202622→**Jun 2026**, 202627→Jul 2026, 202631→Aug 2026, 202635→Sep 2026, 202640→Oct 2026, 202644→Nov 2026, 202648→**Dec 2026**, 202701→Jan 2027, 202722→Jun 2027, 202748→Dec 2027.

## WIF coverage summary

- Burn-In WIF: 3 specs (K LOT%, Stress Time, REBI) verified with multi-DLCP matching.
- Class WIF: 10 spec combinations verified across 5 products (single-spec, multi-spec, multi-range, site-level).

## Legacy per-product WIF scripts (superseded by tci_wif_template.ps1)

Kept for reference only — all new WIF work uses the template:
- `tci_adln_sda_class_wif_multi.ps1` — multi-spec (ETT + RCS, per-spec red font)
- `tci_ptlu404_sda_class_wif_multi.ps1` — PTL U404 multi-spec (fresh POST pattern)
- `tci_ptlu404_sda_class_wif_sqa.ps1` — PTL U404 ShortEQA ETT single-spec
- `tci_ptlu404_sda_class_wif_mps_mon_ss_pg8.ps1` — site-level MPS Monitor WIF (Site filter)
- `tci_ptl12xe_sda_class_wif.ps1` — single-spec multi-range
- `tci_arls8161lga_class_wif.ps1` — single month
- `tci_wcl_class_wif.ps1` — WCL
