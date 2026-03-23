# Rubin Sizing Model — Limitations & Gaps

This document identifies known limitations in the current model logic and areas
that would benefit from refinement to increase the model's robustness.

---

## 0. Hardware refresh & DRP document

| # | Gap | Impact | Suggested Fix |
|---|-----|--------|---------------|
| 0.1 | **Hardware Refresh tab is cohort-based** — Replacement counts assume strict `MOD(FY−delivery, lifetime)=0`; real procurement may slip. | Planning noise. | Adjust cohorts/dates in `sizing_params.yaml` as POs close. |
| 0.2 | **JBOD cohort quantities** — JTM “Total JBODs” (62) may not equal the sum of per-PO lines in YAML; cohort list is for EOL timing, not audit. | Minor inventory mismatch. | Reconcile with asset DB. |
| 0.3 | **`core_hours_per_input_tb` vs DRP doc** — The DRP Resource Usage Estimates doc now includes DM-53697 / DM-52836 CPU sections; the model still uses a single calibrated `core_hours_per_input_tb`. | May diverge from latest task-level totals until recalibrated. | Re-derive from updated doc or Jim’s node-day totals. |

## 1. Compute

| # | Gap | Impact | Suggested Fix |
|---|-----|--------|---------------|
| 1.1 | **International compute shares are static** — France (40%) and UK (25%) DRP shares are fixed constants. Actual shares may vary by DR campaign. | USDF compute purchases could be over- or under-estimated. | Make shares per-LOY configurable in `sizing_params.yaml`. |
| 1.2 | **GPU workload projection is rudimentary** — Only 2 H200 nodes are tracked; no growth model for ML/DIA/deep-learning pipelines. | GPU cost under-represented in later LOYs. | Add a GPU growth curve (e.g., per-DR GPU core-hour estimate). |
| 1.3 | **No per-pipeline breakdown** — DRP is treated as a single block. Individual pipelines (ISR, Coadd, multifit, deblending) are not sized separately. | Cannot identify which sub-pipeline drives growth. | Add a pipeline-level sheet or YAML section with per-pipeline core-hour factors. |
| 1.4 | **DP2/DP3 compute not modelled explicitly** — DP2 is approximated as 10% of DR1. Actual Data Preview runs may differ in scope and timing. | LOY1 compute need may be slightly misrepresented. | Add explicit DP schedule entries with per-DP core-hour budgets. |
| 1.5 | **K8s growth is a flat percentage** — 5% annual growth is an assumption without workload backing. | K8s node purchases may be inaccurate. | Tie K8s growth to the number of services / Prompt Processing cadence. |
| 1.6 | **No interactive / RSP compute budget** — DAC/LSP cores are a fixed fraction of DRP. Actual notebook/portal usage varies. | Interactive capacity may be under-provisioned. | Track RSP usage metrics and feed back into the model. |
| 1.7 | **Compute depreciation is a cliff** — nodes are assumed 100% useful until year N, then 0%. In practice, older nodes lose efficiency. | Overestimates available capacity just before refresh. | Apply an age-based efficiency curve. |

## 2. Storage

| # | Gap | Impact | Suggested Fix |
|---|-----|--------|---------------|
| 2.1 | **No explicit storage depreciation / refresh** — disks have a 5-year lifecycle but the model does not subtract end-of-life capacity. | Available storage may be overstated after LOY5. | Add a storage refresh row analogous to compute refresh. |
| 2.2 | **Tape capacity is not constrained** — tape is assumed to be infinitely purchasable. Slot/library limits are not modelled. | Cost is correct, but physical planning is missing. | Add tape library slot accounting. |
| 2.3 | **Object Store ≠ Ceph HDD** — the model lumps Object Store into the "online" bucket, but Object Store may sit on different Ceph pools with different cost/performance. | Cost granularity is lost. | Split Object Store into its own cost line in Ops Costs. |
| 2.4 | **EFD and node-local storage are not tracked as purchasable items** — EFD telemetry (~1 TB/yr) lives on K8s node-local NVMe. When K8s nodes are added, local storage comes for free, but this is implicit. | Minor — EFD is small. | Note in README; no code change needed unless EFD grows. |
| 2.5 | **No Weka vs. Ceph flash split** — Flash storage is modelled as a single tier. Weka (parallel NVMe) and Ceph NVMe have different costs. | Undercounts cost if Weka is used heavily. | Add separate Weka and Ceph NVMe rows. |
| 2.6 | **Compression is simplified** — Science/engineering raws use Key Numbers ratio; calibrations use the same factor as raws (0.464) until better measurements exist. Lossy output uses a single DRP-derived factor. | Uncertainty on calibration and coadd-related storage. | Update `imaging.calibration_compression` and per-product factors from ops data. |
| 2.7 | **Sliding window boundaries are exact years** — e.g., a "3-year window" drops data at the LOY boundary. In reality, retention may be calendar-date based. | Minor mismatch between model and policy enforcement. | Acceptable approximation for planning. |

## 3. Cost

| # | Gap | Impact | Suggested Fix |
|---|-----|--------|---------------|
| 3.1 | **Unit costs are point-in-time** — the model uses a power-law deflation curve. Actual vendor pricing fluctuates and depends on procurement volume. | Costs diverge further out. | Refresh unit costs annually from procurement quotes. |
| 3.2 | **No network / bandwidth costs** — data transfer between USDF, France, and UK is not costed. | Relevant if large datasets ship internationally. | Add WAN transfer cost estimate. |
| 3.4 | **Torino pricing is provisional** — the Torino node cost ($16,000) may change at procurement. | Budget uncertainty for future LOYs. | Update `sizing_params.yaml` when RFQ is received. |

## 4. Data Products

| # | Gap | Impact | Suggested Fix |
|---|-----|--------|---------------|
| 4.1 | **Alert stream volume is approximated** — 10,000 alerts/visit × 80 KB, but real alert counts vary with sky depth and galactic latitude. | Minor overall, but affects Prompt Processing storage. | Use commissioning alert statistics. |
| 4.2 | **APDB growth model is linear** — the alert processing database grows at a fixed rate per year. Actual growth depends on survey strategy and transient density. | Could be ±20% by LOY10. | Use the DPDD's APDB sizing formula. |
| 4.3 | **No difference between full-depth and shallow data** — DR output image sizes are based on full-depth. Early DRs with fewer visits produce smaller images. | Overestimates early-DR storage. | Scale output image size by cumulative visit fraction. |
| 4.4 | **Prompt Tier-1 and Tier-2 partition is fixed** — 6-month Tier-1 and 30-day Tier-2. Policy may evolve. | Minor. | Make retention periods configurable (they already are in YAML). |

## 5. Model Structure

| # | Gap | Impact | Suggested Fix |
|---|-----|--------|---------------|
| 5.1 | **No Monte Carlo / sensitivity analysis** — all parameters are single point estimates. | Cannot quantify confidence intervals on projections. | Add a sensitivity analysis script that varies key params ±20%. |
| 5.2 | **No integration with actual procurement data** — the model does not automatically ingest what was actually purchased. | Drift between plan and reality. | Add a "Procurement Actuals" tab that feeds into Available. |
| 5.3 | **Charts are generated as openpyxl objects** — they may render differently across Excel versions. | Visual artifacts in LibreOffice / Google Sheets. | Acceptable for now; PDF export recommended for sharing. |
| 5.4 | **No versioning / changelog in the spreadsheet itself** — only the YAML + script are version-controlled. | Hard to tell which parameter set produced a given .xlsx. | Embed a metadata sheet with generation timestamp and git hash. |

---

## Prioritized Next Steps

1. **Storage refresh accounting** (2.1) — highest risk of silently inflating available storage.
2. **GPU growth model** (1.2) — increasingly important as ML pipelines mature.
3. **Per-pipeline compute breakdown** (1.3) — essential for capacity planning conversations with pipeline teams.
4. **Sensitivity analysis** (5.1) — critical for confidence when presenting to stakeholders.
