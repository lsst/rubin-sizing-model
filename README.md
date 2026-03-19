# Rubin Sizing Model

Compute, storage, and cost projections for Rubin Observatory USDF operations
across 10 years of operations (LOY1–LOY10, FY2026–FY2035).

## Quick Start

```bash
pip install -r requirements.txt
python generate_sizing_model.py          # → rubin_sizing_model_2026.xlsx + charts/
```

All tunable parameters live in **`sizing_params.yaml`**. Edit that file and
re-run the script to regenerate the spreadsheet with updated formulas and
values. The spreadsheet itself contains live Excel formulas, so values can also
be changed directly in Excel and downstream cells will update automatically.

---

## Repository Layout

| File / Directory | Purpose |
|------------------|---------|
| `generate_sizing_model.py` | Python script that reads YAML and writes `.xlsx` with live Excel formulas, plus PNG charts |
| `sizing_params.yaml` | All tunable numbers: data product sizes, DRP estimates, hardware fleet, costs, retention policies, international compute shares |
| `requirements.txt` | Python dependencies (`openpyxl`, `pyyaml`, `matplotlib`) |
| `rubin_sizing_model_2026.xlsx` | Generated spreadsheet (regenerate with the script) |
| `charts/` | Generated PNG forecast charts (storage by tier, detail, data products, DRP core-hours, cost) |
| `LIMITATIONS.md` | Known gaps and suggested improvements |
| `dp-drp-derivations/` | DRP resource usage estimates, RFC-1134 prompt products doc, K8s batch usage report |
| `sizing-model-spreadsheet/` | Rubin Key Numbers spreadsheet and machine/storage pricing reference |

## Spreadsheet Tabs

| Tab | Description |
|-----|-------------|
| **Reference Data** | Key numbers, fleet inventory, data product baselines |
| **Ops Storage** | Per-LOY storage by tier (Flash, HDD, Object Store, Qserv, Tape) with sliding-window retention policies |
| **Ops Compute** | DRP, AP, DAC, Staff, Qserv, GPU, K8s core and node requirements per LOY |
| **Model** | Machine catalog, price factors, storage cost curves |
| **Ops Costs** | Annual CapEx for compute and storage, broken out by line item |
| **Purchase Plan** | Actionable per-year purchase requirements (batch nodes, K8s, storage by tier), with LOY1 (FY2026) detail for the current purchase cycle |
| **International Compute** | DRP compute split: USDF (35%), France / CC-IN2P3 (40%), UK / UKDF (25%) |
| **Yearly Readiness** | Per-LOY readiness (LOY1–LOY10) with milestone markers (DP2, DR1–DR9) and on-track / gap status |
| **Charts** | Forecast charts for storage and compute growth |

---

## References and How They Were Applied

### Documents Included in the Repository

| Document | Applied To |
|----------|------------|
| **DRP Resource Usage Estimates** (`dp-drp-derivations/DRP+Resource+Usage+Estimates.doc`) | DR1 core-hours, processing window, DRP output image and parquet sizes per visit, compression ratios. Used to calibrate `core_hours_per_input_tb`, `output_image_tb_per_visit`, and `parquet_tb_per_visit`. |
| **RFC-1134** (`dp-drp-derivations/RFC-1134.doc`) | Prompt products: alert stream sizing, Tier-1 (6-month) and Tier-2 (30-day) retention, long-term data preservation products stored at USDF. |
| **Rubin Key Numbers** (`sizing-model-spreadsheet/Rubin Key Numbers.xlsx`) | Visits per night (700), image size (8.19 GB logical), raw compression ratio (0.464), calibration images per day (450), images per visit (1). |
| **Rubin k8s_batch usage** (`dp-drp-derivations/Rubin k8s_batch usage .pdf`) | S3DF-wide utilization context for K8s and batch partitions. |
| **New sizing machines/storage** (`sizing-model-spreadsheet/new sizing machines_storage .xlsx`) | Torino node specs and flash/HDD tier pricing cross-reference. |

### External References (Not in Repository)

These documents were used during model construction but are not tracked in this
repository. Their key numbers have been captured in `sizing_params.yaml`.

| Document | Applied To |
|----------|------------|
| **DMTN-135** — DM Sizing Model & Cost Plan | Storage retention policies (sliding windows for Qserv 3 yr, Output Images 2 yr, Coadds 3 yr, Parquet 3 yr), tape model, compute scaling formulas, international compute shares, and the overall spreadsheet methodology. |
| **2026 HW Initial Capacity Capture** spreadsheet | Rubin USDF-specific fleet: Milano/Torino batch nodes, K8s nodes, interactive nodes, GPU nodes, installed storage. |
| **Original sizing model spreadsheet** (March 2026) | Blueprint for all formulas and inter-tab references. Storage retention logic, compute growth curves, tape accumulation, and chart structure were replicated from this spreadsheet. |
---

## YAML Configuration

`sizing_params.yaml` is organized into these sections:

- **general** — scale year, number of LOYs, first fiscal year
- **prompt** — alert rates, tier-1/tier-2 retention, DR1 target year
- **qserv** — storage per node, replication factor
- **observing** — nightly visits, image count, calibration images, engineering fraction
- **imaging** — image size, raw compression (science only), calibration compression, lossy compression, per-visit DRP output constants
- **storage_fractions** — APDB, scratch, misc overhead, sims output
- **catalogs** — object/source/forced-source row sizes and counts
- **users** — science user counts and storage per user
- **precursor** — HSC RC2/PDR2 reference data for coadd scaling
- **current_fleet** — Rubin USDF-specific node counts (batch, K8s, interactive, GPU, Qserv) and installed storage
- **costs** — unit costs for nodes and storage tiers, hosting/overhead
- **price_factors** — annual price change rates (currently 0 = flat pricing)
- **lifecycle** — compute (3 yr) and storage (5 yr) refresh periods
- **storage_retention** — sliding window lengths for Qserv, output images, coadds, Parquet
- **qserv_data_per_node** — per-LOY drive density ramp
- **row_sizes** — per-LOY catalog row size growth
- **dr_schedule** — DR1–DR9 target LOYs (annual; 11 total DRs per PSTN-019, DR10–DR11 beyond model window)
- **international_compute** — France (40%) and UK (25%) DRP shares
- **milestones_by_loy** — DP2, DR1–DR9 milestone labels per LOY (annual data releases)
- **dp2** — Data Preview 2 parameters (10% of DR1)
- **k8s** — K8s infrastructure growth rate
- **efd** — Engineering Facility Database storage estimate
- **current_year** — months elapsed/remaining for purchase cycle planning

---

## Methodology

1. **Data product sizing** starts from per-image byte counts (raw, processed,
   coadd, difference, Parquet catalogs) multiplied by nightly visit rates and
   accumulated over survey years.
2. **Compression**: Raw compression (0.464) is applied only to science and
   engineering images where measurements exist. Calibration images use a
   separate factor (default 1.0 = uncompressed) since no measured compression
   data is available. Lossy compression (0.27) is applied to output images on
   the object store.
3. **Storage retention** applies sliding windows (DMTN-135 policy): e.g.,
   Qserv czar tables kept for 3 years, output images 2 years, coadds 3 years
   after initial construction. Raw images are permanent on object store + tape.
4. **Compute** uses a linear scaling model: input TB per year of observations
   multiplied by cumulative years and a calibrated `core_hours_per_input_tb`
   factor (derived from actual DR1 processing). AP and DAC are additive
   fractions.
5. **International sharing** reduces USDF DRP compute by 65% (France 40% +
   UK 25%) in the International Compute tab only. All other tabs assume 100%
   USDF. Storage remains entirely at USDF.
6. **Purchase planning** compares per-LOY needs against cumulative available
   inventory (starting from the current 30 PB / fleet), producing incremental
   purchase amounts per year and per storage tier.
7. **Cost** uses flat unit pricing (no annual deflation). Compute refresh
   triggers after the 3-year lifecycle.
8. **Yearly Readiness** summarizes compute and storage gaps per LOY and flags
   each year as "On Track" or reports the specific gap. Available capacity
   reflects inventory *before* that year's purchase.

---

## Known Limitations

See [`LIMITATIONS.md`](LIMITATIONS.md) for a detailed list of gaps, their
impact, and suggested fixes. Key items include:

- Storage refresh (5-year disk lifecycle) not yet subtracted from available
- GPU growth not projected beyond the 2 existing H200 nodes
- No operating costs (power, cooling, FTEs)
- All parameters are point estimates (no sensitivity analysis)

---

## License

Internal Rubin Observatory planning document. Not for public distribution.
