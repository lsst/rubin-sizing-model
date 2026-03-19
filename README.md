# Rubin Sizing Model

Compute, storage, and cost projections for Rubin Observatory USDF operations
across 10 years of operations (LOY1–LOY10, FY2026–FY2035).

## Quick Start

```bash
pip install -r requirements.txt
python generate_sizing_model.py          # → rubin_sizing_model_2026.xlsx
```

All tunable parameters live in **`sizing_params.yaml`**. Edit that file and
re-run the script to regenerate the spreadsheet with updated formulas and
values. The spreadsheet itself contains live Excel formulas, so values can also
be changed directly in Excel and downstream cells will update automatically.

---

## Repository Layout

| File | Purpose |
|------|---------|
| `generate_sizing_model.py` | Python script that reads YAML and writes `.xlsx` with live Excel formulas |
| `sizing_params.yaml` | All tunable numbers: data product sizes, DRP estimates, hardware fleet, costs, retention policies, international compute shares |
| `requirements.txt` | Python dependencies (`openpyxl`, `pyyaml`) |
| `rubin_sizing_model_2026.xlsx` | Generated spreadsheet (regenerate with the script) |
| `LIMITATIONS.md` | Known gaps and suggested improvements |

## Spreadsheet Tabs

| Tab | Description |
|-----|-------------|
| **Reference Data** | Key numbers, fleet inventory, data product baselines |
| **Ops Storage** | Per-LOY storage by tier (Flash, HDD, Object Store, Qserv, Tape) with sliding-window retention policies |
| **Ops Compute** | DRP, AP, DAC, Staff, Qserv, GPU, K8s core and node requirements per LOY |
| **Model** | Machine catalog, price deflation factors, storage cost curves |
| **Ops Costs** | Annual CapEx for compute and storage, broken out by line item |
| **Purchase Plan** | Actionable per-year purchase requirements (batch nodes, K8s, storage by tier), with LOY1 (FY2026) detail for the current purchase cycle |
| **International Compute** | DRP compute split: USDF (35%), France / CC-IN2P3 (40%), UK / UKDF (25%) |
| **Yearly Readiness** | Per-LOY readiness (LOY1–LOY10) with milestone markers (DP2, DR1–DR4) and on-track / gap status |
| **Charts** | Forecast charts for storage and compute growth |

---

## References and How They Were Applied

### Primary Sizing Documents

| Document | Applied To |
|----------|------------|
| **DMTN-135** — DM Sizing Model & Cost Plan | Storage retention policies (sliding windows for Qserv 3 yr, Output Images 2 yr, Coadds 3 yr, Parquet 3 yr), tape model, compute scaling formulas, international compute shares, and the overall spreadsheet methodology. |
| **DRP Resource Usage Estimates** (`DRP+Resource+Usage+Estimates.doc`) | DR1 core-hours (24.6 M), DRP node-days (2,872), annual DRP growth (1.2×), processing window (200 days), and batch compute safety margin (1.2). |
| **RFC-1134** (`RFC-1134.doc`) | Prompt products: alert stream sizing (10k alerts/visit × 80 KB), Tier-1 (6-month) and Tier-2 (30-day) retention, long-term data preservation products to be stored at USDF. |
| **Rubin Key Numbers** (`Rubin Key Numbers.xlsx`) | Nightly visit count (1,000), image dimensions (4k × 4k, 32 MB raw), total survey nights (3,000), sky coverage (18,000 deg²). |

### Hardware & Fleet

| Document | Applied To |
|----------|------------|
| **2026 HW Initial Capacity Capture** (`2026_HW_Initial_Capacity_Capture.xlsx`) | Rubin USDF-specific fleet: 50 Milano batch nodes, 10 Torino batch nodes, 75 K8s nodes, 11 interactive nodes, 2 H200 GPU nodes, 30 PB installed online storage. |
| **Rubin k8s_batch usage** (`Rubin k8s_batch usage .pdf`) | S3DF-wide utilization context. Node counts are S3DF-wide; Rubin-specific counts were derived from the HW Capacity Capture. |
| **usage_per_month.png** | Core-hours by partition used to infer the number of Milano nodes available to Rubin (~50 nodes based on utilization patterns). |

### Cost & Procurement

| Document | Applied To |
|----------|------------|
| **Summarized conversation with KT** (`summarized-conversation-with-KT.txt`) | Node pricing (Milano $11,800, Torino $16,000), storage cost per TB (HDD $50/TB, NVMe $200/TB, tape $10/TB), 3-year compute lifecycle, 5-year storage lifecycle, annual cost deflation model (power-law factors). |
| **New sizing machines/storage** (`new sizing machines_storage .xlsx`) | Cross-reference for Torino node specs and flash/HDD tier pricing. |

### Original Spreadsheet

| Source | Applied To |
|--------|------------|
| **sizing-model-spreadsheet-03-19-2026.xlsx** | Blueprint for all formulas and inter-tab references. Storage sliding-window retention logic, compute growth curves, tape accumulation, cost deflation, and chart structure were replicated from this spreadsheet. Formulas that reference across tabs (Ops Storage ↔ Ops Compute ↔ Ops Costs) preserve the original dependency chain. |

---

## YAML Configuration

`sizing_params.yaml` is organized into these sections:

- **telescope** — nightly visits, image geometry, survey duration
- **drp** — DR1 baselines (node-days, core-hours), processing window, safety margin, annual growth
- **data_products** — per-image sizes, compression factors, Parquet catalog sizes, alert rates
- **storage_retention** — sliding window lengths for Qserv, output images, coadds, Parquet
- **current_fleet** — Rubin USDF-specific node counts (batch, K8s, interactive, GPU) and installed storage
- **costs** — unit costs for nodes and storage, lifecycle durations, deflation model
- **lifecycle** — compute (3 yr) and storage (5 yr) refresh periods
- **dr_schedule** — DR1/DR2/DR3 target LOYs
- **international_compute** — France (40%) and UK (25%) DRP shares
- **milestones_by_loy** — DP2, DR1–DR4 milestone labels per LOY
- **dp2** — Data Preview 2 parameters (10% of DR1)
- **k8s** — K8s infrastructure growth rate
- **efd** — Engineering Facility Database storage estimate
- **current_year** — months elapsed/remaining for purchase cycle planning

---

## Methodology

1. **Data product sizing** starts from per-image byte counts (raw, processed,
   coadd, difference, Parquet catalogs) multiplied by nightly visit rates and
   accumulated over survey years.
2. **Storage retention** applies sliding windows (DMTN-135 policy): e.g.,
   Qserv czar tables kept for 3 years, output images 2 years, coadds 3 years
   after initial construction. Raw images are permanent on object store + tape.
3. **Compute** converts DRP node-days → core-hours, then divides by the
   processing window to get peak concurrent cores. Annual growth (1.2×) is
   applied. AP and DAC are additive fractions.
4. **International sharing** reduces USDF DRP compute by 65% (France 40% +
   UK 25%). Storage remains entirely at USDF.
5. **Purchase planning** compares per-LOY needs against cumulative available
   inventory (starting from the current 30 PB / 7,600 cores), producing
   incremental purchase amounts per year and per storage tier.
6. **Cost** applies power-law deflation to unit prices, multiplied by
   incremental purchases. Compute refresh triggers after the 3-year lifecycle.
7. **Yearly Readiness** summarizes compute and storage gaps per LOY and flags
   each year as "On Track" or reports the specific gap.

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
