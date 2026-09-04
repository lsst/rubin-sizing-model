# Data Facility Sizing Model

Compute, storage, tape and cost projections for a Rubin data facility across
the ten loop years of operations.

One script reads one parameter file and writes one Excel workbook of live
formulas. The script contains no facility data at all — every number, label
and note lives in `sizing_params.yaml`. Change the YAML, re-run, and the
whole workbook follows.

## Prerequisites

[uv](https://docs.astral.sh/uv/) is the only requirement; it installs Python
and the two dependencies for you.

```bash
# macOS / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Homebrew
brew install uv

# Windows (PowerShell)
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Anything from Python 3.10 works. If you would rather not use uv:
`pip install openpyxl pyyaml` and run `python generate_model.py`.

## Quick start

```bash
uv run generate_model.py
```

Writes the workbook named in `workbook.output` in the YAML. Open it once in
Excel or LibreOffice to populate the formula results — `openpyxl` writes
formulas without cached values, so a freshly generated file looks empty to
anything that reads cached values only.

```bash
uv run generate_model.py --params other.yaml --output other.xlsx
```

## Repository layout

| Path | Purpose |
|---|---|
| `generate_model.py` | The script. Formulas and layout mechanics only. |
| `sizing_params.yaml` | Every input number, label, note and source citation. |
| `rubin_usdf_model_2026_condensed_with_pricing.xlsx` | Generated output, committed so it can be read without running anything. |
| `pyproject.toml` | Dependencies (`openpyxl`, `pyyaml`). |
| `dp-drp-derivations/`, `sizing-model-spreadsheet/` | Source documents the parameters were derived from. |

## What the workbook contains

Two tabs hold raw numbers. Every other tab is calculated from them, so no
figure is stated in two places.

| Tab | Role | Contents |
|---|---|---|
| **Key Numbers** | reference | Sizing and data-product inputs: observing, DRP compute, storage basis, tape, Alert Production, user/DAC, database, fleet, facility shares, and the dataset-type breakdown the campaign total is summed from. |
| **Cost Inputs** | reference | Costs only: unit prices, hardware lifetimes, historic purchases by fiscal year, the current year's actual requests. |
| **All at USDF** | calculated | Single-site scenario — DRP compute in core-hours, node-days and peak cores; the on-floor disk build-up; cumulative tape; Alert Production and services; file counts. |
| **Facility Split** | calculated | The same workload split across partner facilities, with per-year shares and capacity checks. |
| **Qserv** | calculated | Cluster specification, catalogue-driven database size, and node projection. |
| **DF Template** | calculated | Blank-fill calculator so another facility can size its own share. |
| **USDF Pricing Forecast** | calculated | Annual purchases and spend, with a ten-year total. |

Reference tabs may not read another tab. The projection tabs read only the
reference tabs (and, for the split and the pricing forecast, the tabs that
already computed demand and node counts — never re-deriving them, so the tabs
cannot disagree).

### Cell colours

| Colour | Meaning |
|---|---|
| Blue | Raw input — safe to edit |
| Black | Formula |
| Green | Link to another tab |
| Yellow | Key assumption that materially moves the answer |

## How the model works

**Compute.** DRP work is expressed in core-hours, calibrated against measured
campaign usage. It is shown three ways — core-hours, node-days, and peak
concurrent cores over the processing window — because campaigns are reported
in node-days but hardware is bought in cores. Where a fleet mixes CPU
generations, capacity is expressed in equivalents of whichever generation the
measurements were taken on.

**Storage.** *Bytes written* and *on-floor need* are different quantities and
the workbook keeps them apart. A data-release campaign writes far more than it
retains, because most intermediates are created and deleted during the run.
On-floor need is the live fraction of those intermediates, plus retained
releases over a sliding window, plus cumulative raw data and prompt products.
Purchases follow on-floor need, not bytes written.

**Tape** is cumulative — raw data and every release are preserved — so the
footprint only grows. Purchases start once it exceeds the capacity already
available.

**Cost.** Refresh is cohort-based: hardware bought in FY*x* is replaced in
FY*x* + lifetime, driven by the historic purchase grid. Changing a lifetime
moves every replacement automatically. The first column uses actual capture
requests rather than modelled figures.

Each formula block is documented in a notes section at the foot of its tab.

## Assumptions that move the answer

These are marked yellow in the workbook. The first is by far the most
significant in any facility's instance of this model.

| Assumption | Why it matters |
|---|---|
| **Live-intermediate fraction** | Sets how much transient campaign data is resident at once, and dominates the disk budget. Replace the placeholder with campaign telemetry as soon as you have it. |
| Tape capacity assumed available | Shared-facility accounting rarely says what a single tenant may actually draw on; owning media does not reserve library slots or drive bandwidth. Set to 0 for the no-headroom case. |
| Idle / inter-stage multiplier | Task-level CPU is usually measured; wall-clock inflation between stages usually is not. |
| Facility shares | Provisional until each partner supplies its own plan. |
| Hosting and overhead rates | Often carried forward from an older agreement rather than re-quoted. |
| Per-core speedup between CPU generations | Measured on one era of tasks; re-measure per release. |

Network and WAN are deliberately not modelled here: measured throughput
figures matter too much to guess at.

## Adapting this for another facility

Edit `sizing_params.yaml` only. The parameter blocks at the top hold the
numbers; the `workbook:` block at the bottom declares the tabs, sections, rows
and notes. A row draws its value from a `param:` pointer into those blocks, a
literal `value:`, or `formula: true` when the script computes it.

Row numbers appear nowhere. They are derived from the order of the YAML — a
section banner, then its rows, then one blank line — and formulas address
rows by key, so inserting or deleting a row re-points every dependent formula
automatically.

Two conventions the script enforces, both of which raise rather than silently
producing a broken workbook: a note may not begin with `=`, since Excel would
parse it as a formula; and a `param:` path that does not resolve is an error,
not a blank cell.
