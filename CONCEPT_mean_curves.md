# Concept: Mean Curves with Variability — Outlook for the Next AI

**Status:** analysed, agreed as a concept, **not implemented**. Deferred by the user on 2026-08-08.
**Target file:** `Edward/YALLA6.py` (single-file Tkinter + matplotlib app).

---

## 1. What the user asked for

> "I want to include a function to make mean curves from more than 2 curves. Even if the single
> points in x or y axis are not exactly the same it should still work to make mean values of
> U/i and P curves, and also the columns of i at 0.6 V and OCV and Pmax etc. Preferably also
> with variability / error bars, if there are enough values."

Two distinct deliverables:

1. **Mean curves** — average several Polarization (U vs. J) and Power (P vs. J) curves into one,
   with a variability band.
2. **Mean scalar metrics** — mean ± spread for the four extracted numbers
   (`i @ 0.6 V`, `i @ 0.65 V`, `OCV`, `Pmax`) across the selected datasets.

---

## 2. Codebase context you will need

Line numbers are from the state of the file when this concept was written — verify before relying
on them.

| What | Where | Notes |
|---|---|---|
| `data_dict` | `YALLA6.py:28` | `{key: raw DataFrame}` — one entry per measurement folder |
| `current_at_06V` | `YALLA6.py:29` | `{key: {"Pol 1 i @ 0.6 V [A/cm²]": val, ...}}` — the extracted scalars |
| `POL_STYLES` | `YALLA6.py:34` | color/marker/linestyle per Pol 1–6; user-editable via the Pol Styles dialog |
| `interpolate_current()` | `YALLA6.py:89` | existing sort-then-linear-interpolate helper; same *pattern* the new code needs |
| `extract_current()` | `YALLA6.py:293` | fills `current_at_06V`; **rebuilds it from scratch each run** |
| `plot_data()` | `YALLA6.py:442` | single-dataset Polarization plot; shows the block-averaging idiom |
| `plot_power_data()` | `YALLA6.py:502` | single-dataset Power plot |
| `overlay_iv_pairs()` | `YALLA6.py:724` | overlay of `(key, pol)` pairs — Polarization |
| `overlay_power_pairs()` | `YALLA6.py:826` | overlay of `(key, pol)` pairs — Power |
| `open_overlay_keypol_selector()` | `YALLA6.py:935` | the dialog that builds the `(key, pol)` pair list |
| `show_current_table()` | `YALLA6.py:387` | existing results-table popup — copy its style for the metrics table |
| `attach_label_editor()` | `YALLA6.py:713` | call this on any new figure so "press E to edit labels" keeps working |

### Key existing behaviours to preserve

- **Block averaging.** Every curve is first reduced by averaging `bs` consecutive raw points
  (`parse_block_size()`, "Points per average", default 10). The mean-curve feature must consume
  these *same* block-averaged curves, so a mean curve is numerically consistent with the individual
  curves drawn next to it.
- **Active area.** Raw `I Mittel [A]` is divided by `parse_active_area()` to get `J [A/cm²]`.
  The Power path prefers `JFilter [A/cm²]` and `U1 Mittel [V]` when those columns exist —
  see `overlay_power_pairs()` for the exact fallback chain. Reuse it, don't re-derive it.
- **Pol-based styling.** Color and marker encode the Pol number consistently across the whole app.
  Overlay has two opt-in toggles (`vary_marker_by_dataset_var`, `vary_color_by_dataset_var`) that
  reassign marker/color to the *dataset* instead. Mean curves should follow `POL_STYLES` — a mean
  curve represents one Pol, so per-dataset styling is meaningless there.

---

## 3. The core technical problem

Repeated measurements never land on the same J values (different sweep timing, noise, block
boundaries), so curves **cannot be averaged row-by-row**. Averaging requires putting every curve
on a shared abscissa first.

### Proposed approach

1. Build the block-averaged `(J, U)` — and `(J, P)` — curve per `(key, pol)`, exactly as the
   existing overlay functions already do.
2. Construct a **common J grid** spanning the *union* of all selected curves' J ranges
   (not the intersection). Rationale: intersection would truncate the mean curve to the shortest
   run; union keeps the full range and simply has fewer contributing curves near the edges.
   Suggested resolution: a few hundred points, or the median point count of the input curves.
3. Linearly interpolate each curve's `U(J)` / `P(J)` onto that grid. Outside a given curve's own
   J range, yield `NaN` — **do not extrapolate**, and make sure `np.interp` is not used naively,
   because it clamps to the endpoint value instead of returning `NaN`.
4. Per grid point, compute `mean`, `std`, and `n` across the curves that have a value there
   (`np.nanmean` / `np.nanstd` with an explicit `n` count).
5. Plot the mean as a line styled by `POL_STYLES[pol]`, with a shaded ±STD band via
   `ax.fill_between(...)`. **Mask the band where `n < 2`** — a band drawn from a single curve is
   not variability, it is a rendering artifact.

### Grouping rule

Group the selected `(key, pol)` pairs **by Pol code**. Selecting Pol 1 from three datasets gives
one averaged Pol-1 curve; adding Pol 2 pairs adds a second averaged curve to the same axes.

---

## 4. Mean scalar metrics

No interpolation needed — these are already single numbers per dataset+Pol in `current_at_06V`.
For each Pol, aggregate across the selected datasets: `mean`, `std`, `n`. Use
`to_float_or_nan()` (`YALLA6.py:114`) to handle the `"N/A"` sentinel that extraction writes for
missing values, and `round_sig_numeric()` (`YALLA6.py:123`) for display, matching the rest of the app.

Present it in a popup table modelled on `show_current_table()`.

> ### ⚠️ Pitfall worth stating in the UI
> **The mean of the individual Pmax values ≠ the peak of the mean power curve.**
> Because each run peaks at a slightly different J, averaging the curves flattens and lowers the
> peak. Both numbers are legitimate but answer different questions ("typical peak performance of a
> single cell" vs. "peak of the average behaviour"). If both are ever shown, label them distinctly.
> The same caution applies to OCV if curves start at different J.

---

## 5. UI proposal

**Extend the existing Overlay dialog** (`open_overlay_keypol_selector()`, `YALLA6.py:935`) rather
than building a new one — it already has exactly the right selection mechanism (pick key + Pol,
accumulate a pair list, and the pairs are already grouped naturally).

Add to the existing button row:

- `Plot Mean ± STD` — Polarization mean curve(s)
- `Plot Mean Power ± STD` — Power mean curve(s)
- `Save Mean PNG`
- `Show Mean Metrics Table`

Note the button row currently sits at `row=7` of the dialog grid, below the two "vary…" checkboxes
at rows 5 and 6. Adding widgets above it means bumping that row index again.

---

## 6. Decisions taken (defaults, if nobody says otherwise)

| Question | Default | Rationale |
|---|---|---|
| Spread measure | **Standard deviation** | The user asked for "variability", which is STD. SEM (`std/√n`) describes the precision of the mean instead — offer it as a toggle only if asked. |
| Visual style | **Shaded band** (`fill_between`) | Cleaner on a dense interpolated grid; per-point error bars would be visual noise. Discrete error bars remain reasonable for the scalar-metrics bar chart. |
| Minimum n for a band | **≥ 2** to draw anything | Below 2 there is no spread to show. ≥ 3 before it is statistically meaningful — consider annotating `n` on the plot so the reader can judge. |
| J grid extent | **Union** of ranges, no extrapolation | Keeps the full curve; edges naturally show fewer contributing curves. |

---

## 7. Effort estimate

Roughly **2–2.5 hours** of focused work:

| Task | Estimate |
|---|---|
| Interpolation + aggregation helper (grid, resample, mean/std/n) | ~35 min |
| Mean-curve plotting with band, for both Polarization and Power | ~30 min |
| Mean-metrics table popup | ~25 min |
| UI wiring in the Overlay dialog | ~15 min |
| Testing with real repeated datasets | ~15–20 min |

This is a genuine feature, not a tweak — it adds new numerical logic, a new plotting path, a new
table, and new UI. Do not quote it as a quick fix.

---

## 8. Extension points worth considering later

- **Export.** The mean curve and mean metrics are currently plot-only in this concept. They are
  natural additions to `export_results()` (`YALLA6.py:1037`) as extra sheets, e.g.
  `Curves_Mean` (J, U_mean, U_std, n) and `Metrics_Mean` (Pol, metric, mean, std, n).
- **Bar diagram.** `plot_bar_diagram_multi_pol()` (`YALLA6.py:1858`) could gain `yerr=` error bars
  once per-Pol std values exist — small addition on top of this work.
- **eLabFTW export.** Mean ± std could be exported as its own extra_fields group, but only if the
  user asks; the current export is deliberately per-dataset.

---

## 9. Notes for whoever implements this

- The app is a **single 2100+ line file** with module-level Tk widget creation at the bottom.
  Functions defined above reference widgets created below via globals; that is existing style,
  not a bug. Guard optional globals with the `'name' in globals()` idiom already used for the
  overlay toggles.
- Run `python -m py_compile Edward/YALLA6.py` after edits — it is the fastest sanity check, since
  the GUI cannot be exercised headlessly.
- User-facing wording: this project uses **"Polarization"**, never "I–V" (renamed deliberately —
  "I–V" implies current in amps, but the axis is current *density* J in A/cm²). Internal Python
  identifiers such as `overlay_iv_pairs` and `get_iv_axes_limits` still contain `iv` and were
  intentionally left alone. Keep that split: new *labels* say Polarization, new *code names* can
  follow the surrounding convention.
- Companion files (`README.txt`, `elabftw_bol_procedures.json`) are located via `_script_dir()`
  (`YALLA6.py:1358`). If new config files are added, use the same helper so a future PyInstaller
  build keeps working.
- Packaging to `.exe` is a planned next step. Data files next to the script are **not** bundled
  automatically by PyInstaller — they must be declared in the `.spec`. Several older `.spec`
  files already exist in the folder to use as a starting point.
