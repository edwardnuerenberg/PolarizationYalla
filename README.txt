PEMFC YALLA 6 - README
=======================
Zentrum fuer BrennstoffzellenTechnik gGmbH

Programm for automatic processing, visualization as U/I curves and bar Diagramms
of extracted Performance parameters of PEMFC OCV and polarization measurements 
from MS2-Engineering testbenches controled by LabVIEW. 

0. FILES IN THIS FOLDER
------------------------

The program looks for its companion files NEXT TO ITSELF (next to YALLA6.py,
or next to the .exe once it is packaged). A copy placed next to the program
always wins; the packaged .exe also carries its own fallback copies inside,
so a lone .exe still runs - but then the eLabFTW configuration can no longer
be adapted without a rebuild.

Required:
   YALLA6.py           The program itself (or "YALLA6.exe" once packaged).
   README.txt          This file. Opened by the "Open README" button - the
                       button shows an error if it is missing.

Optional but recommended:
   elabftw_bol_procedures.json
                       Defines the metadata / BOL / AST / per-Pol condition
                       fields offered in the eLabFTW export dialog. This file
                       is meant to be edited by you: new fields or procedures
                       can be added here without changing any program code.
                       If it is missing, the eLabFTW export still works but
                       offers only the raw Pol results (see section 5).
   elabftw_export_format.json
                       Controls the SHAPE of the exported JSON itself - group
                       ids/names and the field-name text - as opposed to
                       elabftw_bol_procedures.json above, which controls what
                       the export DIALOG offers. Edit only the values you
                       want to change; anything left out, or the whole file
                       if missing, falls back to the shipped default shape
                       (see section 5).

Created automatically (do not need to be shipped):
   elabftw_export_values.json
                       Remembers the values last typed into the eLabFTW
                       export dialog, so they do not have to be retyped.
                       Written on every export; safe to delete (it will
                       simply start empty again).

Needed only when running YALLA6.py directly (the .exe carries its own copy):
   UnRAR.exe           Used to unpack .rar data archives. Freeware from
                       RARLAB (www.rarlab.com), redistributed unmodified.
                       Without it, .rar archives cannot be opened - .zip
                       archives and plain folders still work.

Development notes (not used by the running program):
   CONCEPT_mean_curves.md
                       Design concept for a planned feature: averaging
                       several curves into one mean curve with a variability
                       band, plus mean/spread of the extracted metrics.
                       Analysed and specified, but NOT implemented.

Your measurement data does NOT belong in this folder - it is picked
separately via "Select Folder" / "Select Archive" and can live anywhere.


1. TYPICAL PROCEDURE
---------------------

1) "Select Folder"  or  "Select Archive (.rar / .zip)"
   Two ways to load the same thing - pick whichever matches how the data
   arrived. An archive is unpacked to a temporary folder, read, and the
   temporary copy is deleted again: your archive is never modified, and
   nothing is left behind on disk.

   Either way the app scans the data (including subfolders) for files named
   "..._01_YYYYMMDD.txt" - these are the Pol curve files that get read. Each
   dataset's "Key" (shown in the Dataset dropdown, and used everywhere else
   in the app) is the name of the folder that directly contains that file,
   e.g. "M4895_TS3_KIMODPEM". For measurement files lying loose at the top
   of an archive, the archive's own name is used as the Key.
   Per folder only ONE run number is read, so that two separate measurements
   can never end up concatenated into one curve. Preferred is the lowest run
   number that actually contains rows marked Pol 1-6 - not merely a run that
   is readable, since an aborted attempt often leaves a few unmarked rows
   that read_csv accepts happily. "_01_" wins whenever it qualifies; if it
   doesn't, the next number that does is used instead ("_02_", then "_03_",
   ...). Only when NO run in the folder has Pol markers does the lowest
   readable one get used anyway, so the folder stays visible instead of
   disappearing outright.

   What happened is never left to a dialog you might click away: it is
   recorded per dataset in the "Pol Data Structure" column, which appears in
   the results table, in both the Wide and Long sheets of the export, and
   (appended automatically) in Update Master Excel.

     OK                                  only "_01_", used as intended
     OK — 01_ used (02_ ignored)         "_01_" has Pol data; a later run
                                          exists too but was not touched -
                                          the dataset itself is fine
     Check data! (used 02_, not 01_)     "_01_" had no usable Pol rows, so
                                          a later run was used INSTEAD
     Check data! (01_ used, ...)         "_01_" was used, but extraction
                                          still came up empty - the reason
                                          is named: no SetMarker column, no
                                          Pol 1-6 markers, no numeric U/I,
                                          or no i @ 0.6 V in any Pol
     Missing                             no run in this folder contained
                                          any data at all

   The first word is always OK, Check data!, or Missing, so a text filter
   in Excel catches all three regardless of the detail in parentheses.

   A folder can still hold several "_01_" files (e.g. Pol and EIS); those are
   read together as before. Files not matching "..._NN_YYYYMMDD.txt" at all
   are ignored and the folder stays invisible.
   Password-protected and multi-part .rar archives are not supported.

2) Set "Active area [cm2]" and "Points per average" (see section 2).

3) "Extract i@0.6 / i@0.65 + OCV + Pmax"
   Computes, for every dataset and every Pol (1-6):
     - current density at 0.6 V
     - current density at 0.65 V
     - OCV (open-circuit voltage)
     - Pmax (maximum power density)
   Results open in a popup table and are kept in memory for the rest of the
   session. Every export/update/plot step below (Export Results, Update
   Master Excel, Export to eLabFTW, Bar Diagram) reads from this in-memory
   result set, NOT from disk - re-run this step after changing the loaded
   data or the Active area / Points per average settings.

4) From here, use any combination of:
     - Plot Polarization / Plot Power for a single dataset
     - Overlay (Key + Pol) for multiple datasets on one graph
     - Bar Diagram to compare a metric across datasets
     - Export Results (CSV/XLSX)
     - Update Master Excel
     - Export to eLabFTW (JSON)


2. GRAPHICAL OPTIONS
---------------------

Dataset dropdown
   Choose which loaded key "Plot Polarization" / "Plot Power (P-J)" acts on.

Active area [cm2] / Points per average
   Applied globally: Active area converts raw current [A] to current
   density [A/cm2]; Points per average is the block size used to average
   consecutive data points before plotting or exporting curve points
   (default 10). Both affect every plot and export below, including the
   Curves_Polarization export sheet.

Polarization curve axes / Power axes
   Optional manual axis limits (J/U or J/P). Leave a field blank to let
   matplotlib auto-scale that side.

Show Pmax table (Power plot)
   Adds a small Pmax summary table to the corner of the Power plot.

Overlay (Key + Pol)...
   Build a custom list of (dataset, Pol) pairs and plot or save them
   together on one Polarization or Power graph.
     - Vary symbols per dataset: gives each dataset its own marker shape
       instead of the marker following the Pol number.
     - Vary colours per dataset: same idea, but for color.
   Normally color AND marker both encode the Pol number, consistently
   across every plot type in the app (e.g. Pol 1 is always a blue circle).
   Turning either toggle on trades that consistency for being able to tell
   datasets apart when the same Pol appears in several overlaid datasets.
   If BOTH toggles are on at once, Pol is only distinguishable via the
   legend text (e.g. "M4895_TS3_KIMODPEM - Pol 1"), not visually.

Bar Diagram (Extracted Values)...
   Compare one metric (i@0.6V, i@0.65V, OCV, or Pmax) across chosen
   datasets and Pols as a grouped bar chart. Both dataset names and the
   metric label can be renamed just for the chart.

Pol Styles...
   Customize the color / marker / line style used for each Pol (1-6),
   applied across all plot types. "Reset to defaults" restores the
   built-in palette.

Any plot window
   Press "E" to open a label editor and rename the title/axis labels
   before saving the figure.


3. EXPORT RESULTS (CSV/XLSX)
------------------------------

"Export Results (CSV/XLSX)" writes the extracted metrics (from step 1.3)
to disk. Choosing a .xlsx path produces one workbook with three sheets
(frozen header row, auto-sized columns); choosing .csv produces up to
three separate files instead:

   Wide      (*_wide.csv)               one row per dataset, all Pol
                                         columns side by side
   Long      (*_long.csv)                one row per dataset + Pol
   Curves_Polarization (*_curves_polarization.csv)
                                         the averaged (block-mean) J/U
                                         points actually used for plotting,
                                         per dataset and Pol - useful if you
                                         want to replot the curves elsewhere

All numeric values are rounded to 4 significant figures.


4. UPDATE MASTER EXCEL
------------------------

"Update Master Excel (.xlsx)" merges the extracted metrics into an
existing lab master spreadsheet, without touching the original file - the
result is always saved as a NEW file that you choose the path for.
One can also overwrite the master excel file. ATTENTION: All SHEETS except the
receiving master sheet will be deleted, same if you overwrite an existing file
the already existing values in the receiving sheet will be preserved.

How matching works:
   The app reads the FIRST sheet of the file you pick, and for each row
   compares the first 5 characters of its "MEA" column against the first 5
   characters of each dataset's Key (folder name). On a match, it fills in
   24 columns: Pol 1-6 i@0.6V, i@0.65V, OCV, and Pmax (columns are created
   automatically if the master file doesn't have them yet; only values
   that aren't "N/A" are written), plus "Pol Data Structure" - the raw-file
   verdict described in section 1.1.

   "Pol Data Structure" is written on every match, even when no metric
   could be filled in: a row whose numbers stayed empty is the one where
   knowing why is worth most. Rows whose MEA matched no dataset are left
   untouched, blank included - nothing was scanned, so nothing is claimed.

If you get "No rows were updated", check:
   - The data table is on the FIRST sheet of the workbook (a later sheet
     is never read).
   - The column header is spelled exactly "MEA" - this is a case-sensitive,
     exact match (not "Mea", not "MEA " with a trailing space).
   - The first 5 characters really match, e.g. "M5102" is NOT the same as
     "M5101" - a one-character typo is enough to silently skip the row.
   - "Extract i@0.6 / i@0.65 + OCV + Pmax" has been run for the dataset in
     question during THIS session - the lookup only ever uses what's
     currently in memory, not a previous export.

Matched values are rounded to 4 significant figures before saving.


5. EXPORT TO ELABFTW (JSON)
------------------------------

"Export to eLabFTW (JSON)..." writes one eLabFTW-compatible extra_fields
JSON file per dataset, ready to import into an eLabFTW experiment.

Scope
   Export just the currently selected dataset, or all extracted datasets
   at once (one JSON file each).

Optional sections
   These only appear if "elabftw_bol_procedures.json" exists next to the
   script/exe. Without it, export still works but only includes the raw
   Pol results (i@0.6V, i@0.65V, OCV, Pmax) - no metadata/BOL/AST/condition
   fields.
     - Metadata (Identification): filled in once per export run.
     - BOL / AST procedure: conditions applied to the whole cell before any
       Pol curve (BOL is typically required if used; AST is optional, for
       accelerated stress tests).
     - Per-Pol operating conditions (stoichiometry, temperature, pressure,
       RH anode/cathode): set individually for each Pol 1-6, and only
       included in the export for Pols that actually have extracted data.

Each Pol that has results becomes its own extra_fields_group in the JSON,
so eLabFTW displays a clean per-Pol block of conditions next to that Pol's
results.

Pol N is written out as "EC N"
   The PEMFC MEA Test Procedure Builder numbers its blocks
   EC(conditionId + catalogSize x phaseIndex) - the same arithmetic the test
   bench's SetMarker uses, and therefore the same number as this program's
   Pol slot. So Pol 4 is exported as EC4: the post-AST measurement of
   condition 1. That is what lets a procedure JSON and a results JSON of the
   same run be loaded into one eLabFTW entry and line up field for field,
   with the measured values landing in the blank slots the procedure left
   for them.

   Field names are a shared contract between this program, the Procedure
   Builder and the eLabFTW Dashboard. Nothing errors if they drift apart -
   a value just silently arrives as a new field next to the empty one that
   was waiting for it. Before changing any exported name, read
   "ec-field-naming.md" in the Procedure Builder repository.

Field values you type are remembered between export runs (saved to
"elabftw_export_values.json" next to the script/exe), so metadata doesn't
need to be retyped every time.

Adjusting the export format
   "elabftw_export_format.json" (next to the script/exe) controls the shape
   of the written JSON, independently of which fields are offered:
     - Group ids and names (Identification, BOL, AST, per-Pol results and
       per-Pol conditions).
     - The field-name text for the four result metrics (i@0.6V, i@0.65V,
       OCV, Pmax) - handy if your eLabFTW template expects different wording.
     - The output filename pattern.
   The file is self-documenting (see its "_readme" keys) and only needs to
   list the values you actually want to change - anything left out keeps
   the shipped default. If the file is missing, invalid, or a value is left
   out, that value's default is used and a note is printed to the console;
   the export itself never fails because of this file.
