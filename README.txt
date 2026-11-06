MRD Traceability Proof of Concept - Python 3.15-compatible
==========================================================

No pip installation required. Uses only the Python standard library.
Does not use Snakemake, PyYAML, pandas, or openpyxl.
For demonstration ONLY; not validated for clinical or diagnostic use.

1. Extract the ZIP file into a folder.
2. Open PowerShell inside the MRD_traceability_POC_py315 folder.
3. Confirm Python version:

   py --version

4. Dry run (does not execute scripts or create an audit package):

   py run_mrd.py --workflow workflow_test.json --config config.csv --dry-run

5. Execute the two demonstration scripts:

   py run_mrd.py --workflow workflow_test.json --config config.csv

6. Inspect the newly created Runs/MRD_... folder:
   - manifest.json
   - execution.log
   - Traceability_Log.xlsx (three Excel worksheets)
   - *.stdout.log and *.stderr.log
   - workflow_snapshot.json and config_snapshot.csv

7. Optional failure test: edit workflow_test.json and append
   "--simulate-failure" to the command list for MRD_specific.
   Run command 5 again. The manifest should be FAILED and the launcher
   should exit with code 2.

How this works:
- workflow_test.json defines which scripts run and their arguments.
- run_mrd.py is a generic execution-and-recording launcher.
- fake_scripts/ are dummy analyses used only for testing.
- No analytical script needs modification.
- To configure Linux later, create a separate workflow_production.json with
  approved commands and conda environment handling.

Reminder: Changes to the workflow definition must be controlled/approved before
routine clinical use. A SUCCESS exit code is not analytical QC approval.
