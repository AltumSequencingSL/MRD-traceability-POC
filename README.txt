MRD Traceability Proof of Concept 


How this works:
- workflow_test.json defines which scripts run and their arguments.
- run_mrd.py is a generic execution-and-recording launcher.
- fake_scripts/ are dummy analyses used only for testing.
- No analytical script needs modification.
- To configure Linux later, create a separate workflow_production.json with
  approved commands and conda environment handling.

Reminder: Changes to the workflow definition must be controlled/approved before
routine clinical use. A SUCCESS exit code is not analytical QC approval.
