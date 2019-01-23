# ClinVarEPReports
Python scripts to generate ClinGen Expert Panel reports from ClinVar FTP files.

## About this project
ClinVar outputs a submission_summary.txt.gz file containing a list of all SCV submissions per variant to its FTP site.
The scripts in this project use this file to generate the following files in the subdirectory ClinVarEPReports/Reports_MM-DD-YYYY:

**EPReports.py** - this script outputs an Excel file containing each variant in the submission_summary.txt file that an EP needs to update or review. The Excel contains a README with summary stats and 7 structured tabs as detailed below:
  * \#1. Alert: ClinVar variants with an LP/VUS Expert Panel SCV with a DateLastEvaluated > 2 years from the date of this file (may overlap with variants on Tabs 2, 3 and 4).
  * \#2. Alert: ClinVar variants with a P/LP Expert Panel SCV AND a newer VUS/LB/B non-EP SCV (with a DateLastEvaluated up to 1 year prior of EP DateLastEvaluated; medically-significant conflict).
  * \#3. Alert: ClinVar variants with a VUS Expert Panel SCV AND a newer P/LP non-EP SCV (with a DateLastEvaluated up to 1 year prior of EP DateLastEvaluated).
  * \#4. Alert: ClinVar variants with a VUS Expert Panel SCV AND a newer LB/B non-EP SCV (with a DateLastEvaluated up to 1 year prior of EP DateLastEvaluated).
  * \#5. Priority: ClinVar variants WITHOUT an Expert Panel SCV, but in a gene in scope for the EP, with at least one P/LP SCV and at least one VUS/LB/B SCV (medically-significant conflict).')
  * \#6. Priority: ClinVar variants WITHOUT an Expert Panel SCV, but in a gene in scope for the EP, with at least one VUS SCV and at least one LB/B SCV.
  * \#7. Priority: ClinVar variants WITHOUT an Expert Panel SCV, but in a gene in scope for the EP, with >=3 concordant VUS SCVs from different submitters.
  * \#8. Priority: ClinVar variants WITHOUT an Expert Panel SCV, but in a gene in scope for the EP, with at least one P/LP SCV from (at best) a no assertion criteria provided submitter.

**EPReports.py** also generates an EPReportsStats Excel file containing the summary variant counts for each EP.

## How to run these scripts
All scripts are run as 'python3 *filename.py*'.
All scripts use FTP to take the most recent ClinVar FTP files as input and to output the files with the date of the FTP submission_summary.txt.gz file appended:

  * ftp.ncbi.nih.gov/pub/clinvar/tab_delimited/submission_summary.txt.gz
  * ftp.ncbi.nih.gov/pub/clinvar/tab_delimited/variation_allele.txt.gz
  * ftp.ncbi.nih.gov/pub/clinvar/tab_delimited/variant_summary.txt.gz

These ClinVar files are then removed when finished.
