#  Load Profile Data Processor

This code processes PVsyst energy generation data & Fluke Energy Logger data into 24 hours data for SEDA load profile application.

## Requirement
1. PVsyst energy generation data in .CSV
2. Fluke Energy Logger data in .txt
3. Valid output folder directory

## Instructions
1. Choose energy logger file
2. Choose PVsyst file (optional)
3. Choose output directory (default is current executed directory)
4. Enter company name and PV capacity (optional but recommended)
5. Enter description (optional)
6. Click `Submit`

## Results
1. Generate processed load profile in excel format
2. Generate load profile graph
3. Generate load profile vs inverter generation graph (if PVsyst file is added)
