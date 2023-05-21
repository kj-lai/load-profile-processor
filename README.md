#  Load Profile Data Processor

This code processes PVsyst energy generation data & Fluke Energy Logger data into 24 hours data for SEDA load profile application.

To convert the python file into executable .exe, please use 'pyinstaller'. The file size will be around 53.2MB which has over the GitHub's file size limit.

## Requirement
1. PVsyst energy generation data in .CSV
2. Fluke Energy Logger data in .txt
3. Valid output folder directory

## Instructions
1. Choose the two input files and output folder directory
2. Enter the start date and end date (end date need to add 1 day extra)
3. Enter company name and PV capacity
4. Enter description (optional)
5. Click `Submit`

## Future Work
1. Adding graph generating feature
