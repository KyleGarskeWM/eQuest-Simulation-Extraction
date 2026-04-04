$ErrorActionPreference = "Stop"

$RepoPath = "A:\Users\kyleg\PycharmProjects\eQuest-Simulation-Extraction"
$SimPath = "A:\Users\kyleg\PycharmProjects\eQuest-Simulation-Extraction\sample_data\St Anselm Baseline ABS_Rev_0 - Baseline Design.SIM"
$WorkbookPath = "A:\Users\kyleg\PycharmProjects\eQuest-Simulation-Extraction\output_files\Building Performance Assumptions.xlsm"
$MasterRoomOut = "A:\Users\kyleg\PycharmProjects\eQuest-Simulation-Extraction\output_files\Building Performance Assumptions.master_room.updated.xlsm"
$EcmOut = "A:\Users\kyleg\PycharmProjects\eQuest-Simulation-Extraction\output_files\Building Performance Assumptions.ecm.updated.xlsm"

Set-Location $RepoPath

python --version
python equest_extractor.py --help

# 1) Extract BEPS JSON sanity check
python equest_extractor.py "$SimPath" --report beps

# 2) Populate Master Room List (Baseline)
python equest_extractor.py "$SimPath" --populate-master-room-list "$WorkbookPath" --model-run-type Baseline --output-workbook "$MasterRoomOut"

# 3) Populate ECM Data (ECM-1)
python equest_extractor.py "$SimPath" --update-ecm-data "$WorkbookPath" --model-run-type ECM-1 --output-workbook "$EcmOut"

Write-Host "Done. Output files:"
Write-Host $MasterRoomOut
Write-Host $EcmOut
