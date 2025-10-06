<# 
Set Room Calendar Processing Book In Policy Group from CSV File
    Version: v1.0
    Date: 06/10/2025
    Author: Rob Watts https://github.com/robwatts365
#>

# Import Exchange Online Module
Import-Module ExchangeOnlineManagement

# Connects to Exchange Online
Write-Host "Connecting to Exchange Online..."
Connect-ExchangeOnline

# Enable File Picker
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

# File Picker  (Set File Path - Open File Browser)
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $FilePath = $OpenFileDialog.filename
   
# Store the data from NewUsersFinal.csv in the $ADUsers variable
    Write-Host "Importing CSV..."
    $Spaces = Import-Csv $FilePath

# Loop through each row containing user details in the CSV file
foreach ($Space in $Spaces) {

    # Read Space data from each field in each row and assign the data to a variable as below
    $SpaceID = $Space.SpaceID
    $GroupEmail = $Space.GroupEmail
        
    Set-CalendarProcessing -Identity "$SpaceID" -AutomateProcessing AutoAccept -BookInPolicy "$GroupEmail" -AllBookInPolicy $false

    }
