# Create an Outlook application instance
$outlook = New-Object -ComObject Outlook.Application

# Create a new mail item
$mail = $outlook.CreateItem(0)

# Set mail item properties
$mail.Subject = "Weekly File"
$mail.Body = "Here is the weekly data .zip file using AutoDataCollector <version002>"
$mail.To = "tal.elazar@hp.com; daniel.biton@hp.com"

# Get the directory where the script is located
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Define the folder to compress and the name of the zip file
$folderToCompress = Join-Path $scriptPath "data"
$today = Get-Date -Format "yyyyMMdd"
$zipFileName = "data_$today.zip"
$zipFilePath = Join-Path $scriptPath $zipFileName

# Check if the zip file already exists and delete it
if (Test-Path $zipFilePath) {
    Remove-Item $zipFilePath
    Write-Host "Existing zip file found and deleted."
}

# Compress the folder
Compress-Archive -Path $folderToCompress -DestinationPath $zipFilePath

Write-Host "Full path to file: $zipFilePath"

# Add the attachment
$mail.Attachments.Add($zipFilePath)

# Send the email
$mail.Send()
