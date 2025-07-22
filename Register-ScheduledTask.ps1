<#

.SYNOPSIS
    Registers the UPS Shutdown Monitor task in Task Scheduler.

.DESCRIPTION
    This script registers the UPS Shutdown Monitor task in Windows Task Scheduler
    to run at system startup.

.NOTES
    Ensure you run this script with administrative privileges.
#>


# Check if running as administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script must be run as an administrator."
    exit -1
}

# Get the script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Update the xml file to include the ScriptDirectory
$xmlFilePath = "$scriptDir\win-auto-shutdown.xml"
if (-not (Test-Path $xmlFilePath)) {
    Write-Host "XML file not found: $xmlFilePath"
    exit -1
}

# Load the XML file
try {
    $xmlContent = Get-Content -Path $xmlFilePath -Raw -Encoding Unicode
    [xml]$taskXml = $xmlContent
} catch {
    Write-Host "Failed to load XML file: $($_.Exception.Message)"
    exit -1
}

# Update the script path in the XML
$taskXml.Task.Actions.Exec.Arguments = "-File `"$scriptDir\UPS-Shutdown-Monitor.ps1`""

# Create XML writer settings for proper formatting
$xmlWriterSettings = New-Object System.Xml.XmlWriterSettings
$xmlWriterSettings.Indent = $true
$xmlWriterSettings.IndentChars = "  "
$xmlWriterSettings.Encoding = [System.Text.Encoding]::Unicode # Use UTF-16 encoding
$xmlWriterSettings.OmitXmlDeclaration = $false

# Save the updated XML back to a temp file with proper encoding
$tempXmlFilePath = "$scriptDir\win-auto-shutdown-temp.xml"
try {
    $xmlWriter = [System.Xml.XmlWriter]::Create($tempXmlFilePath, $xmlWriterSettings)
    $taskXml.Save($xmlWriter)
    $xmlWriter.Close()
} catch {
    Write-Host "Failed to save temporary XML file: $($_.Exception.Message)"
    exit -1
}

    
try {
    $schtasksResult = & schtasks.exe /create /tn "win-auto-shutdown" /xml $tempXmlFilePath /f
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Task 'win-auto-shutdown' registered successfully using schtasks.exe."
    } else {
        throw "schtasks.exe failed with exit code: $LASTEXITCODE. Output: $schtasksResult"
    }
} catch {
    Write-Host "Method 2 also failed: $($_.Exception.Message)"
    Write-Host "Error details: $($_.Exception.GetType().FullName)"
    
    # Additional debug information
    if (Test-Path $tempXmlFilePath) {
        Write-Host "Temporary XML file exists at: $tempXmlFilePath"
        Write-Host "File size: $((Get-Item $tempXmlFilePath).Length) bytes"
        Write-Host "First few lines of XML file:"
        Get-Content -Path $tempXmlFilePath -Head 10 | ForEach-Object { Write-Host "  $_" }
    } else {
        Write-Host "Temporary XML file not found: $tempXmlFilePath"
    }
    
    Write-Host ""
    Write-Host "Manual import steps:"
    Write-Host "1. Open Task Scheduler (taskschd.msc)"
    Write-Host "2. Right-click 'Task Scheduler Library' and select 'Import Task...'"
    Write-Host "3. Browse to: $tempXmlFilePath"
    Write-Host "4. Name the task 'win-auto-shutdown' and click OK"
    
    exit -1
}

# Start the task immediately
try {
    & schtasks.exe /run /tn "win-auto-shutdown"
} catch {
    Write-Host "Failed to start task: $($_.Exception.Message)"
}

# Clean up the temporary XML file
Remove-Item -Path $tempXmlFilePath -ErrorAction SilentlyContinue -Force
Write-Host "Cleanup completed."