<#
.SYNOPSIS
    This script monitors the UPS battery status and performs actions based on the battery status.
    It logs the status to a file and shuts down the system if the estimated runtime is below a specified threshold.

.NOTES
    This script was written with the help of AI, please review and test thoroughly before use in production environments.
    Ensure you have the necessary permissions to run shutdown commands on remote machines if configured.

#>


function Invoke-ShutdownScript () {
    Write-Host "Shutting Down"
    # Check if computer names are configured, otherwise shutdown local machine
    if ($config.ComputerNames -and $config.ComputerNames.Length -gt 0) {
        foreach ($computer in $config.ComputerNames) {
            Write-Host "Shutting down $computer"
            Stop-Computer -ComputerName $computer -Force -ErrorAction Continue
        }
    } else {
        Write-Host "Shutting down local machine"
        #Stop-Computer -Force
    }
    Exit
}


# Import the configuration file
$configFilePath = "$($PSScriptRoot)\config.json"

if (Test-Path $configFilePath) {
    try {
        $config = Get-Content -Path $configFilePath | ConvertFrom-Json
    }
    catch {
        Write-Host "Error parsing configuration file: $($_.Exception.Message)"
        exit -1
    }
} else {
    Write-Host "Configuration file not found: $configFilePath"
    exit -1
}

# Validate configuration values
if (-not $config.ShutDownRunTime -or $config.ShutDownRunTime -lt 1) {
    Write-Host "Invalid ShutDownRunTime in configuration. Must be greater than 0."
    exit -1
}

if (-not $config.LogUpdateIntervalMinutes -or $config.LogUpdateIntervalMinutes -lt 1) {
    Write-Host "Invalid LogUpdateIntervalMinutes in configuration. Must be greater than 0."
    exit -1
}

if (-not $config.LogOnBatteryIntervalSeconds -or $config.LogOnBatteryIntervalSeconds -lt 1) {
    Write-Host "Invalid LogOnBatteryIntervalSeconds in configuration. Must be greater than 0."
    exit -1
}

if (-not $config.SleepIntervalSeconds -or $config.SleepIntervalSeconds -lt 1) {
    Write-Host "Invalid SleepIntervalSeconds in configuration. Must be greater than 0."
    exit -1
}

#Below this amount of runtime the script will run Invoke-ShutdownScript
$ShutDownRunTime = $config.ShutDownRunTime

#Log update interval (in minutes)
$logUpdateIntervalMinutes = $config.LogUpdateIntervalMinutes

#Log on battery interval (in seconds)
$logOnBatteryIntervalSeconds = $config.LogOnBatteryIntervalSeconds

# Sleep interval (in seconds)
$sleepIntervalSeconds = $config.SleepIntervalSeconds

#Store time variable
$LastLogTime = Get-Date 0

# Ensure log directory exists
$logDirectory = "$($PSScriptRoot)\log"
if (-not (Test-Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}


# Get current date and time
$currentDate = Get-Date -Format "yyyy-MM-dd"

# Log file path - calculate once outside the loop
$script:logFilePath = "$logDirectory\$($currentDate)_battery_status_log.txt"

trap {
    $errorInfo = $_.Exception
    $errorMessage = "Error: " + $errorInfo.GetType().FullName
    $errorMessage += "`nMessage: " + $errorInfo.Message
    $errorMessage += "`nStackTrace: " + $errorInfo.StackTrace
    $errorMessage += "`n"

    # Append the error details to the log file
    try {
        $errorMessage | Out-File -FilePath $script:logFilePath -Append
    }
    catch {
        Write-Host "Failed to write error to log file: $errorMessage"
    }
    exit -1
}



# Write initial log entry
$currentTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "[$currentTime] UPS Shutdown Monitor started."
"[$currentTime] UPS Shutdown Monitor started." | Out-File -FilePath $script:logFilePath -Append

# Main monitoring loop wrapped in try-finally for cleanup
try {
    # Loop to check UPS battery status
    while ($true) {
        # Get current date and time
        $currentTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $currentDate = Get-Date -Format "yyyy-MM-dd"

        # Update log file path if date has changed (handles overnight runs)
        $newLogFilePath = "$logDirectory\$($currentDate)_battery_status_log.txt"
        if ($newLogFilePath -ne $script:logFilePath) {
            $script:logFilePath = $newLogFilePath
        }    

        # Get UPS battery information using CIM instead of deprecated WMI
        try {
            $ups = Get-CimInstance -ClassName Win32_Battery -ErrorAction Stop
        }
        catch {
            $errorMessage = "[$currentTime] Error retrieving UPS information: $($_.Exception.Message)"
            Write-Host $errorMessage
            $errorMessage | Out-File -FilePath $script:logFilePath -Append
            Start-Sleep -Seconds $sleepIntervalSeconds
            continue
        } 
        
        if ($ups) {
            # Handle potential null EstimatedRunTime
            $estimatedRunTime = if ($ups.EstimatedRunTime -ne $null) { $ups.EstimatedRunTime } else { "Unknown" }
            $statusMessage = "[$currentTime] UPS Battery Status: $($ups.BatteryStatus), EstimatedRunTime: $estimatedRunTime"
            Write-Host $statusMessage            

            if ($ups.BatteryStatus -ne 2) {
                #Write to log if interval has elapsed
                if (((Get-Date) - $LastLogTime).TotalSeconds -gt $logOnBatteryIntervalSeconds) {
                    # Log status message to file
                    $statusMessage | Out-File -FilePath $script:logFilePath -Append
                    $LastLogTime = Get-Date
                } 

                #Run Invoke-ShutdownScript if Estimated time below $ShutDownRunTime
                if ($ups.EstimatedRunTime -ne $null -and $ups.EstimatedRunTime -lt $ShutDownRunTime) {
                    $statusMessage = "[$currentTime] UPS Runtime ($($ups.EstimatedRunTime)) is below $($ShutDownRunTime) minutes. Shutting Down"
                    Write-Host $statusMessage
                    $statusMessage | Out-File -FilePath $script:logFilePath -Append
                    Invoke-ShutdownScript
                }
            } else {
                #Write to log if interval has elapsed
                if (((Get-Date) - $LastLogTime).TotalMinutes -gt $logUpdateIntervalMinutes) {
                    # Log status message to file
                    $statusMessage | Out-File -FilePath $script:logFilePath -Append
                    $LastLogTime = Get-Date
                } 
            }
        }
        else {
            $noUpsMessage = "[$currentTime] No UPS battery found on this system."

            # Write no UPS message to console
            Write-Host $noUpsMessage

            # Log no UPS message to file
            $noUpsMessage | Out-File -FilePath $script:logFilePath -Append

        }

        # Wait for sleep interval before checking again
        Start-Sleep -Seconds $sleepIntervalSeconds
    }
}
finally {
    # Cleanup code that runs when script exits (normally or via CTRL-C)
    $currentTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $exitMessage = "[$currentTime] UPS Shutdown Monitor exiting."
    
    $exitMessage | Out-File -FilePath $script:logFilePath -Append    
    Write-Host $exitMessage
    
}