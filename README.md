# win-auto-shutdown
Windows Auto Shutdown Script for power failures

This simple PowerShell script monitors the attached (USB) battery and runs the configured shutdown script if battery level drops below the supplied threshold.

# Configuration

The configuration file also includes helpful descriptions for each parameter to make it easy to understand what each setting controls. To use this configuration, you would:

 - Copy config-example.json to config.json

 - Adjust the values according to your specific needs

The PowerShell script will automatically load and use these settings
This configuration provides a good balance between monitoring frequency and system resources while ensuring timely shutdown during power outages.