<#
.Synopsis
   Stop Services.
.NOTES
   Created by: Craig Sorensen
   Modified: 12/18/19

   Changelog:
    * Initial release
#>


# Configurable variables

$MaxLogFileSize = 1024  # Set the max file size in KB
$LogfileName = "Service-Manager.log" # Set the logfile name

# $ServiceName accepts wildcards and will search the SERVICE DISPLAYNAME for a match! Service displayname is what you see in the services mmc for name.
# Be advised: This can be different from the actual service name.
# If you need to look at the service name field, modify the line below from $_.DisplayName to $_.Name
# $servicearray = Get-Service | Where-Object {$_.Name -like "$($ServiceName)"}

$ServiceName = "*Spooler*"




function Get-ScriptDirectory {
    Split-Path -parent $PSCommandPath
}

function Create-LogFile {
    If (!(Test-Path $logfile -PathType Leaf)) {
        New-Item $logfile
        Write-Host -ForegroundColor Yellow "Logfile Not found. Created new log file: " $Logfile
    }
}

function Write-Log {
    # Source: https://gallery.technet.microsoft.com/scriptcenter/Write-Log-PowerShell-999c32d0
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [Alias('LogPath')]
        [string]$Path = 'C:\Logs\PowerShellLog.log',

        [Parameter(Mandatory = $false)]
        [ValidateSet("Error", "Warn", "Info")]
        [string]$Level = "Info",

        [Parameter(Mandatory = $false)]
        [switch]$NoClobber
    )

    Begin {
        # Set VerbosePreference to Continue so that verbose messages are displayed.
        $VerbosePreference = 'Continue'
    }
    Process {

        # If the file already exists and NoClobber was specified, do not write to the log.
        if ((Test-Path $Path) -AND $NoClobber) {
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name."
            Return
        }

        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path.
        elseif (!(Test-Path $Path)) {
            Write-Verbose "Creating $Path."
            $NewLogFile = New-Item $Path -Force -ItemType File
        }

        else {
            # Nothing to see here yet.
        }

        # Format Date for our Log File
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Write message to error, warning, or verbose pipeline and specify $LevelText
        switch ($Level) {
            'Error' {
                Write-Error $Message
                $LevelText = 'ERROR:'
            }
            'Warn' {
                Write-Warning $Message
                $LevelText = 'WARNING:'
            }
            'Info' {
                Write-Verbose $Message
                $LevelText = 'INFO:'
            }
        }

        # Write log entry to $Path
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append
    }
    End {
    }
}

$ScriptExcutionPath = Get-ScriptDirectory
$logfile = $ScriptExcutionPath + "\" + $LogfileName



# Check if logfile exists, if not create it.
Create-LogFile

# Get current log file size
$FileSize = (Get-Item $logfile).length / 1KB

# Check if the current log is over max size and needs to be rolled.
If ($FileSize -gt $MaxLogFileSize) {
    Write-log -Level Warn -Path $logfile "Log file reached max size: $FileSize"
    Write-log -Level Info -Path $logfile "Rolling log file.. "

    If ((Test-Path ($logfile + ".1") -PathType Leaf)) {
        Write-log -Level Warn -Path $logfile "backup file found, will remove!"
        Remove-Item ($logfile + ".1")
    }

    Rename-Item -Force -path $logfile -newname ($LogfileName + ".1")
    # Check if logfile exists, if not create it.
    Create-LogFile
    Write-log -Level Info -Path $logfile "Logfile reached maximum configured size. Logfile has been rolled.."

}

$servicearray = Get-Service | Where-Object { $_.DisplayName -like "$($ServiceName)" }

for ($i = 0; $i -lt $servicearray.Length; $i++) {
    Write-log -Level Info -Path $logfile "Starting $($servicearray[$i].DisplayName) Service"
    Start-Service $servicearray[$i].Name
}
