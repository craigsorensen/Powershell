<#
.SYNOPSIS
Search TDX by keyword and bulk close tickets and tasks

.DESCRIPTION
Search TDX by keyword and bulk close tickets and tasks

Author: Craig Sorensen
Version: 1.0
Date: 10/28/2019

.EXAMPLE
.\TicketSearch.ps1 -Environment dev
Search the TDX Sandbox by keyword and bulk close tickets and tasks

.\TicketSearch.ps1 -Environment prod
Search production TDX by keyword and bulk close tickets and tasks

.DISCLAIMER
All scripts and other powershell references are offered AS IS with no warranty.
These script and functions are tested in my environment and it is recommended that you test these scripts in a test environment before using in your production environment.
#>


[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Select TDX environment: dev, development, prod, production")]
    [String]$Environment
)

Write-Host -ForegroundColor Yellow "Select TDX environment: " -NoNewline
Write-Host "dev, development, prod, production"

switch ($Environment) {
    dev { $BaseURI = "https://<TDX_URL>/SBTDWebApi/api/"; Write-Host "Sandbox Selected." }
    development { $BaseURI = "<TDX_URL>/SBTDWebApi/api/"; Write-Host "Sandbox Selected." }
    prod { $BaseURI = "<TDX_URL>/TDWebApi/api/"; Write-Host "Production Selected." }
    production { $BaseURI = "<TDX_URL>/TDWebApi/api/"; Write-Host "Production Selected." }
    Default { Write-Host -ForegroundColor Red "Unknown environment specified. Exiting!"; exit }
}

# TDX Environment Parameters
$appID = ""
$username = ""
$password = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String("<Insert_Base64_Encoded_Password>"))

# Global Variables
$global:AuthToken | out-null
$TicketSearchResults = [System.Collections.ArrayList]@()

<#  UO TDX Ticket Status IDs

New: 17926
In Progress: 17928
WFC: 23328
Hold: 17932

#>


Function Log {
    param( [String]$LogString, [String]$LogLevel )

    # This is appends a log level tag to the write-host command. Used for printing basic strings.
    # Example: Log -LogLevel Error "This will print and error message."
    # Output: [Error] - This will print and error message.

    switch ($LogLevel) {
        "Error" { Write-Host -ForegroundColor Red "[Error] " -NoNewline; Write-Host $LogString ; break }
        "Warn" { Write-Host -ForegroundColor Yellow "[Warn] " -NoNewline; Write-Host $LogString ; break }
        "Info" { Write-Host "[Info] " -NoNewline; Write-Host $LogString ; break }
        default { Write-Host $LogString ; break }
    }
}

Function SleepForABit ([Int]$SleepTime) {
    # Function will make the script pause for X seconds.
    # Example: SleepForABit -SleepTime 30

    Log -LogLevel Warn "Sleeping for $SleepTime seconds.."
    $DefaultSleepTime = 10
    if (!$SleepTime) {
        Log -LogLevel Warn "Sleep time not specified, using default of $DefaultSleepTime seconds.."
        $SleepTime = $DefaultSleepTime
    }

    Start-Sleep -s $SleepTime
}

Function GET-TDXAPI ( [String]$EndPointPath, [String]$Method, [String]$Body, [String]$AuthToken) {
    <# This will call the TDX API and return the raw response. It also will catch and try to handle errors returned by the API. This can be configured
       by adding additional error codes and error handling methods.

       Example: GET-TDXAPI -EndPointPath "$appID/tickets/search" -Method GET -Body <body_in_json> -AuthToken <bearer_token>
    #>

    # Was a bearer token supplied with the function call? If not, request a new token from the API.
    if (!$global:AuthToken) {
        $global:AuthToken = Get-Token; $response = GET-TDXAPI -EndPointPath $EndPointPath -Method $Method -Body $Body -AuthToken $global:AuthToken;
    }

    $Headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $Headers.Add("Authorization", "Bearer " + $AuthToken)

    # write-host $global:AuthToken
    try {
        switch ($Method) {
            "get" { $response = Invoke-RestMethod -URI ($BaseURI + $EndPointPath) -Method GET -Headers $Headers -ContentType "application/json"; break }
            "patch" { $response = Invoke-RestMethod -URI ($BaseURI + $EndPointPath) -Method PATCH -Body $Body -Headers $Headers -ContentType "application/json"; break }
            "post" { $response = Invoke-RestMethod -URI ($BaseURI + $EndPointPath) -Method POST -Body $Body -Headers $Headers -ContentType "application/json"; break }
            "put" { $response = Invoke-RestMethod -URI ($BaseURI + $EndPointPath) -Method PUT -Body $Body -Headers $Headers -ContentType "application/json"; break }
            default { Log "Exiting! No API Method defined!" -LogLevel Error ; break }
        }
    }
    catch {
        Log "$($_.Exception.Response.StatusCode.value__) $($_.Exception.Response.StatusDescription)" -LogLevel Error
        $ErrorCode = $_.Exception.Response.StatusCode.value__
        Write-Host "Oops we've got an error here folks! Trying to handle it..."

        switch ($ErrorCode) {
            "401" { $global:AuthToken = Get-Token; $response = GET-TDXAPI -EndPointPath $EndPointPath -Method $Method -Body $Body -AuthToken $global:AuthToken; break }
            "429" { SleepForABit -SleepTime 15; $response = GET-TDXAPI -EndPointPath $EndPointPath -Method $Method -Body $Body -AuthToken $global:AuthToken; break }
            default { Log "Unable to handle the error. Exiting!" -LogLevel Error ; break }
        }

    }

    return $response
}

Function Get-Token {
    # Requests a new bearer token from the TDX API
    Write-Host "Requesting API Auth Token!"

    $creds = @{
        username = $username
        password = $password
    }

    $creds = $creds | ConvertTo-Json

    try {
        Invoke-RestMethod -URI ($BaseURI + "auth") -Method POST -Body $creds -Headers $headers -ContentType "application/json";
    }
    catch {
        Log "$($_.Exception.Response.StatusCode.value__) $($_.Exception.Response.StatusDescription)" -LogLevel Error
    }
}

Function Search-Tickets {
    param(
        [System.Collections.ArrayList]$AccountIDs,
        [nullable[Boolean]]$AssignmentStatus,
        [String]$ClosedDateFrom,
        [String]$ClosedDateTo,
        [String]$CreatedDateFrom,
        [String]$CreatedDateTo,
        [String]$CustomAttributes,
        [Int]$DaysOldFrom,
        [Int]$DaysOldTo,
        [nullable[Boolean]]$IsOnHold,
        [Int]$MaxResults,
        [String]$RequestorNameSearch,
        [System.Collections.ArrayList]$ResponsibilityGroupIDs,
        [System.Collections.ArrayList]$ResponsibilityUids,
        [String]$SearchText,
        [Int[]]$StatusIDs,
        [Int]$TicketID,
        [boolean]$ContainsString
    )
    # This function searches the TDX API using any of the following parameters. For more information on these parameters see:
    # https://api.teamdynamix.com/TDWebApi/Home/section/Tickets#POSTapi/{appId}/tickets/search

    # This function will search and try to match the searchtext to the text in the ticket title.
    # Returns ticket object with array of tickets.

    $Body = [ordered]@{ }

    if ($AccountIDs) { $Body.Add( "AccountIDs", $AccountIDs ) }
    if ($AssignmentStatus -ne $null) { $Body.Add( "AssignmentStatus", $AssignmentStatus ) }
    if ($ClosedDateFrom) { $Body.Add( "ClosedDateFrom", $ClosedDateFrom ) }
    if ($ClosedDateTo) { $Body.Add( "ClosedDateTo", $ClosedDateTo ) }
    if ($CreatedDateFrom) { $Body.Add( "CreatedDateFrom", $CreatedDateFrom ) }
    if ($CreatedDateTo) { $Body.Add( "CreatedDateTo", $CreatedDateTo) }
    if ($CustomAttributes) { $Body.Add( "CustomAttributes", $CustomAttributes ) }
    if ($DaysOldFrom) { $Body.Add( "DaysOldFrom", $DaysOldFrom ) }
    if ($DaysOldTo) { $Body.Add( "DaysOldTo", $DaysOldTo ) }
    if ($IsOnHold -ne $null) { $Body.Add( "IsOnHold", $IsOnHold) }
    if ($MaxResults) { $Body.Add( "MaxResults", $MaxResults ) }
    if ($RequestorNameSearch) { $Body.Add( "RequestorNameSearch", $RequestorNameSearch ) }
    if ($ResponsibilityGroupIDs) { $Body.Add( "ResponsibilityGroupIDs", $ResponsibilityGroupIDs ) }
    if ($ResponsibilityUids) { $Body.Add( "ResponsibilityUids", $ResponsibilityUids ) }
    if ($SearchText) { $Body.Add( "SearchText", $SearchText ) }
    if ($StatusIDs) { $Body.Add( "StatusIDs", $StatusIDs ) }
    if ($TicketID) { $Body.Add( "TicketID", $TicketID ) }


    $Body = $Body | ConvertTo-Json

    $Tickets = GET-TDXAPI -EndPointPath "$appID/tickets/search" -Method POST -Body $Body -AuthToken $global:AuthToken

    if ($ContainsString) {
        $ExactStringMatchTickets = @()
        foreach ($t in $Tickets) {

            if ($t.title -like "*$SearchText*") {
                $ExactStringMatchTickets += $t
                Write-Host $t.title $t.statusid -ForegroundColor green
            }

        }
        return $ExactStringMatchTickets
    }

    return $Tickets;
}

Function Search-TicketTasks {
    param(
        [Parameter(Mandatory = $true)]
        [Int]$TicketID)
    # Take a ticked (by ticket ID) and return any sub-tasks, if they exist
    GET-TDXAPI -EndPointPath "$appID/tickets/$TicketID/tasks/" -Method GET -AuthToken $global:AuthToken
}

Function Close-TicketTask {
    param(
        [Parameter(Mandatory = $true)]
        [Int]$TicketID,
        [string]$TaskID
    )
    # Will take a ticketID and taskID then close the associated task.
    $Body = @{ }

    #Get all ticket tasks
    $tasks = Search-TicketTasks -TicketID $TicketID

    # Grab properties from the ticket task. These are needed when closing the task, otherwise the API will set the values to null when the ticket is closed.
    foreach ($t in $tasks) {
        if ($t.ID -eq $TaskID) {

            $Body.Add( "Title", $t.Title )
            $Body.Add( "Description", $t.Description )
            $Body.Add( "StartDate", $t.StartDate )
            $Body.Add( "EndDate", ((get-date).ToUniversalTime()).ToString("yyyy-MM-ddTHH:mm:ssZ") )
            $Body.Add( "CompleteWithinMinutes", $t.CompleteWithinMinutes )
            $Body.Add( "EstimatedMinutes", $t.EstimatedMinutes )
            $Body.Add( "PercentComplete", 100 )
            $Body.Add( "ResponsibleUid", $t.ResponsibleUid )
            $Body.Add( "ResponsibleFullName", $t.ResponsibleFullName )
            $Body.Add( "ResponsibleEmail", $t.ResponsibleEmail )
            $Body.Add( "ResponsibleGroupID", $t.ResponsibleGroupID )
            $Body.Add( "Order", $t.Order )
        }
    }

    $Body = $Body | ConvertTo-Json

    GET-TDXAPI -EndPointPath "$appID/tickets/$TicketID/tasks/$TaskID" -Method PUT -Body $Body -AuthToken $global:AuthToken
}

Function Close-Ticket {
    param(
        [Parameter(Mandatory = $true)]
        [Int]$TicketID
    )
    # Close ticket by ticketID, using PATCH method

    # Format: [{ "op": "add", "path": "/StatusID", "value": 17930}]
    $Body = [ordered]@{ }

    $Body.Add( "op", "add" )
    $Body.Add( "path", "/StatusID" )
    $Body.Add( "value", 17930 )

    $Body = $Body | ConvertTo-Json

    GET-TDXAPI -EndPointPath "$appID/tickets/$TicketID" -Method PATCH -Body "[$Body]" -AuthToken $global:AuthToken
}

Function PopulateListView([String]$SearchCriteria, $ExcludeClosedTickets) {
    # This is used to populate the ListView GUI object with the search results. It also controls the way the list view object behaves.

    if ($ExcludeClosedTickets) { $ExcludeClosedTickets = @(17926, 17928, 23328, 17932) }

    Log -LogLevel Info "Excluding Closed tickets!" #$ExcludeClosedTickets
    Log -LogLevel Info "Search Criteria: $SearchCriteria"


    $ResultsListView.Items.Clear();
    $ResultsListView.Columns.Clear();
    $Progress_Tickets.Value = 0



    $SearchResults = Search-Tickets -SearchText $SearchCriteria -ContainsString $true -StatusIDs $ExcludeClosedTickets

    Write-Host "Found: " $SearchResults.Count " Tickets"
    $ticketTotalCount = $SearchResults.Count
    $Progress_Tickets.Maximum = $ticketTotalCount

    # Column heading, -2 means to auto resize the column width to the search results.
    $ResultsListView.Columns.Add("Ticket ID", -2)
    $ResultsListView.Columns.Add("Requestor", -2)
    $ResultsListView.Columns.Add("Title", -2)
    $ResultsListView.Columns.Add("Ticket Status", -2)
    $ResultsListView.Columns.Add("Task Exists", -2)

    Foreach ($ticket in $SearchResults) {
        $Progress_Tickets.PerformStep();
        $tasks = Search-TicketTasks -TicketID $ticket.ID

        if ($tasks) {
            $ticket.tasks = $tasks
            $ticket | Add-Member -Name "HasTasks" -Value "True" -MemberType NoteProperty
        }
        else {
            $ticket | Add-Member -Name "HasTasks" -Value "False" -MemberType NoteProperty
        }
        $TicketSearchResults.Add($ticket)

        #write-host "Tasks " $ticket.HasTasks

        $ListView_Item = New-Object System.Windows.Forms.ListViewItem($ticket.ID)
        $ListView_Item.SubItems.Add($ticket.RequestorName) | Out-Null
        $ListView_Item.SubItems.Add($ticket.Title) | Out-Null
        $ListView_Item.SubItems.Add($ticket.StatusName) | Out-Null
        $ListView_Item.SubItems.Add($ticket.HasTasks)
        $ResultsListView.Items.AddRange(($ListView_Item))
        $ResultsListView.AutoResizeColumns("HeaderSize")
    }
}

Function SearchBarActions() {
    # On enter code here?
    # Would be nice to have the ticket search be submitted on button push (enter).
}

# Build the GUI

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$CloseTicketForm = New-Object system.Windows.Forms.Form
$CloseTicketForm.Text = "TDX Bulk Ticket Manager"
$CloseTicketForm.ControlBox = $True
$CloseTicketForm.TopMost = $False
$CloseTicketForm.ClientSize = '870,700'

$SearchCriteriaLabel = New-Object system.Windows.Forms.Label
$SearchCriteriaLabel.text = "Search Criteria"
$SearchCriteriaLabel.AutoSize = $true
$SearchCriteriaLabel.width = 25
$SearchCriteriaLabel.height = 10
$SearchCriteriaLabel.location = New-Object System.Drawing.Point(10, 21)
$SearchCriteriaLabel.Font = 'Microsoft Sans Serif,10'

$SearchCriteriaTextBox = New-Object system.Windows.Forms.TextBox
$SearchCriteriaTextBox.multiline = $false
$SearchCriteriaTextBox.width = 532
$SearchCriteriaTextBox.height = 18
$SearchCriteriaTextBox.location = New-Object System.Drawing.Point(105, 18)
$SearchCriteriaTextBox.Font = 'Microsoft Sans Serif,10'

$SearchButton = New-Object system.Windows.Forms.Button
$SearchButton.text = "Search"
$SearchButton.width = 60
$SearchButton.height = 22
$SearchButton.location = New-Object System.Drawing.Point(645, 18)
$SearchButton.Font = 'Microsoft Sans Serif,10'
$SearchButton.Add_Click( {
        PopulateListView -SearchCriteria $SearchCriteriaTextBox.text -ExcludeClosedTickets ($CheckBox_ExcludeClosedTickets.Checked.ToString() -eq [bool]::TrueString)
        #Write-Host $TicketSearchResults
    })

$CheckBox_ExcludeClosedTickets = New-Object system.Windows.Forms.CheckBox
$CheckBox_ExcludeClosedTickets.text = "Exclude closed"
$CheckBox_ExcludeClosedTickets.AutoSize = $false
$CheckBox_ExcludeClosedTickets.Checked = $True
$CheckBox_ExcludeClosedTickets.width = 130
$CheckBox_ExcludeClosedTickets.height = 20
$CheckBox_ExcludeClosedTickets.location = New-Object System.Drawing.Point(105, 47)
$CheckBox_ExcludeClosedTickets.Font = 'Microsoft Sans Serif,10'

$Tip_ExcludeClosedTickets = New-Object system.Windows.Forms.ToolTip
$Tip_ExcludeClosedTickets.ToolTipTitle = "Exclude Closed Tickets"
#$Tip_ExcludeClosedTickets.isBalloon  = $true

$Tip_ExcludeClosedTickets.SetToolTip($CheckBox_ExcludeClosedTickets, 'This will exclude closed tickets from the search results')

$ResultsListView = New-Object system.windows.Forms.ListView
$ResultsListView.CheckBoxes = $true
$ResultsListView.View = "Details"
$ResultsListView.Text = "ListView"
$ResultsListView.Font = "Microsoft Sans Serif,10"
$ResultsListView.location = New-Object System.Drawing.Point(10, 80)
$ResultsListView.Width = 850
$ResultsListView.Height = 580
$ResultsListView.Sorting.SortOrder.Ascending;

$Progress_Tickets = New-Object system.Windows.Forms.ProgressBar
$Progress_Tickets.width = 720
$Progress_Tickets.height = 21
$Progress_Tickets.location = New-Object System.Drawing.Point(10, 670)
$Progress_Tickets.Visible = $true;
$Progress_Tickets.Minimum = 0;
$Progress_Tickets.Value = 0;
$Progress_Tickets.Step = 1;

$CloseTicketsButton = New-Object system.Windows.Forms.Button
$CloseTicketsButton.text = "Close Tickets"
$CloseTicketsButton.width = 88
$CloseTicketsButton.height = 21
$CloseTicketsButton.location = New-Object System.Drawing.Point(755, 670)
$CloseTicketsButton.Font = 'Microsoft Sans Serif,10'

$CloseTicketForm.controls.AddRange(@($SearchCriteriaTextBox, $SearchCriteriaLabel, $SearchButton, $CheckBox_ExcludeClosedTickets, $Progress_Tickets, $CloseTicketsButton, $ResultsListView))

$CloseTicketsButton.Add_Click(
    {
        foreach ($item in $ResultsListView.Checkeditems) {
            Write-Host $item.Text
            $tasks = Search-TicketTasks -TicketID $item.Text

            if ($tasks) {
                foreach ($task in $tasks) {
                    Write-Host "Closing task " $task.ID " on ticket " $item.Text
                    Close-TicketTask -TicketID $item.Text -TaskID $task.ID
                }
            }
            $tasks = Search-TicketTasks -TicketID $item.Text

            foreach ($task in $tasks) {
                if ($task.PercentComplete -ne 100) {
                    Write-Host
                }

            }
            if (!$tasks) {
                Write-Host "Closing ticket called for ticket " $item.Text
                Close-Ticket -TicketID $item.Text
            }
            else {
                Log "Tasks still present. Will skip setting ticket to closed!" -LogLevel Warn
            }
        }


    }
)

# Draw the GUI
[void]$CloseTicketForm.ShowDialog()

#$TicketSearchResults[0].ID
