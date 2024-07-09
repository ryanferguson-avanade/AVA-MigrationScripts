<#
.SYNOPSIS
    This script will download log analytics data and save the output to a CSV file.  The directory structure that it saves the files to is defined in the CSV input
    The ServerList input takes the form of a list of Servers
    You will need to have access to the log analytics workspace in order for this script to function correctly!

.DESCRIPTION
    Usage: .\ExportLogAnalyticsData.ps1 -ServerList serverlist.csv -OutputDir "Output directory Path" -CheckType Pre -Timespan 30d/14d/7d/24h/1h -SubscriptionName  "name of the subscription that LogAnalytics resides" -WorkspaceRG "name of the LA resource group" -WorkspaceName "name of the Log Analytics Workspace" 
    -ServerList serverlist.csv          CSV or text file containing a list of servers Servers - one per line.  Use -GenerateSampleCSV to generate a template file to use.
                                        Default is serverlist.csv
    -OutputDir                          Specify the directory for the output files.
                                        Default is .\ (the current directory)
    -CheckType [Pre/Post/Verify         This sets the checks to either pre or post migration.  This is important for the naming of files.
                                        [Pre] -> This will set the check type to Pre Migration which will output the Pre_LogAnalytics file as well as the PortScanTest.csv
                                        [Post] -> This will set the check type to Post Migration which will output the Pre_LogAnalytics file
                                        [Verify] -> This will simply verify if the servers in serverlist.csv have data in LA.  If they don't, the server is written to errorlist.csv
    -Timespan  [##d/##h]         ]      This sets how long the query will look back.  You can use any number followed by d or h.  30d = 30 days, 14d = 14 days, 7d = 7 days, 24h = 24 hours, 1h = 1 hour.  
                                        Default is 30d.       
    -SubscriptionName                   This is the name of the subscription that LogAnalytics resides
    -WorkspaceRG                        This is the name of the resource group that holds the log analytics workspace
    -WorkspaceName                      This is the name of the Log Analytics Workspace we will query
    -ErrorList ErrorList.csv            This is the name of the files that servers with no data will be written to.  ErrorList.csv can be used as an input to this script to re-run the errors.
                                        Default is ErrorList.csv
    -ServerName                         This is the name of the server you want to run the script against.  This will override the serverlist.csv file.                                        
    -GenerateSampleCSV                  This will generate a sample serverlist.csv so you can modify it.                               
    -Help                               These usage instructions are displayed.
.LINK
    Discovery Toolkit:  https://aka.avanade.com/Discovery

.EXAMPLE
    .\ExportLogAnalyticsData.ps1 -CheckType Post -OutputDir C:\Temp
    This will generate the Post Migration Log Analytics output using serverlist.csv as an input and files are written to c:\temp
.EXAMPLE
    .\ExportLogAnalyticsData.ps1 -ServerList MyServerList.csv -CheckType Pre
    This will generate the Pre Migration Log Analytics output using MyServerList.csv as an input file
.EXAMPLE
    .\ExportLogAnalyticsData.ps1 -GenerateSampleCSV
    This will create a new serverlist.csv file that you can use as a template.  It will not overwrite a file if it is there already.
.NOTES
    Organization    - Avanade
    Owner           - Cloud Managed Services
    Author          - Ryan Ferguson
    ScriptVersion   - 2024.04.03
#>

Param(
    [string]$ServerList = 'serverlist.csv',
    [Parameter(Mandatory = $false)]
    [string]$CheckType,
    [ValidatePattern('^\d+[hd]$')][String]$Timespan = '30d',
    # Change this value for each project.  This will generally only need to be set once so make it default.
    [string]$SubscriptionName,
    # Change this value for each project.  This will generally only need to be set once so make it default.
    [string]$WorkspaceRG,
    # Change this value for each project.  This will generally only need to be set once so make it default.
    [string]$WorkspaceName,
    [string]$OutputDir = '.\',
    [string]$ErrorList = 'ErrorList.csv',
    [switch]$GenerateSampleCSV,
    [string]$ServerName,
    [switch]$Help
)

if ($GenerateSampleCSV) {
    write-host "Generating sample CSV:  $ServerList"
    $Header = "Server`n"
    $Header += """localhost""`n"
    $Header | out-file $ServerList -NoClobber -encoding UTF8 
    return
}


elseif ($Help ) {
    # -or ($PSBoundParameters.count -eq 0)) {
    write-host 'ERROR: No CheckType Selected.' -ForegroundColor Red
    Get-Help $MyInvocation.MyCommand.Definition
    return
}

# This function will create a directory provided it doesn't already exist
function CreateDir {
    param(
        [string]$NewPath  
    )
    If (!(test-path $NewPath)) {
        # write-host "Creating $NewPath"
        New-Item -ItemType directory -force -Path $NewPath | Out-Null
    } 
}

# This function downloads data from Log Analytics.  Since this is done multiple times, we put this into a function.
function DownloadData {
    param(
        [string]$Query
    )
    if ($Timespan.Contains('d')) {
        $days = $Timespan.Replace('d', '')
        $DownloadData = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceID -Query $Query -Timespan (New-TimeSpan -days $days)
        #$ResultCount = $($result.results | select-object Computer | measure).count
    }
    else {
        $hours = $Timespan.Replace('h', '')
        $DownloadData = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceID -Query $Query -Timespan (New-TimeSpan -hours $hours)
        #$ResultCount = $($result.results | select-object Computer | measure).count     
    }
    Return $DownloadData
}



if ([string]::IsNullOrWhiteSpace($SubscriptionName)) {
    Write-Host "Subscription name is empty." -ForegroundColor Red
    return
}

if ([string]::IsNullOrWhiteSpace($WorkspaceName)) {
    Write-Host "Workspace name is empty." -ForegroundColor Red
    return
}

if ([string]::IsNullOrWhiteSpace($WorkspaceRG)) {
    Write-Host "Workspace resource group is empty." -ForegroundColor Red
    return
}


if ([string]::IsNullOrWhiteSpace($ServerName)){
    $ServerCSVList = Import-Csv -Path $ServerList -ErrorAction SilentlyContinue 
}
else {
    $ServerCSVList = @{
        Server = $ServerName}
    
}

if (!$ServerCSVList) {
    write-host 'There are no servers within the file.'
    return
}
else {
    $Length = $($ServerCSVList | Measure-object).count
    write-host 'Servers listed in CSV: ' $Length -ForegroundColor Green
}
write-host 'CheckType: ' $CheckType
# Check if subscription name, workspace name, or workspace resource group is empty

# This will open an interactive login screen to authenticate to Azure
Connect-AzAccount
# If there are multiple subscriptions, we need to select the one that has the workspace we need
Select-AzSubscription $SubscriptionName
# Grab the workspace ID using the information provided
$WorkspaceID = (Get-AzOperationalInsightsWorkspace -Name $workspaceName -ResourceGroupName $workspaceRG).CustomerID

if ($CheckType -eq "Verify") {
    $Query = "VMConnection
    | summarize count() by Computer
    | where count_ > 0"
    $Timespan = '1h'
    write-host "Getting Results... "
    $result = DownloadData $Query
    $LogAnalyticsServers = @()
    foreach ($computer in $result.results.computer) {
        $LogAnalyticsServers += $computer.Split('.')[0].ToUpper()
    }
    $CSVServerList = $ServerCSVList | Where-Object { ![string]::IsNullOrWhiteSpace($_.Server) } | ForEach-Object {
        $_.Server.ToUpper()
    }
    
    $MissingServers = Compare-Object -ReferenceObject $CSVServerList -DifferenceObject $LogAnalyticsServers | Where-Object { $_.SideIndicator -eq '<=' }
    write-host $missingservers.count " Servers are missing from Log Analytics" -foregroundcolor red
    write-host 'Exporting Error Log: ' $([IO.Path]::Combine($OutputDir, 'ErrorList.csv'))
    $MissingServers | Select-Object @{Name = 'Server'; Expression = { $_.InputObject } } | Export-Csv -NoTypeInformation -Path $([IO.Path]::Combine($OutputDir, 'ErrorList.csv'))

    return      
}


$summaryReport = @()

foreach ($Item in $ServerCSVList) {
    # Put the servername into a variable to make things slightly more readable.                
    $Server = $Item.Server.ToUpper()
    # This query will pull the LogAnalytics file which has all of the connection data in it for a record.        
    
    $Query = "VMConnection
                | where Protocol == 'tcp' and Computer contains('$Server') and not(SourceIp == '127.0.0.1')  and not(SourceIp == DestinationIp) and not(DestinationIp == '169.254.169.254')
                | summarize
                    Responses = sum(Responses),
                    LinksFailed = sum(LinksFailed),
                    MaxLinksLive = max(LinksLive),
                    TotalBytesSent = sum(BytesSent),
                    TotalBytesReceived = sum(BytesReceived),
                    AverageResponseTime = 1.0 * sum(ResponseTimeSum) / sum(Responses)
                    by
                    Computer,Direction,ProcessName,SourceIp,DestinationIp,DestinationPort,RemoteIp
                | order by
                Direction asc"

    $result = DownloadData $Query
    $ResultCount = $($result.results | select-object Computer | Measure-Object).count
    # If the results returned is > 0, process the output
    if ($ResultCount -gt 0) {  
        $OutputPath = [IO.Path]::Combine($OutputDir, $Server, $Server + "_$($CheckType)_LogAnalytics_$Timespan.csv")
        createdir -newpath $([IO.Path]::Combine($OutputDir, $Server))
        $result.results | export-csv $OutputPath -Delimiter ',' -NoTypeInformation
        Write-Host "Exported $ResultCount results to output file for the server $Server to $OutputPath" -ForegroundColor Yellow
               
        # Run this query to pull the information for the network testing script.
        $Query = "VMConnection
                | where Protocol == 'tcp' and Computer contains('$Server') and not(SourceIp == '127.0.0.1')  and not(SourceIp == DestinationIp) and Direction == 'outbound' and not(DestinationIp == '169.254.169.254')
                | summarize by 
                    SourceIp, ProcessName, RemoteIp, DestinationPort
                |order by RemoteIp"
        $result = DownloadData $Query
        $ResultCount = $($result.results | select-object Computer | Measure-Object).count
        if ($ResultCount -ne 0 -and $CheckType -eq 'Pre') {
            $OutputPath = [IO.Path]::Combine($OutputDir, $Server, $Server + '_PortScanInput.csv')
            createdir -newpath $([IO.Path]::Combine($OutputDir, $Server))
            $result.results | export-csv $OutputPath -Delimiter ',' -NoTypeInformation
            Write-Host "Exported $ResultCount results to output file for the server $Server to $OutputPath" -ForegroundColor Yellow
            
        }
    }
    else {
        Write-Host "$ResultCount results found for server" $Server -ForegroundColor Red
        # If there are 0 results found, log the servername, environment, workload to a new file so we can re-run that file later on.
        $MyObject = New-Object -TypeName PSObject
        $MyObject | Add-Member @{Server = $Server }
        $summaryReport += $MyObject
    }
}
# Export the errors to a new file that can be used as a serverlist.csv input to the script.  this helps when we need to run things again.
write-host 'Exporting Error Log: ' $([IO.Path]::Combine($OutputDir, 'ErrorList.csv'))
$summaryReport | Export-Csv -NoTypeInformation -Path $([IO.Path]::Combine($OutputDir, 'ErrorList.csv'))