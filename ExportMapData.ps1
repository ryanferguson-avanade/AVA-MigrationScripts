
<#
.SYNOPSIS
    This script will extract data from a local MAP Toolkit database and write it into a series of CSV files.  
    Reqires Powershell V5 at a minimum
    Version: 2020.06.01
.DESCRIPTION
    Usage: .\ExportMapData.ps1 -SQLServer  "(LocalDB)\MAPToolkit" -SQLDatabase "MyMAPDB" -OutputDir "c:\temp" -AssessmentName "FirstExport"
     -SQLServer            Name of SQL Server to connect to.  "(LocalDB)\MAPToolkit" is what is needed for a local install of MAP.
     -SQLDatabase          Name of the database used for the MAP Toolkit scan.
     -OutputDir            Directory to run the script from.  This is where the output will be put. The default will be to use the directory the script is run from.
     -AssessmentName       This is a unique name for this MAP extract.  It will default to "Export-YYYYMMDD"
     -InstallSQLModules    This will install SQL Powershell modules to allow for the export to run.
     -SQLModuleZipFile     Location of the zip file to use to extract the Powershell modules.  
     -Help                 These usage instructions are displayed.
.EXAMPLE
    .\ExportMapData.ps1 -SQLServer  "(LocalDB)\MAPToolkit" -SQLDatabase "MyMAPDB" -OutputDir "c:\temp" -AssessmentName "FirstExport"
    Export all MAP Toolkit data from MyMAPDB on the local MAP installation into c:\temp and name the files FirstExport-
.NOTES
    Avanade - June 2020
    Author: Ryan Ferguson
#>

#-----VARIABLES TO MODIFY
param (
    [string]$SQLDatabase, #This is the name of the DB.  If using MAP, you can see this by going to 'manage databases' within the interface,
    [string]$SQLServer = "(LocalDB)\MAPToolkit", #This is the local instance of the MAP Toolkit.  If you are using this against a restored DB in SQL, this will change to your SQL server instance
    [string]$OutputDir, #The default will be to use the directory the script is run from.
    [string]$AssessmentName = "Export-$(get-date -Format yyyyMMdd)",  #The name of the assessment.  If there are multiple MAP Toolkit extracts, this name needs to be unique.
    [switch]$InstallSQLModules = $false,
    [string]$SQLModuleZipFile  = ".\sqlserver.zip"
)

function CreateDir {
  param(
    [string]$NewPath  
  )
  If (!(test-path $NewPath)){
      write-host "Creating $NewPath"
      New-Item -ItemType directory -force -Path $NewPath|Out-Null
  } 
}

#-------Start of script.  Do not modify past this point.

if (!$OutputDir) {
  $OutputDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
}

if ($Help -or ($PSBoundParameters.count -eq 0)) {
  Get-Help $MyInvocation.MyCommand.Definition
  return
}

if ($InstallSQLModules){
  Write-Host "Installing SQL Modules"
  Expand-Archive $SQLModuleZipFile -DestinationPath 'C:\Program Files\WindowsPowerShell\Modules'
  return
}

#Create Directories
$ServerDir = join-path -path $OutputDir -childpath "Servers"
$SoftwareDir = join-path -path $OutputDir -childpath "Software"
$SQLDir = join-path -path $OutputDir -childpath "SQL"
$IPDir = join-path -path $OutputDir -childpath "IPAddr"
$IISDir = join-path -path $OutputDir -childpath "IIS"
$ServicesDir = join-path -path $OutputDir -childpath "Services"
$FeaturesDir = join-path -path $OutputDir -childpath "Features"
$DisksDir = join-path -path $OutputDir -childpath "Disks"

$ServerFile = join-path -path $ServerDir -childpath "$AssessmentName-Servers.csv"
$SoftwareFile = join-path -path $SoftwareDir -childpath "$AssessmentName-Software.csv"
$SQLFile = join-path -path $SQLDir -childpath "$AssessmentName-SQL.csv"
$IPFile = join-path -path $IPDir -childpath "$AssessmentName-IPAddr.csv"
$IISFile = join-path -path $IISDir -childpath "$AssessmentName-IIS.csv"
$ServicesFile = join-path -path $ServicesDir -childpath "$AssessmentName-Services.csv"
$FeaturesFile = join-path -path $FeaturesDir -childpath "$AssessmentName-Features.csv"
$DisksFile = join-path -path $DisksDir -childpath "$AssessmentName-Disks.csv"




CreateDir -NewPath $ServerDir
CreateDir -NewPath $SoftwareDir
CreateDir -NewPath $SQLDir
CreateDir -NewPath $IPDir
CreateDir -NewPath $IISDir
CreateDir -NewPath $ServicesDir
CreateDir -NewPath $FeaturesDir
CreateDir -newPath $DisksDir


#Get the servers
$ServerQuery = "
select Devices.DeviceNumber,AdDnsHostName,AdDomainName,AdFullyQualifiedDomainName,AdOsVersion,BuildNumber,ComputerSystemName,CreationDate,CsdVersion,CurrentLanguage,DistinguishedName,DnsHostName,EnclosureManufacturer,EnclosureSerialNumber,FreePhysicalMemory,FreeSpaceInPagingFiles,FreeVirtualMemory,HostNameForVm,InventoryRowversion,InventoryWatermark,LastBootupTime,Locale,Model,MuiLanguages,NetworkServerModeEnabled,Devices.NumberOfCores,NumberOfLicensedUsers,Devices.NumberOfLogicalProcessors,Devices.NumberOfProcessors,NumberOfUsers,OperatingSystem,OperatingSystemServicePack,OperatingSystemSku,Organization,OsArchitecture,OsBuildType,OsCaption,OsInstallDate,OsManufacturer,PcSystemType,PowerOnPasswordStatus,PowerState,PowerSupplyState,Roles,SiteName,SystemType,TotalPhysicalMemory,TotalVirtualMemorySize,TotalVisibleMemorySize,VmFriendlyName,WmiDnsHostName,WmiDomainName,WmiOsVersion,WMIStatus,IPAddress,MACAddress,DNSServer,SubnetMask,IPGateway,WINSServer,[Domain/Workgroup],MachineType,
 convert(int,round([TotalPhysicalMemory]/1024/1024/1024.0,0)) as 'Memory (GB)'
	   ,T2.Interfaces as NICs
	   -- Ensure you take away the OS drive from the Total Drive
     ,t5.OSDriveSize as OSDISK
     ,convert(int,round(t3.TotalDriveSize/1024/1024/1024.0,0)) as 'Total HD Size (GB)'
	   ,convert(int,round((t3.TotalDriveSize - t3.TotalFreeSpace)/1024/1024/1024.0,0)) as 'Total HD Used (GB)'
	   ,convert(int,round(t3.TotalFreeSpace/1024/1024/1024.0,0)) 'Total Free Space (GB)'
	   	   
  FROM [Core_Inventory].[Devices]
left join AllDevices_Assessment.HardwareInventoryEx on devices.DeviceNumber = AllDevices_Assessment.HardwareInventoryEx.DeviceNumber
  left join AzureMigration_Reporting.AzureSizingView  on devices.DeviceNumber = AzureSizingView.DeviceNumber
  
  left join (
  SELECT [DeviceNumber], count(*) as Interfaces      
  FROM [Win_Inventory].[NetworkAdapters]
  where NetConnectionId IS NOT NULL
  group by DeviceNumber) as T2
  on devices.DeviceNumber = T2.DeviceNumber

      left join (
  SELECT [DeviceNumber], count(*) as LogicalDriveCount, sum(Size) as TotalDriveSize, sum(FreeSpace) as TotalFreeSpace
  FROM [Win_Inventory].[LogicalDisks]
  group by [DeviceNumber]) as T3
  on devices.DeviceNumber = T3.DeviceNumber

    left join (
  SELECT [DeviceNumber], count(*) as PhysicalDriveCount
  FROM [Win_Inventory].[DiskDrives]
  group by [DeviceNumber]) as T4
  on devices.DeviceNumber = T4.DeviceNumber
  
   left join (
  SELECT [DeviceNumber],   convert(int,round(Size/1024.0/1024/1024.0,0)) as OSDriveSize, FreeSpace as OSDriveFreeSpace
  FROM [Win_Inventory].[LogicalDisks]
  where Name = 'C:'
	) as T5
  on devices.DeviceNumber = T5.DeviceNumber
  
   left join (
	 select Devicenumber, 
	 MetricMax as 'CPUMetrixMax' from Perf_Assessment.MetricAggregation
	left join Perf_Inventory.MetricTypes on MetricTypes.MetricType = Perf_Assessment.MetricAggregation.MetricType
	where MetricName = 'cpu_percentage'
	) as T6
  on devices.DeviceNumber = T6.DeviceNumber

     left join (
	 select Devicenumber, 
	 MetricMax as 'IOPSMax' from Perf_Assessment.MetricAggregation
	left join Perf_Inventory.MetricTypes on MetricTypes.MetricType = Perf_Assessment.MetricAggregation.MetricType
	where MetricName = 'diskiops_total'
	) as T7
  on devices.DeviceNumber = T7.DeviceNumber

       left join (
	 select Devicenumber, 
	 Metricavg/1024/1024 as 'DiskTP' from Perf_Assessment.MetricAggregation
	left join Perf_Inventory.MetricTypes on MetricTypes.MetricType = Perf_Assessment.MetricAggregation.MetricType
	where MetricName = 'disk_bytes_sec'
	) as T8
  on devices.DeviceNumber = T8.DeviceNumber

"
write-host "Collecting Server Data..."
Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $ServerQuery | Export-Csv -Path $ServerFile -Force -NoTypeInformation
Import-Csv $ServerFile | Measure-Object | select-object count |Format-List

#Get the software
$SoftwareQuery = "
SELECT 	distinct devices.ComputerSystemName as 'VM Name',
	 [Name]
      ,[Vendor],version,products.Description,InstallDate,InstallLocation	  
  FROM [Win_Inventory].[Products]
  left join [Core_Inventory].[Devices]  on devices.DeviceNumber = [Win_Inventory].[Products].DeviceNumber
  "
write-host "Collecting Software Data..."
Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $SoftwareQuery | Export-Csv -Path $SoftwareFile -Force -NoTypeInformation
Import-Csv $SoftwareFile | Measure-Object | select-object count |Format-List

#Get IIS Details
$IISQuery = "
SELECT ComputerSystemName
      ,[ApplicationPool]
      ,[ManagedRuntimeVersion]
      ,[IISVirtualDirApplications].[Name]
	  ,[AppRoot]
      ,[Path]
      ,[AppPoolId]
      ,[HasAspx]
      ,[HasAsp]
      ,[HasPhp]
      ,[HasHtml]
      ,[HasJava]
      ,[HasRuby]
	  ,ServerComment as 'Website'
  FROM [AzureMigration_Inventory].[IISApplicationPools]
  left join [AzureMigration_Inventory].[IISVirtualDirApplications] on [IISVirtualDirApplications].DeviceNumber = [IISApplicationPools].DeviceNumber and IISApplicationPools.ApplicationPool = IISVirtualDirApplications.AppPoolId
  left join [AzureMigration_Inventory].[IISWebServerSetting] on IISWebServerSetting.DeviceNumber = [IISVirtualDirApplications].DeviceNumber and IISWebServerSetting.Name = left([IISVirtualDirApplications].Name,7)
  left join  [Core_Inventory].[Devices] on devices.DeviceNumber = [IISApplicationPools].DeviceNumber
"
write-host "Collecting IIS Data..."
Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $IISQuery | Export-Csv -Path $IISFile -Force -NoTypeInformation
Import-Csv $IISFile | Measure-Object| select-object count |format-list

#Get IPAddr Details
$IPAddrQuery = "
SELECT [IpAddress],computersystemname

  FROM [Win_Inventory].[NetworkAdapterConfigurations]
  left join Core_Inventory.Devices on devices.DeviceNumber = NetworkAdapterConfigurations.DeviceNumber
"
write-host "Collecting IP Data..."
Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $IPAddrQuery | Export-Csv -Path $IPFile -Force -NoTypeInformation
Import-Csv $IPFile | Measure-Object| select-object count |format-list

#Get SQL Details
$SQLQuery = "
select distinct ComputerSystemName,Servicename,InstanceName,Sqlservicetype,VersionCoalesce,Skuname,Splevel,[Clustered],instanceid,SqlServer_Assessment.SqlInstances.SqlConnected
      ,[DatabaseName]
      ,[DataFilesSizeKB]
      ,[LogFilesSizeKB]
      ,[LogFilesUsedSizeKB]
      ,[PercentLogUsed]
      ,[Size]
      ,[CompatibilityLevel]
      ,[Status]
      ,[Owner]
      ,[CreatedTimestamp]
      ,[LastBackupDate]
      ,[NumberTables]
      ,[NumberViews]
      ,[NumberSp]
      ,[NumberFunction]

from SqlServer_Assessment.SqlInstances left join  [Core_Inventory].[Devices] on devices.DeviceNumber = SqlServer_Assessment.SqlInstances.DeviceNumber
left join [SqlServer_Reporting].[SqlDatabasesView] on [SqlServer_Reporting].[SqlDatabasesView].[DeviceNumber] = SqlServer_Assessment.SqlInstances.DeviceNumber
"
write-host "Collecting SQL Data..."
Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $SQLQuery | Export-Csv -Path $SQLFile -Force -NoTypeInformation
Import-Csv $SQLFile | Measure-Object| select-object count |Format-List

#Get Service Details
$SQLQuery = "
SELECT distinct devices.ComputerSystemName
      ,[Name]
	  ,[DisplayName]
      ,[Caption]
      ,[Services].[CreateCollectorId]
      ,[Services].[CreateDatetime]
      ,[Services].[Description]
      ,[DesktopInteract]
      ,[PathName]
      ,[ProcessId]
      ,[ServiceType]
      ,[StartMode]
      ,[StartName]
      ,[Started]
      ,[State]
      ,[Status]
      ,[TagId]
      ,[ExecutablePath]      
  FROM [Win_Inventory].[Services]
  left join [Core_Inventory].[Devices]  on devices.DeviceNumber = [Win_Inventory].[Services].DeviceNumber
"
write-host "Collecting Services Data..."
Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $SQLQuery | Export-Csv -Path $ServicesFile -Force -NoTypeInformation
Import-Csv $ServicesFile | Measure-Object| select-object count |Format-List


#Get Features Details
$SQLQuery = "
SELECT devices.ComputerSystemName
      ,[Id]
      ,[Name]
      ,[ParentId]
  FROM [WinServer_Inventory].[ServerFeatures]
    left join [Core_Inventory].[Devices]  on devices.DeviceNumber = [WinServer_Inventory].[ServerFeatures].DeviceNumber
"
write-host "Collecting Features Data..."
Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $SQLQuery | Export-Csv -Path $FeaturesFile -Force -NoTypeInformation
Import-Csv $FeaturesFile | Measure-Object| select-object count |Format-List




#Get Disk Details
$SQLQuery = "
SELECT devices.ComputerSystemName as 'VM Name', [LogicalDisks].Caption, [LogicalDisks].Compressed, [LogicalDisks].CreateDatetime, [LogicalDisks].Description, [LogicalDisks].DeviceId, 
[LogicalDisks].DriveType, [LogicalDisks].FileSystem, convert(int,round([LogicalDisks].FreeSpace/1024.0/1024/1024.0,0))  as 'Free Space (GB)', [LogicalDisks].Name as 'Drive Letter',convert(int,round([LogicalDisks].Size/1024.0/1024/1024.0,0))  as 'Size (GB)',[LogicalDisks].VolumeName, [LogicalDisks].VolumeSerialNumber
 FROM [Win_Inventory].[LogicalDisks]
 left join [Core_Inventory].[Devices]  on devices.DeviceNumber = [Win_Inventory].[LogicalDisks].DeviceNumber
"
write-host "Collecting Disk Data..."
Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $SQLQuery | Export-Csv -Path $DisksFile -Force -NoTypeInformation
Import-Csv $DisksFile | Measure-Object| select-object count |Format-List
