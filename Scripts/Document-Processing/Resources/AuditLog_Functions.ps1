#Requires -Version 5.0
#Requires -Modules SharePointPnPPowerShellOnline
<#
.SYNOPSIS
  Audit Log Retrieval Function Library
.NOTES
  Version:        4.0.190520
  Author:         Jeremy M. Morel, Axis Business Solutions, Ltd
  Creation Date:  8/31/2018
#>
#region Core Utility Functions =============================================

function Connect-AuditLog
{

    Param (

        [Parameter(Mandatory = $false)]
        [String]$credentialLocation

    )
    Try {Remove-PSSession $AuditLogSession |Out-Null}catch {<#Bury#>}
    try
    {
        if (Test-Path $credentialLocation) { $credential = Import-CliXml -Path $credentialLocation }
        else { Get-Credential | Export-CliXml -Path $credentialLocation ; $credential = Import-CliXml -Path $credentialLocation; }
        $AuditLogSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection 
        Import-PSSession $AuditLogSession -AllowClobber -DisableNameChecking
        Update-Status -Message "Established Audit Log Session." -level Info
    }
    Catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error }
}

function Connect-SharePoint
{

    Param (
        [Parameter(Mandatory = $true)]
        [String]$siteUrl,

        [Parameter(Mandatory = $false)]
        [String]$credentialLocation,

        [Parameter(Mandatory = $false)]
        [Boolean]$useMFA = $false
    )
    Try {Disconnect-PnPOnline |Out-Null}catch {<#Bury#>}
    try
    {
        If (!$useMFA)
        {
            if (Test-Path $credentialLocation) { $credential = Import-CliXml -Path $credentialLocation }
            else { Get-Credential | Export-CliXml -Path $credentialLocation ; $credential = Import-CliXml -Path $credentialLocation; }
            Connect-PnPOnline –Url $siteUrl –Credentials $credential
        }
        else
        { Connect-PnPOnline –Url $siteUrl –UseWebLogin }
        Update-Status -Message "Conncted to SharePoint Site: $((Get-PnPConnection).Url)" -level Info
    }
    Catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error }
}

function Disconnect-AuditLog
{ 
    Try { Remove-PSSession $AuditLogSession |Out-Null}catch {<#Bury#>}
    
}

function Disconnect-SharePoint
{
    Try {Disconnect-PnPOnline |Out-Null}catch {<#Bury#>}
}

function Get-LogData
{
    Param (
        [Parameter(Mandatory = $true)]
        [System.TimeSpan]$timeToGoBack = $config.AuditTimeToGoBack,

        [Parameter(Mandatory = $true)]
        [Int]$intervalMinutes = $config.AuditIntervalMinutes,

        [Parameter(Mandatory = $true)]
        [String[]]$operations = $config.AuditOperations,

        [Parameter(Mandatory = $true)]
        [String]$recordType = $config.RecordType,

        [Parameter(Mandatory = $true)]
        [Int]$resultSize = $config.AuditResultChunkSize,

        [Parameter(Mandatory = $true)]
        [Int]$retries = $config.AuditRetries,

        [Parameter(Mandatory = $true)]
        [String]$auditListName = $config.AuditListName


    )

$startDate = (Get-Date -date ($(get-date) - $timeToGoBack) -format g)
$endDate = (Get-Date -format g) #(today)
[DateTime]$currentStart = $startDate
[DateTime]$currentEnd = $startDate
$currentTries = 0

Update-Status -Message "Retrieving Audit Logs between $($startDate) and $($endDate)" -Level Info
while ($true) {
    $currentEnd = $currentStart.AddMinutes($intervalMinutes)
    if ($currentEnd -gt $endDate) {
        break
    }
    $currentTries = 0
    $sessionID = [DateTime]::Now.ToString().Replace('/', '_')
    Update-Status -Message " --- Collecting a $($intervalMinutes) Minute interval from $($currentStart) to $($currentEnd)." -Level Info
    $currentCount = 0
    while ($true) {
  
        [Array]$results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -RecordType $recordType -Operations $operations -ObjectIds $((Get-PnPWeb).ServerRelativeUrl) -SessionId $sessionID -SessionCommand ReturnNextPreviewPage -ResultSize $resultSize
        if ($results -eq $null -or $results.Count -eq 0) {
            #Retry if needed. This may be due to a temporary network glitch
            if ($currentTries -lt $retries) {
                $currentTries = $currentTries + 1
                Update-Status -Message "     --- No Results Returned.  Retrying, $($currentTries) of $($retries)." -Level Warn
                continue
            }
            else {
                Update-Status -Message "     --- Empty data set returned (Retry count reached)." -Level Warn
                break
            }
        }
        $currentTotal = $results[0].ResultCount
        if ($currentTotal -gt 5000) {
            Update-Status -Message " --- $($currentTotal) total records match the search criteria. Some records may get missed. (Consider reducing the time interval.)" -Level Warn
        }
        $currentCount = $currentCount + $results.Count
        Update-Status -Message "Successfully retrieved $($currentTotal) records for the current time range." -Level Info
        if ($results.Count -gt 0) {
           Write-AuditList -auditListName $auditListName -results $results
        }
        if ($currentTotal -eq $results[$results.Count - 1].ResultIndex) {
            break
        }
    }
    $currentStart = $currentEnd
}
}

function Get-SharePointUser
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$userName
    )
    Try
    {
        $safeUser = (Get-PNPUser | Where-Object LoginName -match $userName).Email ;
        $user = (Get-PNPWeb).EnsureUser($safeUser);
        $ctx = Get-PnPContext;
        $ctx.Load($user);
        Invoke-PnPQuery;
        return $user;
    }
    Catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error }
}

function Update-Status
{
    Param (
        [Parameter(Mandatory = $true)]
        [Object]$Message,

        [Parameter(Mandatory = $false)]
        [string]$Path = $config.AuditOutputLog,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Error", "Warn", "Info", "Data")]
        [string]$Level = "Info"
    )
    if (!$Path) { $path = "$($pwd)\output.log" }
    # Format Date for our Log File
    $FormattedDate = Get-Date -Format g

    # Write message to error, warning, or verbose pipeline and specify $LevelText
    switch ($Level)
    {
        'Error'
        {
            $txt = "[ERROR]: $($Message)"
            Write-Host  $txt -ForegroundColor Red
            "$($FormattedDate) $($txt)" | Out-File $Path -Append

        }
        'Warn'
        {
            $txt = "[WARN]: $($Message)"
            Write-Host  $txt -ForegroundColor Yellow
            "$($FormattedDate) $($txt)" | Out-File $Path -Append
        }
        'Info'
        {
            $txt = "[INFO]: $($Message)"
            Write-Host  $txt -ForegroundColor Cyan
            "$($FormattedDate) $($txt)" | Out-File $Path -Append
        }
        'Data'
        {
            Write-Host ($Message|Format-Table|Out-String) -ForegroundColor Gray
            $("`t`t`t`t") + ($Message|Format-Table|Out-String) | Out-File $Path -Append
        }
    }
}

Function Write-AuditList
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$auditListName,

        [Parameter(Mandatory = $true)]
        [Array]$results
        )

    $filterCount = 0
    $writtenCount = 0
    foreach ($record in $results) {

        $auditData = $record.AuditData | ConvertFrom-Json
        if (!($auditData.SourceFileExtension -in $config.AuditDocTypes))
        { $filterCount = $filterCount + 1; }
        else {
                if (!$config.SimulationMode) {
                Add-PnPListItem -List $auditListName -Values @{ "AuditUserId" = $auditData.UserId;  
                                                                "AuditObjectId" = $auditData.ObjectId;
                                                                "AuditSiteUrl" = $auditData.SiteUrl;
                                                                "AuditSourceRelativeUrl" = $auditData.SourceRelativeUrl;
                                                                "AuditSourceFileName" = $auditData.SourceFileName;
                                                                "AuditOperation" = $auditData.Operation;
                                                                "AuditDate" = $auditData.CreationTime; }
                $writtenCount = $writtenCount + 1;
             }
            else {
                Update-Status -Message "(Simulation Mode) Would have created List Item: $($auditData)" -Level Warn
            } 
        }
    }
    Update-Status -Message "Removed $($filterCount) Items which did not match the file extension filter." -Level Info
    Update-Status -Message "Wrote $($writtenCount) Items to the list." -Level Info
    Update-Status -Message "__________________________________________________________________________________"  
}

#endregion Core Utility Functions ==========================================