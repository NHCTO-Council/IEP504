#Load the assemblies needed to run the SharePoint Online Client. Visit the following link if assemblies are not available:
#https://www.microsoft.com/en-us/download/details.aspx?id=35588

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

$VerbosePreference = 'Continue'
$ErrorActionPreference = 'Continue'

$config = @{
    #################################### Configuration Section ###################################################
    AuditListTitle     = "IEP 504 Audit Logs";
    IntervalMinutes    = 15;
    LogFilePath        = "./AuditLogProcessing.log";
    NumRetries         = 1;
    Operations         = "FileAccessed,FileDeleted,FilePreviewed";
    RecordType         = "SharePointFileOperation";
    ResultSize         = 1000;
    SimulationMode     = $false;
    SiteUrl            = "https://moreltechnology.sharepoint.com/sites/dev/IEP504v2";
    TenantAdminUrl     = "https://moreltechnology-admin.sharepoint.com";
    TimeToGoBack       = New-TimeSpan -Days 0 -Hours 1 -Minutes 0;
    ValidDocTypes      = @('XLS', 'XLSX', 'DOC', 'DOCX', 'PDF', 'PPT', 'PPTX');
    VerboseMode        = $true;

    #Specify the path where a Service User credential file resides. If it does not exist, 
    #the script will prompt to create one and store it in the location specified

    CredentialLocation = "./securecredential.credential";
    #################################### End Configuration Section ###################################################
}


Function Get-Creds([String]$credentialLocation)
{
    if (Test-Path $credentialLocation)
    { $credential = Import-CliXml -Path $credentialLocation }
    else
    { 
        Get-Credential | Export-CliXml -Path $credentialLocation 
        $credential = Import-CliXml -Path $credentialLocation 
    }
    return $credential
}
Function Get-SPContext([PSCredential]$credential, [String]$adminUrl, [String]$siteUrl)
{ 
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
    $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($credential.UserName, $credential.Password)
    Connect-SPOService -Url $adminUrl -Credential $credential
    return $context
} 
Function Write-LogFile ([String]$Message)
{
    if ($config.VerboseMode) { Write-Verbose -Message $Message }
    $final = [DateTime]::Now.ToString() + ":" + $Message
    $final | Out-File $config.LogFilePath -Append
}
Function Write-AuditList ([Microsoft.SharePoint.Client.ClientContext]$context, [Microsoft.SharePoint.Client.List]$list, [Array]$results)
{
    $filterCount = 0
    $writtenCount = 0
    foreach ($record in $results)
    {

        $auditData = $record.AuditData | ConvertFrom-Json
        if (!($auditData.SourceFileExtension -in $config.ValidDocTypes))
        { $filterCount = $filterCount + 1; }
        else
        {
            $user = $context.Web.EnsureUser($auditData.UserId)
            $context.Load($user)
            $context.ExecuteQuery()

            if (!$config.SimulationMode)
            {
                $ListItemCreationInformation = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
                $NewListItem = $list.AddItem($ListItemCreationInformation)
                $NewListItem["AuditUserId"] = $user;
                $NewListItem["AuditObjectId"] = $auditData.ObjectId;
                $NewListItem["AuditSiteUrl"] = $auditData.SiteUrl;
                $NewListItem["AuditSourceRelativeUrl"] = $auditData.SourceRelativeUrl;
                $NewListItem["AuditSourceFileName"] = $auditData.SourceFileName;
                $NewListItem["AuditOperation"] = $auditData.Operation;
                $NewListItem["AuditDate"] = $auditData.CreationTime;
                $NewListItem.Update();
                $context.ExecuteQuery();
                $writtenCount = $writtenCount + 1;
            }
            else
            {
                Write-LogFile "INFO: (Simulation Mode) Would have created List Item: $($auditData)"
            } 
        }
    }
    Write-LogFile "INFO: Removed $($filterCount) Items which did not match the file extension filter."
    Write-LogFile "INFO: Wrote $($writtenCount) Items to the list."
    Write-LogFile "__________________________________________________________________________________"  
}
$startDate = (Get-Date -date ($(get-date) - $config.TimeToGoBack) -format g)
$endDate = (Get-Date -format g) #(today)
[DateTime]$currentStart = $startDate
[DateTime]$currentEnd = $startDate
$currentTries = 0

$context = Get-SPContext -credential (Get-Creds -credentialLocation $config.CredentialLocation) -adminUrl $config.TenantAdminUrl -siteUrl $config.SiteUrl;
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential (Get-Creds -credentialLocation $config.CredentialLocation) -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber

$list = $context.Web.Lists.GetByTitle($config.AuditListTitle)
$context.Load($list)
$context.ExecuteQuery()

Write-LogFile "INFO: Retrieving Audit Logs between $($startDate) and $($endDate)"
while ($true)
{
    $currentEnd = $currentStart.AddMinutes($config.IntervalMinutes)
    if ($currentEnd -gt $endDate)
    {
        break
    }
    $currentTries = 0
    $sessionID = [DateTime]::Now.ToString().Replace('/', '_')
    Write-LogFile " --- INFO: Collecting a $($config.IntervalMinutes) Minute interval from $($currentStart) to $($currentEnd)."
    $currentCount = 0
    while ($true)
    {
        $context.Load($context.Web); $context.ExecuteQuery();
        [Array]$results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -RecordType $config.RecordType -Operations $config.Operations -ObjectIds $context.Web.ServerRelativeUrl -SessionId $sessionID -SessionCommand ReturnNextPreviewPage -ResultSize $config.ResultSize
        if ($results -eq $null -or $results.Count -eq 0)
        {
            #Retry if needed. This may be due to a temporary network glitch
            if ($currentTries -lt $config.NumRetries)
            {
                $currentTries = $currentTries + 1
                Write-LogFile "     --- WARNING: No Results Returned.  Retrying, ($($currentTries) of $($config.NumRetries))."
                continue
            }
            else
            {
                Write-LogFile "     --- WARNING: Empty data set returned (Retry count reached)."
                break
            }
        }
        $currentTotal = $results[0].ResultCount
        if ($currentTotal -gt 5000)
        {
            Write-LogFile " --- WARNING: $($currentTotal) total records match the search criteria. Some records may get missed. (Consider reducing the time interval.)"
        }
        $currentCount = $currentCount + $results.Count
        $message = "INFO: Successfully retrieved $($currentTotal) records for the current time range."
        Write-LogFile $message
        if ($results.Count -gt 0)
        {
            Write-AuditList -context $context -list $list -results $results
        }
        if ($currentTotal -eq $results[$results.Count - 1].ResultIndex)
        {
            break
        }
    }
    $currentStart = $currentEnd
}
Remove-PSSession $Session
