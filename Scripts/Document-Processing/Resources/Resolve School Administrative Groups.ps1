. "c:\Axis\resources\IEP504_Functions.ps1"
. "c:\Axis\IEP504_Configuration.ps1" 
cls
# Configuraton ####################################################################################

# Ensure this value contain all of your school abbreviations
$schoolAbbreviations = @("ACE","AHS","ELC","HGS","HMS","NWE");

# End Configuration ###############################################################################

function Get-DocumentBySchool
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$ListName,
        [Parameter(Mandatory = $true)]
        [String]$SchoolAbbreviation
    )  try
    {
        $result = Get-PnPListItem -List $ListName | Where-Object { $_["School0"] -EQ $SchoolAbbreviation; }
        if ($result.count -eq 0) # Either nothing or not unique
        { throw [System.IO.FileNotFoundException] "School $($SchoolAbbreviation) Result count: $($result.count)." }
        return $result;
    }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
}


function Update-AdminGroups
{
    Param (
        [Parameter(Mandatory = $true)]
        [object]$documentgroup,
        [Parameter(Mandatory = $true)]
        [String]$SchoolAbbreviation
    )  try
    {
       #Iterate the documents in the group, and add the special services group and district group
        Update-Status -Message " -- Processing $($documentgroup.Count) Documents for $($SchoolAbbreviation)." -Level Info;
        foreach ($document in $documentgroup)
        {
            Update-Status -Message " ---- Processing document '$($document.FieldValues.FileLeafRef)'." -Level Info; 
            Set-PnPListItemPermission -List $config.DocumentLibraryName -Identity $document.Id -Group $($config.SchoolTeamGroupPrefix+$SchoolAbbreviation) -AddRole 'Contribute'
            Set-PNPListItemPermission -List $config.DocumentLibraryName -Identity $document.Id -Group $($config.SchoolTeamGroupPrefix+"District") -AddRole 'Contribute'
        } 
    }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
}


Update-Status -Message "=========== Running: Admin Group Resolution Routine =================." -Level Info; 


#Connect to SharePoint
Connect-SharePoint -siteUrl $config.SiteUrl -useMFA $config.UseMFA -credentialLocation $config.CredentialLocation

foreach ($abbreviation in $schoolAbbreviations)
{
  Update-Status -Message "Gathering Documents for $($abbreviation)"
  $docArray = Get-DocumentBySchool -ListName $config.DocumentLibraryName -SchoolAbbreviation $abbreviation
  Update-AdminGroups -documentgroup $docArray -SchoolAbbreviation $abbreviation
}

Disconnect-SharePoint;
Update-Status -Message "=========== Finished: Admin Group Resolution Routine =================." -Level Info; 

