
#Requires -Version 5.0
#Requires -Modules SharePointPnPPowerShellOnline

#region 0: Setup
#Initialize Function Library and Configurations
#==============================================
. ".\resources\IEP504_Functions.ps1"
. ".\IEP504_Configuration.ps1" 

Clear-Host;     
Update-Status -Message  "=[ PROCESS STARTED ]======================================="
If ($config.SimulationMode)
{Update-Status -Message "=[ Simulation Mode ]=======================================" -Level Warn}
#endregion

#region 1: Read Export File
#================================================================================================================================
$allIEPStudents = Get-DataFromExport -columns $config.Columns -fileToRead $config.ExportFile -mergeColumn "STUDENT_NUMBER"
#endregion

#region 2: Look for a bookmark file from the last run. If it exists, determine what's different, and hold results for later.
#================================================================================================================================
If (!$config.SimulationMode)
{
    #If Not in Simulation...
    If ($config.BookMarkFile -and !$config.IgnoreBookmark)  #If Bookmark is Specified and We're not Ignoring it...
    { $updatedIEPStudents = Resolve-Bookmark -bookmarkFile $config.BookMarkFile -recentData $allIEPStudents -columns $config.Columns }
    else
    {
        #Bookmark either not specified or set to Ignore. 
        Update-Status -Message "A Bookmark file will not be processed per Configuration settings. All documents will be reviewed." -Level Warn
        $updatedIEPStudents = $allIEPStudents; # Treat the entire list as updates.  
    }
}
else
{
    #We are in simulation. 
    Update-Status -Message "Bookmark file is not processed in simulation mode. All documents will be reviewed." -Level Warn
    $updatedIEPStudents = $allIEPStudents; # Treat the entire list as updates. 
}
#endregion   

#region 2: Connect to SharePoint
#================================================================================================================================
Connect-SharePoint -siteUrl $config.SiteUrl -useMFA $config.UseMFA -credentialLocation $config.CredentialLocation
#endregion

#region 3: Process files Dropbox (if applicable)
#================================================================================================================================
If ($config.DropBoxName)
{ Resolve-Dropbox -dropBoxName $config.DropBoxName -documentLibraryName $config.DocumentLibraryName -studentData $allIEPStudents}
#endregion

#region 4: Process Changes in the Main Document Library. For each student, find all their documents and update each one.
#================================================================================================================================
If ($UpdatedIEPStudents)
{ Resolve-Documents -documentLibraryName $config.DocumentLibraryName -studentData $updatedIEPStudents }
else { Update-Status -Message "There are currently no updates to be made." -Level Warn}
#endregion

#region 5:
Disconnect-SharePoint;
If ($config.SimulationMode)
{Update-Status -Message "=[ End Simulation Mode ]===================================" -Level Warn}
Update-Status -Message  "=[ PROCESS END ]==========================================="
#endregion
