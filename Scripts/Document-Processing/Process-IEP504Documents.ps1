param ([string[]]$studentIds = "",[string[]]$teacherEmails = "")
#Requires -Version 5.0
#Requires -Modules SharePointPnPPowerShellOnline
<#
.SYNOPSIS
  IEP/504 Document Processing Script
.DESCRIPTION
  Requires configuration variables to be specified in IEP_504Configuration.ps1 file.
  Proces makes use of exported data from student information system for importing documents into a 
  master document library; updates document metadata with student information and secures documents to 
  the appropriate case managers, faculty and staff responsible for document stewardship. 
.INPUTS
  None.
.OUTPUTS
  Log file as defined in IEP_504Configuration.ps1
.NOTES
  Version:        4.0.190906
  Author:         Jeremy M. Morel, Axis Business Solutions, Ltd
  Creation Date:  8/31/2018

  Change log:
  09/06/2019 - Added parameters and logic to support limiting process scope to certain students and/or teachers.
#>

#region 0: Setup
#Initialize Function Library and Configurations
#==============================================
. ".\resources\IEP504_Functions.ps1"
. ".\IEP504_Configuration.ps1" 

Clear-Host;     
Update-Status -Message  "=[ PROCESS STARTED ]=======================================" -level Info
If ($config.SimulationMode)
{Update-Status -Message "=[ Simulation Mode ]=======================================" -Level Warn}
#endregion

#region 1: Read Export File
#================================================================================================================================
$allIEPStudents = Get-DataFromExport -columns $config.Columns -fileToRead $config.ExportFile -mergeColumn "STUDENT_NUMBER"
#endregion

#region 2: Filter the export using optional parameters for StudentIDs and/or TeacherEmails
#================================================================================================================================
if($studentIds -or $teacherEmails) 
{ 
   $config.IgnoreBookmark = $true; Update-Status -Message "A Bookmark file will not be processed when selecting specific students or teachers." -Level Warn 
   $config.DropBoxName = ""; Update-Status -Message "The dropbox is not processed when selecting specific students or teachers." -Level Warn  
}

if($studentIds)
{ 
   Update-Status -Message "Per configuration settings, the process scope will be limited to $($studentIds.Count) students." -Level Warn
   $allIEPStudents = $allIEPStudents | ? {$_.STUDENT_NUMBER -in $studentIds} 
}
if($teacherEmails)
{
  Update-Status -Message "Per configuration settings, the process scope will be limited to $($teacherEmails.Count) Teachers." -Level Warn
  $allIEPStudents = $allIEPStudents | ? {$_.Teacher_Email -match $teacherEmails}
}
#endregion

#region 3: Look for a bookmark file from the last run. If it exists, determine what's different, and hold results for later.
#================================================================================================================================
If (!$config.SimulationMode)
{
    #If Not in Simulation...
    If ($config.BookMarkFile -and !$config.IgnoreBookmark)  #If Bookmark is Specified and We're not Ignoring it...
    { $updatedIEPStudents = Resolve-Bookmark -bookmarkFile $config.BookMarkFile -recentData $allIEPStudents -columns ($allIEPStudents | gm -MemberType Property).Name }
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

#region 4: Connect to SharePoint
#================================================================================================================================
Connect-SharePoint -siteUrl $config.SiteUrl -useMFA $config.UseMFA -credentialLocation $config.CredentialLocation
#endregion

#region 5: Process files Dropbox (if applicable)
#================================================================================================================================
If ($config.DropBoxName)
{ Resolve-Dropbox -dropBoxName $config.DropBoxName -documentLibraryName $config.DocumentLibraryName -studentData $allIEPStudents}
#endregion

#region 6: Process Changes in the Main Document Library. For each student, find all their documents and update each one.
#================================================================================================================================
If ($UpdatedIEPStudents)
{ Resolve-Documents -documentLibraryName $config.DocumentLibraryName -studentData $updatedIEPStudents }
else { Update-Status -Message "There are currently no updates to be made." -Level Warn}
#endregion

#region 7: Cleanup and End Process
#================================================================================================================================
Disconnect-SharePoint;
If ($config.SimulationMode)
{Update-Status -Message "=[ End Simulation Mode ]===================================" -Level Warn}
Update-Status -Message  "=[ PROCESS END ]===========================================" -level Info
#endregion
