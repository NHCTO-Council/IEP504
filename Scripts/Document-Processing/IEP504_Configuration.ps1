<<<<<<< HEAD
<#
.SYNOPSIS
  IEP/504 Document Processing and Audit Log Script Configuration 
.NOTES
  Version:        4.0.190828
  Author:         Jeremy M. Morel, Axis Business Solutions, Ltd
  Creation Date:  8/31/2018

  Change log:
  07/31/2019 - Added field 'unknownUserId' to Document Processing Configuration Section.
#>  
$scriptPath = "C:\Axis"
$config =
@{
   #region Document Processing Configuration Section ##########################################################################
    # Student Export File
    ExportFile                 =  $scriptPath + "\Exports\ps-export.txt";

    #Output Log File
    OutputLog                  = $scriptPath + "\Logs\"+$(get-date -UFormat %Y-%m-%d) + "_Output.log";

    # Number of Days Prior to Class Start that a Teacher should be allowed access to student's documents
    GrantAccessDaysBeforeStart = 30;

    # Number of Days After class ends that a Teacher should retain access to student's documents
    GrantAccessDaysAfterEnd    = 0;

    # List the relevant columns from your export file
    Columns                    = @("DATEENROLLED", "DATELEFT", "STUDENT_NUMBER", "STUDENT_WEB_ID", "FIRST_NAME", "LAST_NAME", "SCHOOL_ABBREVIATION", "CLASSOF", "GRADE_LEVEL", "HOME_ROOM", "TEAM", "CASE_MANAGER", "TEACHER_EMAIL", "IEP504");

    # Specify DropBox name in quotes. (i.e.: DropBoxName = "DropBox") Set Value to $null to disable use of the Dropbox).
    DropBoxName                = "IEP 504 Dropbox";

    # Document Library Name
    DocumentLibraryName        = "IEP 504 Documents";

    # Force Check-in of checked out document
    ForceCheckIn               = $true;

    # Use Multi-Factor Auth.
    UseMFA                     = $false;

    # Stored Credential Location
    CredentialLocation         = $scriptPath + "\securecredential.credential";

    # SharePoint Site Url
    SiteUrl                    = "https://YOURTENANT.sharepoint.com/Sites/IEP-504";

    # Bookmark File (Set to $null to disable use of a Bookmark File.)
    BookMarkFile               = "BookMark.txt";

    # Ignore Bookmark file (default is $false. Set to $true to Force processing of all records)
    IgnoreBookmark             = $false;

    # Simulation Mode (default is $false. Set to $true to simulate a run event.)
    SimulationMode             = $false;
   
    # Perform File rename based on type of document.  All scanned files will have name changed from current value to a value of <studentID>_<document_type>.ext (i.e. 0123456_IEP.pdf)
    RenameFile                 = $true;

    #School Team SharePoint Group Prefix.  Pre-existing SharePoint Groups used to grant access to school teams.  
    # Groups should be named with the specified prefix, followed by school abbreviation as contained in the export.
    # (i.e. if prefix is "Student Services - ",and a school abbreviation is "HMS" the sharepoint group name should be "Student Services - HMS")
    SchoolTeamGroupPrefix      ="School Team - ";
  
    # Unknown User Assignment
    # When the process is unable to locate a particular user it encounters a user field which is blank, this user will be added to avoid process crashing.
    unknownUserId = "IEP504UnknownTeacher@YOURDOMAIN.COM";

   #endregion Document Processing Configuration Section #######################################################################

   #region Audit Log Configuration Section ####################################################################################
    AuditListTitle     = "IEP 504 Audit Logs";
    AuditIntervalMinutes    = 15;
    AuditOutputLog        =  $scriptPath + "\Logs\"+$(get-date -UFormat %Y-%m-%d) + "_AuditLogProcessing.log";
    AuditRetries         = 1;
    AuditOperations         = "FileAccessed,FileDeleted,FilePreviewed";
    AuditRecordType         = "SharePointFileOperation";
    AuditResultChunkSize         = 1000;
    AuditTimeToGoBack       = New-TimeSpan -Days 0 -Hours 2 -Minutes 0;
    AuditDocTypes      = @('XLS', 'XLSX', 'DOC', 'DOCX', 'PDF', 'PPT', 'PPTX');
   #endregion Audit Log Configuration Section #################################################################################

=======
$scriptPath = "C:\Users\jmorel\Source Code\IEP-504\Current\Scripts\Document-Processing"
$config =
@{
    # Student Export File
    ExportFile                 = $scriptPath + "\SampleExport.txt";

    #Output Log File
    OutputLog                  = $scriptPath + "\" + (get-date -UFormat %Y-%m-%d) + "_Output.log";

    # Number of Days Prior to Class Start that a Teacher should be allowed access to student's documents
    GrantAccessDaysBeforeStart = 30;

    # Number of Days After class ends that a Teacher should retain access to student's documents
    GrantAccessDaysAfterEnd    = 0;

    # List the relevant columns from your export file
    Columns                    = @("DATEENROLLED", "DATELEFT", "STUDENT_NUMBER", "STUDENT_WEB_ID", "FIRST_NAME", "LAST_NAME", "SCHOOL_ABBREVIATION", "CLASSOF", "GRADE_LEVEL", "HOME_ROOM", "TEAM", "CASE_MANAGER", "TEACHER_EMAIL");

    # Specify DropBox name in quotes. (i.e.: DropBoxName = "DropBox") Set Value to $null to disable use of the Dropbox).
    DropBoxName                = "IEP 504 DropBox";

    # Document Library Name
    DocumentLibraryName        = "IEP 504 Documents";

    # Force Check-in of checked out document
    ForceCheckIn               = $true;

    # Use Multi-Factor Auth.
    UseMFA                     = $false;

    # Stored Credential Location
    CredentialLocation         = $scriptPath + "\securecredential.credential";

    # SharePoint Site Url
    SiteUrl                    = "https://axisbusiness.sharepoint.com/sites/Morel/IEP504v2";

    # Bookmark File (Set to $null to disable use of a Bookmark File.)
    BookMarkFile               = "BookMark.txt";

    # Ignore Bookmark file (default is $false. Set to $true to Force processing of all records)
    IgnoreBookmark             = $false;

    # Simulation Mode (default is $false. Set to $true to simulate a run event.)
    SimulationMode             = $true;
    #
>>>>>>> 77cac6119b1b583aaef6a5803e5f77a558c98e3f
}