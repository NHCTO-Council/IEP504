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
}