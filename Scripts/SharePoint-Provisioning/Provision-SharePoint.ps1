#Requires -Version 5.0
#Requires -Modules SharePointPnPPowerShellOnline
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$UI = New-Object system.Windows.Forms.Form
$UI.ClientSize = '600,325'
$UI.text = "Educational Plan Document Management Framework"
$UI.TopMost = $false

$lblTenant = New-Object system.Windows.Forms.Label
$lblTenant.text = "Tenant Name:"
$lblTenant.AutoSize = $true
$lblTenant.width = 25
$lblTenant.height = 10
$lblTenant.location = New-Object System.Drawing.Point(26, 25)
$lblTenant.Font = 'Microsoft Sans Serif,10'

$lblSiteTitle = New-Object system.Windows.Forms.Label
$lblSiteTitle.text = "Site Title:"
$lblSiteTitle.AutoSize = $true
$lblSiteTitle.width = 25
$lblSiteTitle.height = 10
$lblSiteTitle.location = New-Object System.Drawing.Point(26, 27)
$lblSiteTitle.Font = 'Microsoft Sans Serif,10'

$lblTargetSite = New-Object system.Windows.Forms.Label
$lblTargetSite.text = "Target Site:"
$lblTargetSite.AutoSize = $true
$lblTargetSite.width = 2
$lblTargetSite.height = 10
$lblTargetSite.location = New-Object System.Drawing.Point(26, 57)
$lblTargetSite.Font = 'Microsoft Sans Serif,10'

$lblOwner = New-Object system.Windows.Forms.Label
$lblOwner.text = "Owner Email:"
$lblOwner.AutoSize = $true
$lblOwner.width = 2
$lblOwner.height = 10
$lblOwner.location = New-Object System.Drawing.Point(26, 87)
$lblOwner.Font = 'Microsoft Sans Serif,10'

$txtTenantName = New-Object system.Windows.Forms.TextBox
$txtTenantName.multiline = $false
$txtTenantName.width = 255
$txtTenantName.height = 20
$txtTenantName.TextAlign = "Right"
$txtTenantName.location = New-Object System.Drawing.Point(145, 21)
$txtTenantName.Font = 'Microsoft Sans Serif,10'

$lblTenantSampleText = New-Object system.Windows.Forms.Label
$lblTenantSampleText.text = ".sharepoint.com"
$lblTenantSampleText.AutoSize = $true
$lblTenantSampleText.width = 25
$lblTenantSampleText.height = 10
$lblTenantSampleText.location = New-Object System.Drawing.Point(402, 25)
$lblTenantSampleText.Font = 'Microsoft Sans Serif,10'
$lblTenantSampleText.ForeColor = "#696969"#9b9b9b

$txtSiteTitle = New-Object system.Windows.Forms.TextBox
$txtSiteTitle.multiline = $false
$txtSiteTitle.text = "Educational Plan Document Center"
$txtSiteTitle.width = 255
$txtSiteTitle.height = 20
$txtSiteTitle.location = New-Object System.Drawing.Point(145, 25)
$txtSiteTitle.Font = 'Microsoft Sans Serif,10'

$grpTenant = New-Object system.Windows.Forms.Groupbox
$grpTenant.height = 87
$grpTenant.width = 580
$grpTenant.text = "Tenant"
$grpTenant.location = New-Object System.Drawing.Point(8, 17)

$lblUseMFA = New-Object system.Windows.Forms.Label
$lblUseMFA.text = "Tenant requires Multi-Factor Authentication"
$lblUseMFA.AutoSize = $true
$lblUseMFA.width = 25
$lblUseMFA.height = 10
$lblUseMFA.location = New-Object System.Drawing.Point(26, 59)
$lblUseMFA.Font = 'Microsoft Sans Serif,10'

$rdoYesMFA = New-Object system.Windows.Forms.RadioButton
$rdoYesMFA.text = "Yes"
$rdoYesMFA.AutoSize = $true
$rdoYesMFA.width = 104
$rdoYesMFA.height = 20
$rdoYesMFA.location = New-Object System.Drawing.Point(297, 55)
$rdoYesMFA.Font = 'Microsoft Sans Serif,10'

$rdoNoMFA = New-Object system.Windows.Forms.RadioButton
$rdoNoMFA.text = "No"
$rdoNoMFA.AutoSize = $true
$rdoNoMFA.width = 104
$rdoNoMFA.height = 20
$rdoNoMFA.visible = $true
$rdoNoMFA.enabled = $true
$rdoNoMFA.Checked = $true
$rdoNoMFA.location = New-Object System.Drawing.Point(347, 55)
$rdoNoMFA.Font = 'Microsoft Sans Serif,10'

$grpSiteOptions = New-Object system.Windows.Forms.Groupbox
$grpSiteOptions.height = 125
$grpSiteOptions.width = 580
$grpSiteOptions.text = "Site Options"
$grpSiteOptions.location = New-Object System.Drawing.Point(8, 108)

$lblTargetSampleText = New-Object system.Windows.Forms.Label
$lblTargetSampleText.text = "/sites/"
$lblTargetSampleText.AutoSize = $true
$lblTargetSampleText.width = 40
$lblTargetSampleText.height = 10
$lblTargetSampleText.location = New-Object System.Drawing.Point(103, 57)
$lblTargetSampleText.Font = 'Microsoft Sans Serif,10'
$lblTargetSampleText.ForeColor = "#696969"

$txtTargetSiteUrl = New-Object system.Windows.Forms.TextBox
$txtTargetSiteUrl.multiline = $false
$txtTargetSiteUrl.text = "IEP-504"
$txtTargetSiteUrl.width = 255
$txtTargetSiteUrl.height = 20
$txtTargetSiteUrl.location = New-Object System.Drawing.Point(145, 55)
$txtTargetSiteUrl.Font = 'Microsoft Sans Serif,10'

$txtOwner = New-Object system.Windows.Forms.TextBox
$txtOwner.multiline = $false
$txtOwner.width = 255
$txtOwner.height = 20
$txtOwner.location = New-Object System.Drawing.Point(145, 85)
$txtOwner.Font = 'Microsoft Sans Serif,10'

$btnStart = New-Object system.Windows.Forms.Button
$btnStart.text = "Start"
$btnStart.width = 60
$btnStart.height = 30
$btnStart.location = New-Object System.Drawing.Point(465, 245)
$btnStart.Font = 'Microsoft Sans Serif,10'

$btnExit = New-Object system.Windows.Forms.Button
$btnExit.text = "Exit"
$btnExit.width = 60
$btnExit.height = 30
$btnExit.location = New-Object System.Drawing.Point(528, 245)
$btnExit.Font = 'Microsoft Sans Serif,10'

$txtStatus = New-Object system.Windows.Forms.TextBox
$txtStatus.multiline = $true
$txtStatus.BackColor = "#f0f0f0"
$txtStatus.BorderStyle = "none"
$txtStatus.width = 445
$txtStatus.height = 80
$txtStatus.location = New-Object System.Drawing.Point(8, 240)
$txtStatus.Font = 'Courier,10'

$grpTenant.controls.AddRange(@($lblTenant, $txtTenantName, $lblTenantSampleText, $lblUseMFA, $rdoYesMFA, $rdoNoMFA))
$grpSiteOptions.controls.AddRange(@($lblSiteTitle, $lblTargetSite, $txtSiteTitle, $lblTargetSampleText, $txtTargetSiteUrl, $lblOwner, $txtOwner))
$UI.controls.AddRange(@($grpTenant, $grpSiteOptions, $txtStatus, $btnStart, $btnExit))

function provision()
{
    $tenant = $txtTenantName.Text;
    $SiteTitle = $txtSiteTitle.Text;
    $siteDescription = "This Site Collection is used to process, store and sure documentation and artifacts related to Individualized Educational Plans (IEPs) and Disability Accommodation Plans. (Section 504 Plans)."
    $targetSite = "/sites/{0}" -f $txtTargetSiteUrl.text;
    $templateFile = ".\SiteTemplate.xml"
    $adminUrl = "https://{0}-admin.sharepoint.com" -f $tenant;
    $webUrl = "https://{0}.sharepoint.com{1}" -f $tenant, $targetSite;
    $ownerEmail = $txtOwner.text
    $useMFA = $rdoYesMFA.Checked;
    if (!$useMFA) { $creds = Get-Credential -Message "Enter Tenant Administrator Credentials"}
    Set-PnPTraceLog -On -LogFile "Installation.log" -Level Debug

    #Connect to Admin Site
    try
    {
        $msg = "Connecting to SharePoint Administration Site..."
        setStatus -msg $msg
        if ($useMFA) { Connect-PnPOnline -Url $adminUrl -UseWebLogin }
        else { Connect-PnPOnline -Url $adminUrl -Credentials $creds }      
    }
    catch
    {
        $msg = "Error encountered while attempting to connect to SharePoint Adminstration Site,'{0}'. Message: {1} See the Installation log for more details." -f $adminUrl, $Error[0].Exception.Message
        displayFatalError -msg $msg;
    }
    
    #Create Site Collection
    try
    { 
        $msg = "Provisioning a new Site Collection called '{0}'. This may take several minutes. Please wait..." -f $SiteTitle
        setStatus -msg $msg
        New-PnPTenantSite -Title $SiteTitle -Description $siteDescription -Url $webUrl -Owner $ownerEmail -TimeZone 0 -Template "STS#0" -RemoveDeletedSite -Force -Wait
    }
    catch
    { 
        if ($Error[0].Exception.Message -match "A Site already exists")
        {
            $siteExists = $true
            $msg = "There is an existing Site found at {0}. Do you wish to modify the existing site for use with this application?" -f $webUrl
            $msgBoxInput = [System.Windows.Forms.MessageBox]::Show($msg, 'Warning', 'YesNo', 'Question')
            switch ($msgBoxInput)
            {
                'Yes' { setStatus -msg "Preparing to use Existing Site..."       }
                'No' { [System.Windows.Forms.MessageBox]::Show("The existing site will not be changed. The process will now end.", "Information", "Ok", "Information"); [System.Environment]::Exit(0); }
                'Cancel' { [System.Windows.Forms.MessageBox]::Show("The existing site will not be changed. The process will now end.", "Information", "Ok", "Information") }
            } 
        }
        else
        {
            $msg = "Error encountered while attempting to Create a new Site Collection at '{0}'. Message: {1} See the Installation log for more details." -f $webUrl, $Error[0].Exception.Message
            displayFatalError -msg $msg;       
        }
    }
    
    #Connect to New Site
    try
    {
        if ($siteExists) { $msg = "Connecting to existing site collection..." }
        else { $msg = "Connecting to new site collection..." }
        setStatus -msg $msg
        if ($useMFA) { Connect-PnPOnline -Url $webUrl -UseWebLogin }
        else { Connect-PnPOnline -Url $webUrl -Credentials $creds }        
    }
    catch
    { 
        $msg = "Error encountered while attempting to connect to the Site Collection at '{0}'. Message: {1} See the Installation log for more details." -f $webUrl, $Error[0].Exception.Message 
        displayFatalError -msg $msg;
    }
    
    #Apply Educational Plan Site Template
    try
    {
        $msg = "Provisioning Educational Plan Site Framework. This will take several minutes.  Please Stand by..." 
        setStatus -msg $msg
        Apply-PnPProvisioningTemplate -Path $templateFile
    }
    catch
    { 
        $msg = "Error encountered while attempting to provision framework at {0}. Message: {1} See the Installation log for more details." -f $webUrl, $Error[0].Exception.Message
        displayFatalError -msg $msg;
    }
    
    if (!$siteExists)
    {
        #if this was a brand new site
        $msg = "Finalizing Installation..."
        setStatus -msg $msg
        try
        {
            #Disaable Site Notebook
            $notebookFeature = Get-PnPFeature SiteNotebook
            $notebookFeature | % { Disable-PnPFeature -Identity $_.DefinitionId }
    
            #Get rid default links and document library.
            Get-PnPNavigationNode -Location QuickLaunch | Remove-PnPNavigationNode -Force
            Remove-PnPList -Identity Documents -Force   
        }
        catch
        {
            $msg = "Error encountered while setting site defaults. This is NOT critical to the functionality of the site. Message: {0} See the Installation log for more details." -f $Error[0].Exception.Message
            [System.Windows.Forms.MessageBox]::Show($msg, "Warning", "Ok", "Warning")
        }
    }
    $msg = "Site has been provisioned successfully. Your site is located at `n{0}." -f $webUrl 
    [System.Windows.Forms.MessageBox]::Show($msg, "Success", "Ok", "Information" )
}

function setStatus($msg)
{
    if ($msg.Length -ge 240) {$txtStatus.Text = "$($msg.Substring(0,200))... See Installation Log for more Info."}
    else {$txtStatus.Text = $($msg) } 
}

function displayFatalError($msg)
{
    [System.Windows.Forms.MessageBox]::Show($msg, 'Error', 0, 'Error');
    [System.Environment]::Exit(0)
}

$btnStart.Add_Click( { provision })
$btnExit.Add_Click( { $UI.Close() })

[void]$UI.ShowDialog()