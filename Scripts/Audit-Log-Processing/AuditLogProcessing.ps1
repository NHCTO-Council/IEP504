#Requires -Version 5.0
#Requires -Modules SharePointPnPPowerShellOnline

<#
.SYNOPSIS
  IEP/504 Site Audit Log Processing 
.DESCRIPTION
  Requires configuration variables to be specified in IEP_504Configuration.ps1 file.
  Queries a data extract from MSOnline Unified Audit Logging system, for use in auditing educational plan document access.
.INPUTS
  None
.OUTPUTS
  Log file as defined in IEP_504Configuration.ps1
.NOTES
  Version:        4.0.190520
  Author:         Jeremy M. Morel, Axis Business Solutions, Ltd
  Creation Date:  8/31/2018
#>

#Initialize Function Library and Configurations
#==============================================
. ".\resources\AuditLog_Functions.ps1"
. ".\IEP504_Configuration.ps1" 

Update-Status -Message "Process Awakened." -Level Info

Connect-AuditLog -credentialLocation $config.CredentialLocation

Connect-SharePoint -siteUrl $config.SiteUrl -useMFA $config.UseMFA -credentialLocation $config.CredentialLocation

Get-LogData -timeToGoBack $config.AuditTimeToGoBack `
            -intervalMinutes $config.AuditIntervalMinutes `
            -operations $config.AuditOperations `
            -recordType $config.AuditRecordType `
            -resultSize $config.AuditResultChunkSize `
            -retries $config.AuditRetries `
            -auditListName $config.AuditListTitle 

Disconnect-SharePoint

Disconnect-AuditLog

Update-Status -Message "Process Ended.  Awaiting Next Cycle." -Level Info