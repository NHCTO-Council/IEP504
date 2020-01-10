<<<<<<< HEAD
#Requires -Version 5.0
#Requires -Modules SharePointPnPPowerShellOnline
<#
.SYNOPSIS
  IEP/504 Document Processing Function Library
.NOTES
  Version:        4.0.190919
  Author:         Jeremy M. Morel, Axis Business Solutions, Ltd
  Creation Date:  8/31/2018

  Change log:
  07/31/2019 - Updated date range logic to use simpler function, improve performance.
             - Resolved character errors.
             - Resolved issue where file renaming logic displayed during simulation mode despite being disabled.
  08/16/2019 - Resolved Issue with internal name of school abbreviation field, which impacts whether it is treated as multi or single valued. (School vs School0).
  08/28/2019 - Added new logic to substitute a default user for case manager when it is invalid or blank.
             - Revised error handling logic around case manager and teacher permissions so that the process will not crash on exception.
             - Revised School Group logic so that the district group is added to every item.
  09/04/2019 - Added Get-IdsForEmailAddresses function to work around an issue that PnP seems to have with adding multiple users by an email address array. Ids prove to work better.
  09/10/2019 - Revised routines which add Teachers to the library column and grant item permission.  This resolves an issue with a dependency due to deferred update.
  09/16/2019 - Ensured trimming of white space when inspecting student ID numbers from the dropbox.
  09/19/2019 - Resolved issue with Grant-ItemPermission function that will enable the process to still process a permission update by falling back to the name if the email is invalid.
#>

#region Core Utility Functions =============================================

function Coalesce
{
    param(
        [PSObject]$a,
        [PSObject]$b
    )
    if ($a -ne $null) { $a } else { $b }
}

Function Compare-ObjectProperties
{
    Param(
        [Parameter(Mandatory = $true)]
        [PSObject]$ReferenceObject,
        [Parameter(Mandatory = $true)]
        [PSObject]$DifferenceObject,
        [Parameter(Mandatory = $false)]
        [PSObject]$Key = "SideIndicator"
    )
    $objprops = $ReferenceObject | Get-Member -MemberType Property, NoteProperty | ForEach-Object Name
    $objprops += $DifferenceObject | Get-Member -MemberType Property, NoteProperty | ForEach-Object Name
    $objprops = $objprops | Sort | Select-Object -Unique
    $diffs = @()
    foreach ($objprop in $objprops)
    {
        $diff = Compare-Object $ReferenceObject $DifferenceObject -Property $objprop
        if ($diff)
        {
            $diffprops = @{
                $($Key)  = $ReferenceObject.$Key
                Property = $objprop
                Current  = ($diff | Where-Object {$_.SideIndicator -eq '<='} | ForEach-Object $($objprop))
                Update   = ($diff | Where-Object {$_.SideIndicator -eq '=>'} | ForEach-Object $($objprop))
            }
            $diffs += New-Object PSObject -Property $diffprops
        }
    }
    if ($diffs) {return ($diffs | Select-Object $Key, Property, Current, Update | Where-Object Property -NE "SideIndicator") }

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
            Connect-PnPOnline -Url $siteUrl -Credentials $credential
        }
        else
        { Connect-PnPOnline -Url $siteUrl -UseWebLogin }
        Update-Status -Message "Conncted to SharePoint Site: $((Get-PnPConnection).Url)" -level Info
    }
    Catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error }
}

function Disconnect-SharePoint
{
    Disconnect-PnPOnline
}

function Get-AllItems
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$listName
    )
    try { return Get-PnPListItem -List $listName; }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
}

function Get-DataFromExport
{
    [OutputType([System.Data.DataTable])]
    Param (
        [Parameter(Mandatory = $true)]
        [String]$fileToRead,

        [Parameter(Mandatory = $true)]
        [String[]]$columns,

        [Parameter(Mandatory = $false)]
        [String]$delimeter,

        [Parameter(Mandatory = $false)]
        [String]$mergeColumn, #If specified, will merge rows, using this as the unique identifier.

        [Parameter(Mandatory = $false)]
        [String]$mergeCharacter = ";" #Separator character to use for any merged values.
    )
    try
    {

        # Create DataTable and Schema (columns)
        $datatable = New-Object System.Data.DataTable("data")
        $columns| ForEach-Object { $datatable.Columns.Add($_) } | Out-Null;

        # Load Export File to memory.
        Update-Status -Message "Reading file '$(Split-Path $fileToRead -leaf)' to Memory." -Level Info;
        if (!$delimeter) { $importedData = Import-CSV $fileToRead -Delimiter "`t"; }
        else { $importedData = Import-CSV $fileToRead -Delimiter $delimeter ; }

        foreach ($item in $importedData)
        {

            $classStartDate = ([DateTime]$item.DATEENROLLED);
            $classEndDate = ([DateTime]$item.DATELEFT);
            $startDateBuffered = $classStartDate.AddDays(-$config.GrantAccessDaysBeforeStart);
            $endDateBuffered = $classEndDate.AddDays($config.GrantAccessDaysAfterEnd);
            if (IsTodayInRange $startDateBuffered $endDateBuffered)           
            {
                $row = $datatable.NewRow()
                foreach ($column in $columns) { $row[$column] = $item.$column }
                $datatable.Rows.Add($row) | Out-Null
            }
            else {$removedCounter++; }
        }

        Update-Status -Message "$($importedData.count) rows read from '$(Split-Path $fileToRead -leaf)'." -level Info
        if ($removedCounter -gt 0)
        { Update-Status -Message "$($removedCounter) rows removed which were outside the Allowed Access Date Range." -level Warn}
        if (($datatable.Rows.Count -gt 0) -and $mergeColumn)
        {
            # If we should merge multiple rows, Get Unique values of merge column.
            # then use those values to iterate the larger list.  Attempt to merge values found in a given column, when the mergecolumn isn't unique.
            # we need to remove the date columns because they cause issues with merging.
               $datatable.Columns.Remove("DATEENROLLED")
               $datatable.Columns.Remove("DATELEFT")
               $columns = $datatable.Columns|Select ColumnName -ExpandProperty ColumnName

            $uniqueEntries = $datatable.DefaultView.ToTable($true, $mergeColumn)
            Update-Status -Message "Merging Values found in $($uniqueEntries.Rows.Count) unique entries." -level Warn

            # Create a new datatable with the same columns, used for holding the merge results.
            $datatable_merged = $datatable.Clone();
            foreach ($dataRow in $uniqueEntries)
            {
                #get all matching rows for the given unique entry
                $dataview	= New-Object System.Data.DataView($datatable)
                $dataview.RowFilter = "$($mergeColumn) = '$($dataRow.$mergeColumn)'"

                #Perform merge where necessary
                $mergedRow = $datatable_merged.NewRow()
                foreach ($column in $columns)
                { $mergedRow[$column] = ($dataview.$($column)|Select-Object -Unique) -join $mergeCharacter }

                $datatable_merged.Rows.Add($mergedRow) | Out-Null
            }
            return [System.Data.DataTable]$datatable_merged;
        }
        else
        {
            return [System.Data.DataTable]$datatable;
        }
    }
    catch
    {
        Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error
    }

}

function Get-DocumentByFileName
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$ListName,
        [Parameter(Mandatory = $true)]
        [String]$FileName
    )  try
    {
        $result = Get-PnPListItem -List $ListName | Where-Object { $_["FileLeafRef"] -EQ $FileName; }
        if ($result.count -ne 1) # Either nothing or not unique
        { throw [System.IO.FileNotFoundException] "Could not return a single file. Result count: $($result.count)." }
        return $result;
    }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
}

function Get-SharePointUser
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$userName
    )
    Try
    {
       # Account for cases where we are looking up the unknown user and they might not have an email. 
       if ($userName -eq $config.unknownUserId){ $user = (Get-PNPUser | Where-Object LoginName -match $config.unknownUserId) }
       else
       {
        $safeUser = (Get-PNPUser | Where-Object LoginName -match $userName);
        if (!$safeUser)
        {
          # Couldn't match on a LoginName. Try e-mail.
          $safeUser = (Get-PNPUser | Where-Object Email -match $userName);
        }
        if (!$safeUser)
        {
          #Still couldn't match.  Return the default user and issue a warning.
          Update-Status -Message "Could not find a SharePoint user associated with '$($userName)'. Using the default user." -Level Error
          $safeUser = (Get-PNPUser | Where-Object LoginName -match $config.unknownUserId)
        }
        $user = (Get-PNPWeb).EnsureUser($safeUser.LoginName);
        $ctx = Get-PnPContext;
        $ctx.Load($user);
        Invoke-PnPQuery;
       }
        return $user;
    }
    Catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error }
}

function Grant-ItemPermission
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$listName,

        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.ListItem]$item,

        [Parameter(Mandatory = $true)]
        [String]$userName,

        [Parameter(Mandatory = $false)]
        [Boolean]$isGroup = $false,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Full Control", "Edit", "Contribute", "Read")]
        [string]$role,
        
        [Parameter(Mandatory = $false)]
        [Boolean]$clearExisting = $false

    )
    try
    {
        if ($isGroup)
        { 
            if ($clearExisting) 
            { 
                Set-PnPListItemPermission -list $listName -Identity $item -Group $userName -AddRole $role -ClearExisting
            }
            else 
            {
                Set-PnPListItemPermission -list $listName -Identity $item -Group $userName -AddRole $role
            }
            
        }
        else
        {
            $user = Get-SharePointUser $userName;
            if ($clearExisting)
            {
                #Prefer adding user by email, but title will work.
                try { Set-PnPListItemPermission -list $listName -Identity $item -user $user.Email -AddRole $role -ClearExisting }
                catch 
                { 
                  Update-Status -Message "The email address assigned to user '$($userName)' ($($user.Email)) may be invalid." -Level Warn
                  Set-PnPListItemPermission -list $listName -Identity $item -user $user.Title -AddRole $role -ClearExisting
                 }
            }

            else 
            { 
                #Prefer adding user by email, but title will work.
                try { Set-PnPListItemPermission -list $listName -Identity $item -user $user.Email -AddRole $role }
                catch 
                { 
                  Update-Status -Message "The email address assigned to user '$($userName)' ($($user.Email)) may be invalid." -Level Warn
                  Set-PnPListItemPermission -list $listName -Identity $item -user $user.Title -AddRole $role
                 }
            }   
        }
    }
    catch
    { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
}

Function IsTodayInRange([Datetime]$start,[Datetime]$end)
{
	$d = get-date
	if (($d -ge $start) -or ($d -le $end))
	{
		return $true
	}
	else
	{
		return $false
	}
}

function Move-DocumentItem
{
    Param (
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.ListItem]$itemToMove,
        [Parameter(Mandatory = $true)]
        [String]$DestinationLibraryNameOrURL,
        [Parameter(Mandatory = $false)]
        [Boolean]$OverWriteIfExists = $false
    )
    $fileName = $itemToMove.FieldValues.FileLeafRef;
    if ($DestinationLibraryNameOrURL -match "/")
    { $targetFile = $DestinationLibraryNameOrURL / $fileName ; }
    else { $targetPath = (Get-PnPList $DestinationLibraryNameOrURL).RootFolder.ServerRelativeUrl; $targetFile = "$($targetPath)/$($fileName)"; }
    Update-Status -Message "Moving '$($fileName)' from '$($itemToMove["FileDirRef"])' to '$($targetFile.Replace($fileName,''))'." -level Info;
    try
    {
        if ($OverWriteIfExists)
        { Move-PnPFile -ServerRelativeUrl $itemToMove.FieldValues.FileRef -TargetUrl $targetFile -OverwriteIfAlreadyExists -Force; }
        else { Move-PnPFile -ServerRelativeUrl $itemToMove.FieldValues.FileRef -TargetUrl $targetFile -Force; }
    }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
    Invoke-PnPQuery #needed to committ the move
}

function Rename-Key
{
    [cmdletbinding(SupportsShouldProcess = $True, DefaultParameterSetName = "Pipeline")]

    Param(
        [parameter(Position = 0, Mandatory = $True,
            HelpMessage = "Enter the name of your hash table variable without the `$",
            ParameterSetName = "Name")]
        [ValidateNotNullorEmpty()]
        [string]$Name,
        [parameter(Position = 0, Mandatory = $True,
            ValueFromPipeline = $True, ParameterSetName = "Pipeline")]
        [ValidateNotNullorEmpty()]
        [object]$InputObject,
        [parameter(position = 1, Mandatory = $True, HelpMessage = "Enter the existing key name you want to rename")]
        [ValidateNotNullorEmpty()]
        [string]$Key,
        [parameter(position = 2, Mandatory = $True, HelpMessage = "Enter the NEW key name")]
        [ValidateNotNullorEmpty()]
        [string]$NewKey,
        [switch]$Passthru,
        [ValidateSet("Global", "Local", "Script", "Private", 0, 1, 2, 3)]
        [ValidateNotNullOrEmpty()]
        [string]$Scope = "Global"
    )

    Begin
    {
        Write-Verbose -Message "Starting $($MyInvocation.Mycommand)"
        Write-verbose "using parameter set $($PSCmdlet.ParameterSetName)"
    }

    Process
    {
        #validate Key and NewKey are not the same
        if ($key -eq $NewKey)
        {
            Write-Warning "The values you specified for -Key and -NewKey appear to be the same. Names are NOT case-sensitive"
            Return
        }

        Try
        {
            #validate variable is a hash table
            if ($InputObject)
            {
                $name = "tmpInputHash"
                Set-Variable -Name $name -Scope $scope -value $InputObject
                $Passthru = $True
            }

            Write-Verbose (get-variable -Scope $scope | out-string)
            Write-Verbose "Validating $name as a hashtable in $Scope scope."
            #get the variable
            $var = Get-Variable -Name $name -Scope $Scope -ErrorAction Stop

            if ( $var.Value -is [hashtable])
            {
                #create a temporary copy
                Write-Verbose "Cloning a temporary hashtable"
                <#
                Use the clone method to create a separate copy.
                If you just assign the value to $temphash, the
                two hash tables are linked in memory so changes
                to $tempHash are also applied to the original
                object.
            #>
                $tempHash = $var.Value.Clone()
                #validate key exists
                Write-Verbose "Validating key $key"
                if ($tempHash.Contains($key))
                {
                    #create a key with the new name using the value from the old key
                    Write-Verbose "Adding new key $newKey to the temporary hashtable"
                    $tempHash.Add($NewKey, $tempHash.$Key)
                    #remove the old key
                    Write-Verbose "Removing $key"
                    $tempHash.Remove($Key)
                    #write the new value to the variable
                    Write-Verbose "Writing the new hashtable"
                    Write-Verbose ($tempHash | out-string)
                    Set-Variable -Name $Name -Value $tempHash -Scope $Scope -Force -PassThru:$Passthru |
                        Select-Object -ExpandProperty Value
                }
                else
                {
                    Write-Warning "Can't find a key called $Key in `$$Name"
                }
            }
            else
            {
                Write-Warning "The variable $name does not appear to be a hash table."
            }
        } #Try

        Catch
        {
            Write-Warning "Failed to find a variable with a name of $Name."
        }

        Write-Verbose "Rename complete."
    } #Process

    End
    {
        #clean up any temporary variables
        Get-Variable -Name tmpInputHash -Scope $scope -ErrorAction SilentlyContinue |
            Remove-Variable -Scope $scope
        Write-Verbose -Message "Ending $($MyInvocation.Mycommand)"
    } #end

}

function Reset-ItemPermission
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$listName,

        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.ListItem]$item
    )

    try { Set-PnPListItemPermission -List $listName -Identity $item -InheritPermissions; }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
}

function Resolve-MultiValuedFields
{
    Param (
        [Parameter(Mandatory = $true)]
        [Hashtable]$proposedMetadata,

        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.ListItem]$item
    )
    $keys = $proposedMetadata.Keys
    $keys | ForEach-Object {if (!$item.FieldValues.ContainsKey($_) -and $item.FieldValues.ContainsKey($_ + "0")) {$proposedMetadata = $proposedMetadata | Rename-Key -Key $_ -NewKey $_"0"}}
    return $proposedMetadata
}

function Update-Status
{
    Param (
        [Parameter(Mandatory = $true)]
        [Object]$Message,

        [Parameter(Mandatory = $false)]
        [string]$Path = $config.OutputLog,

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

#endregion Core Utility Functions ==========================================

#region IEP/504 Processing Helper Functions ================================

function Get-IdsForEmailAddresses
{
    Param (
        [Parameter(Mandatory = $true)]
        [object]$emailAddressArray
    )

    #Check if we have retrieved site users before, if so use that to save time
    if(!$thisSitesUsers) { $thisSitesUsers = Get-PnPUser }

    $idArray = @()
    
    foreach ($emailAddress in $emailAddressArray)
    {
      try 
      {
        $idArray += $thisSitesUsers | ? { $_.Email -eq $emailAddress.trim() } | select -ExpandProperty Id
      }
      catch 
      {
        Update-Status -Message "User '$($emailAddress)' is does not appear to be a member of this site." -Level Error;
      }
      
    }
   
    return $idArray
}

function Get-NewFileName
{
    Param (
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.ListItem] $Document,
        [Parameter(Mandatory = $true)]
        [String] $StudentId,        
        [Parameter(Mandatory = $true)]
        [String] $DocumentType
    )
    
    Try
    {
         # Rename file by concatenating Student Id and Document Type. #########################
         $oldFileName = [System.IO.Path]::GetFilenameWithoutExtension($document["FileLeafRef"])
         $newFileName = "$($StudentId)_$($DocumentType)"
         return $($Document["FileLeafRef"]).Replace($oldfileName, $newFileName)           
    }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
}

function Resolve-Bookmark
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$bookmarkFile,
        [Parameter(Mandatory = $true)]
        [Object[]]$recentData,
        [Parameter(Mandatory = $true)]
        [Object[]]$columns
    )

    If ($bookmarkFile -and (Test-Path $bookmarkFile))
    {
        #Remove Date Columns from header since they will have been stripped.
        
        Update-Status  -Message "Using Bookmark file '$($bookmarkFile)' to resolve recent changes." -level Info
        $bookmark = (import-csv $bookmarkFile -Delimiter "`t")
        $comparisons = Compare-Object $recentData $bookmark -Property $columns -PassThru| Group-Object -Property STUDENT_NUMBER
        $additions = $comparisons| Where-Object Count -eq 1 | Select-Object -ExpandProperty Group | Select-Object STUDENT_NUMBER, FIRST_NAME, LAST_NAME

        if ($comparisons)
        {
            Update-Status -Message "Bookmark Comparison has identified $($comparisons.Values.count) student records which are new or updated." -level Info
            $updates = $comparisons| Where-Object Count -eq 2 |Select-Object $_ | ForEach-Object {Compare-ObjectProperties $_.Group[0] $_.Group[1] -Key STUDENT_NUMBER}

            If ($updates)
            {
                $updateMessage = $updates | Sort-Object STUDENT_NUMBER |Format-Table Property, @{L = 'Current Value'; E = {$_.Current}}, @{L = 'Will be updated to'; E = {$_.Update}} -GroupBy STUDENT_NUMBER;
                Update-Status -Message "The Following Updates will be made:" -Level Warn
                Update-Status -Message $UpdateMessage -Level Data
            }

            If ($additions)
            {
                Update-Status -Message "The Following are new entries which will be inspected:" -Level Warn
                Update-Status -Message $additions -Level Data
            }
            Update-Status -Message "Updating Bookmark." -level Info
            $recentData |Export-Csv -Path $config.BookMarkFile -NoTypeInformation -Delimiter "`t" -Force
            return ($recentData | Where-Object STUDENT_NUMBER -in $comparisons.Name);
        }
    }
    else
    {
        Update-Status -Message "No Existing Bookmark was located.  Setting one for next time..." -Level Warn
        $recentData |Export-Csv -Path $config.BookMarkFile -NoTypeInformation -Delimiter "`t" -Force
        return $recentData;
    }

}

function Resolve-Documents
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$documentLibraryName,
        [Parameter(Mandatory = $true)]
        [Object[]]$studentData
    )
    try
    {
        $totalCount = $studentData.count;
        $i = 0
        foreach ($student in $studentData)
        {
          $roundPercent = [math]::Round($i*100/$totalCount)
          Write-Progress -PercentComplete ($i*100/$totalCount) -Activity "Records Processed: $i/$totalCount ($roundPercent%)"  -Status "Updating Document Metadata and Permissions"
          Update-Status -Message "Records Processed: $i/$totalCount ($roundPercent%)" -level Info

            $itemQuery = "<View><Query><Where><Eq><FieldRef Name='StudentId'/><Value Type='Text'>$($student.STUDENT_NUMBER.trim())</Value></Eq></Where></Query></View>"
            [Array]$studentDocs = Get-PnPListItem -List $documentLibraryName -Query $itemQuery
            Update-Status -Message "Located $($studentDocs.Count) Document(s) for student $($student.STUDENT_NUMBER)." -level Info
            foreach ($document in $studentDocs)
            { Set-MetadataAndPermission -listName $documentLibraryName -Document $document -studentData $student }
        $i++
        }

    }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }

}

function Resolve-Dropbox
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$dropBoxName,
        [Parameter(Mandatory = $true)]
        [String]$documentLibraryName,
        [Parameter(Mandatory = $true)]
        [Object[]]$studentData
    )
    try
    {
        Update-Status -Message "Retreiving Items from the Dropbox Library ($($dropBoxName))." -level Info
        $allDocuments = Get-AllItems -listName $dropBoxName

        if (!$allDocuments.Count)
        { Update-Status -Message $("No documents were found in the Dropbox Library.") -Level Warn; }
        else { Update-Status -Message $("Found $($allDocuments.Count) document(s) in the Dropbox Library.") -level Info; }

        foreach ($document in $allDocuments)
        {
            $docName = $document['FileLeafRef'];
            $studentId = $document['StudentId'];
            $docType = $document['SSDocumentType'];
            if (!$studentId -or !$docType) { Update-Status -Message $("Skipping Document '$($docName)' because it has not been assigned a Student ID and/or Document Type.") -level Warn; continue; }
            #Document has been assigned a student ID, Process Intake Procedure.

            Update-Status -Message $("Processing Document: '$($docName)'.") -level Info;
            $student = $studentData | Where-Object { $_.STUDENT_NUMBER.trim() -eq $document["StudentId"].trim() | Select-Object -First 1 };
            if (!$student) { Update-Status -Message $("The Student ID, $($studentId), assigned to document '$($docName)' does not match Valid Student Data. The file has been skipped.") -level Error; continue; }

            if (!$config.SimulationMode)
            {
                Move-DocumentItem -itemToMove $document -DestinationLibraryNameOrURL $documentLibraryName -OverWriteIfExists $true
                #Get Moved Document and set the metadata
                $movedDocument = Get-DocumentByFileName -ListName $documentLibraryName -FileName $document["FileLeafRef"] ;
                if ($movedDocument) { Set-MetadataAndPermission -listName $documentLibraryName -Document $movedDocument -studentData $student }
            }
            else {Update-Status -Message "Simulation Mode: Would have moved file '$($document["FileLeafRef"])' to '$($documentLibraryName)'." -level Info }
        }
    }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
}
function Set-MetadataAndPermission
{
    Param (
        [Parameter(Mandatory = $true)]
        [String] $listName,
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.ListItem] $Document,
        [Parameter(Mandatory = $true)]
        [Object[]] $studentData
    )
    Try
    {

        $metadata = @{
            CaseManager      = [String](Coalesce $studentData.CASE_MANAGER $config.unknownUserId);
           #Teachers         = [String[]](Coalesce @($studentData.TEACHER_EMAIL.TrimStart(';').TrimEnd(';').Split(';')) $config.unknownUserId);
            School           = [String](Coalesce $studentData.SCHOOL_ABBREVIATION $null);
            StudentFirstName = [String](Coalesce $studentData.FIRST_NAME $null);
            StudentLastName  = [String](Coalesce $studentData.LAST_NAME $null);
            StudentWebId     = [String](Coalesce $studentData.STUDENT_WEB_ID $null);
            GradeLevel       = [String](Coalesce $studentData.GRADE_LEVEL $null);
            GraduationYear   = [String](Coalesce $studentData.CLASSOF $null);
            HomeRoom         = [String](Coalesce $studentData.HOME_ROOM $null);
            SchoolTeam       = [String](Coalesce $studentData.TEAM $null);
            SSDocumentType   = [String](Coalesce $studentData.IEP504 "MISC");
        }
        $teacherEmailArray = [String[]](Coalesce @($studentData.TEACHER_EMAIL.TrimStart(';').TrimEnd(';').Split(';')) $config.unknownUserId);
        
        #remove any blank values
        $teacherEmailArray = [String[]]($teacherEmailArray | ? {$_.trim()});

        Update-Status -Message "Setting Metadata on '$($Document["FileLeafRef"])'. ------------------------------------" -level Info

        If (!$config.SimulationMode)
        {
            try
            {

                #resolve multi-valued field issue
                $metadata = Resolve-MultiValuedFields -proposedMetadata $metadata -item $Document


                if ($Document.FieldValues.CheckoutUser -and $config.ForceCheckIn)
                {
                    Update-Status -Message "File '$($Document["FileLeafRef"])' is Locked for editing by $($Document.FieldValues.CheckoutUser.Email). Will try to force check-in." -Level Warn
                    Set-PNPFileCheckedIn -url $Document["FileRef"] -CheckinType OverwriteCheckIn -Comment "Checked in by Administrator."
                }
                if ($config.Renamefile)
                {
                  $proposedFileName = Get-NewFileName -Document $Document -StudentId $Document['StudentId'] -DocumentType $metadata.SSDocumentType
                  If ($proposedFileName -notlike $Document["FileLeafRef"]) 
                  {
                    Update-Status -Message "Renaming file '$($Document["FileLeafRef"])' to '$($proposedFileName)'." -level Info 
                    #Document is not already named using the desired convention, so perform the rename.
                    try 
                    {
                        $oldName = $Document["FileLeafRef"];
                        $Document["FileLeafRef"] = $proposedFileName;
                        $Document.Update();
                        Set-PnPListItem -List $listName -Identity $Document -ErrorAction stop |Out-Null
                    }
                    catch
                    { 
                      Update-Status -Message "Unable to rename the file '$($oldName)' to '$($proposedFileName)'.ErrorDetail: $($Error[0].Exception.Message)" -level Error
                      Update-Status -Message "Document could not be remaned, but will still be secured accordingly." -level Warn
                    }
                  }
                  else { Update-Status -Message "File does not require renaming." -Level Info} 
                  
                }
                try 
                {
                  Update-Status -Message "Updating Metadata..." -Level Info;
                  Set-PnPListItem -List $listName -Identity $Document -Values $metadata -ErrorAction Stop |Out-Null;
                  Update-TeacherColumn -teacherEmailArray $teacherEmailArray -document $Document;
                }
                catch
                {
                  Update-Status "An error occured while updating the metadata. The error was: '$($Error[0].Exception.Message)'." -Level Warn;
                }

                if ($teacherEmailArray.Count -ne $Document["Teachers"].Count) 
                { Update-Status -Message "Deferred Data write: The Teacher column count could not be verified due to SharePoint deferred update. This does not affect permissions, only the column 'Teachers' on the individual item." -Level Warn }
            }
            catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }

            #Set Permissions
            Reset-ItemPermission -listName $listName -item $Document
            Update-Status -Message "Resetting Document Permissions." -level Info
            #Set Permission for Case Manager

            if ($Document["CaseManager"])
            {
                try
                {
                    Update-Status -Message "Granting Permission to Case Manager: $($Document["CaseManager"].LookupValue)" -level Info
                    Grant-ItemPermission -listName $listName -item $Document -userName $Document["CaseManager"].Email -role Contribute -clearExisting $true
                }
                catch
                {
                    #Couldn't add case manager for some reason. Log an error, then assign the default user and issue a warning.
                    Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; 
                    Update-Status -Message "Error while trying to set the case manager set on document '$($Document["FileLeafRef"])'. Assigning the default user $($config.unknownUserId)." -Level Warn
                    Grant-ItemPermission -listName $listName -item $Document -userName $config.unknownUserId -role Contribute -clearExisting $true
                }
            }
            else
            {
                #No Case Manager was set. Use the default user and issue a warning.
                Update-Status -Message "No Case Manager was set on document '$($Document["FileLeafRef"])'. Assigning the default user $($config.unknownUserId)." -Level Warn
                Grant-ItemPermission -listName $listName -item $Document -userName $config.unknownUserId -role Contribute -clearExisting $true
            }
            #Set Permissions for teachers.
            Update-Status -Message "Granting Permission to $($teacherEmailArray.Count) teachers." -level Info 

            foreach ($teacher in $teacherEmailArray)
            {
                Update-Status -Message "Attempting to Grant Permission to teacher: $($teacher)." -Level Info
                try { Grant-ItemPermission -listName $listName -item $Document -userName $teacher -role Read }
                catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
            }
            #Set Permissions for School Team (if applicable)
            if ($Document["School"])
            {
                Update-Status -Message "Granting Permission to school Team: $($config.SchoolTeamGroupPrefix + $Document["School"])." -level Info
                Grant-ItemPermission -listName $documentLibraryName -item $Document -userName $($config.SchoolTeamGroupPrefix + $Document["School"]) -isGroup $true -role Contribute
                Update-Status -Message "Granting Permission to District Team: $($config.SchoolTeamGroupPrefix + "District")." -Level Info
                Grant-ItemPermission -listName $documentLibraryName -item $Document -userName $($config.SchoolTeamGroupPrefix + "District") -isGroup $true -role Contribute
            }

            #The below section will resolve any old entries where the School field may have been multi-valued.
            if ($Document["School0"])
            {
                Update-Status -Message "Granting Permission to school Team: $($config.SchoolTeamGroupPrefix + $Document["School0"])." -level Info
                Grant-ItemPermission -listName $documentLibraryName -item $Document -userName $($config.SchoolTeamGroupPrefix + $Document["School0"]) -isGroup $true -role Contribute
                Update-Status -Message "Granting Permission to District Team: $($config.SchoolTeamGroupPrefix + "District")." -Level Info
                Grant-ItemPermission -listName $documentLibraryName -item $Document -userName $($config.SchoolTeamGroupPrefix + "District") -isGroup $true -role Contribute
            }
        }
        else
        {
            Update-Status -Message "Simulation Mode: Would have set the following Metadata:" -Level Info
            Update-Status -Message $metadata -Level Data;
            Update-Status -Message "Document would have been shared with the following Teachers:"
            Update-Status -Message $teacherEmailArray -level Data;
            $proposedFileName = Get-NewFileName -Document $Document -StudentId $Document['StudentId'] -DocumentType $metadata.SSDocumentType
            If ($config.Renamefile)
            {
               If ($proposedFileName -notlike $Document["FileLeafRef"]) { Update-Status -Message "Would have renamed file '$($Document["FileLeafRef"])' to '$($proposedFileName)'." -Level Info}
            }
        }
    }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
    Update-Status -Message "------------------------------------------------------------------------" -level Info
}

 function Update-TeacherColumn
 {
     Param (
        [Parameter(Mandatory = $true)]
        [string[]] $teacherEmailArray,
        [Parameter(Mandatory = $true)] 
        [Microsoft.SharePoint.Client.ListItem] $document
    )
      # Process Teachers
      
      $lookupValueCollection = @()
      $teacherEmailArray | % {  try 
                                {
                                    $context = Get-PnPContext;
                                    $spUser = $context.Web.EnsureUser($_);
                                    $context.Load($spUser);
                                    $context.ExecuteQuery();
                                 }
                                 catch
                                 {
                                   Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error;
                                   $spUser = $null;
                                 }
                                 
                                if($spUser -ne $null){
                                
                                $spuserValue = New-Object Microsoft.SharePoint.Client.FieldUserValue;
                                $spuserValue.LookupId = $spuser.Id;
                                $lookupValueCollection += $spuserValue;
                                }  
                              }
      $userValueCollection = [Microsoft.SharePoint.Client.FieldUserValue[]]$lookupValueCollection;
      $document["Teachers"] = $userValueCollection;
      $document.Update();
      Invoke-PnPQuery;
}
=======
#Requires -Version 5.0
#Requires -Modules SharePointPnPPowerShellOnline

#region Core Utility Functions =============================================

function Coalesce
{
    param(
        [PSObject]$a,
        [PSObject]$b
    )
    if ($a -ne $null) { $a } else { $b }
}

Function Compare-ObjectProperties
{
    Param(
        [Parameter(Mandatory = $true)]
        [PSObject]$ReferenceObject,
        [Parameter(Mandatory = $true)]
        [PSObject]$DifferenceObject,
        [Parameter(Mandatory = $false)]
        [PSObject]$Key = "SideIndicator"
    )
    $objprops = $ReferenceObject | Get-Member -MemberType Property, NoteProperty | ForEach-Object Name
    $objprops += $DifferenceObject | Get-Member -MemberType Property, NoteProperty | ForEach-Object Name
    $objprops = $objprops | Sort | Select-Object -Unique
    $diffs = @()
    foreach ($objprop in $objprops)
    {
        $diff = Compare-Object $ReferenceObject $DifferenceObject -Property $objprop
        if ($diff)
        {
            $diffprops = @{
                $($Key)  = $ReferenceObject.$Key
                Property = $objprop
                Current  = ($diff | Where-Object {$_.SideIndicator -eq '<='} | ForEach-Object $($objprop))
                Update   = ($diff | Where-Object {$_.SideIndicator -eq '=>'} | ForEach-Object $($objprop))
            }
            $diffs += New-Object PSObject -Property $diffprops
        }
    }
    if ($diffs) {return ($diffs | Select-Object $Key, Property, Current, Update | Where-Object Property -NE "SideIndicator") }

}

function Connect-SharePoint
{

    Param (
        [Parameter(Mandatory = $true)]
        [String]$siteUrl,

        [Parameter(Mandatory = $false)]
        [SecureString]$credentialLocation,

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
        Update-Status -Message "Conncted to SharePoint Site: $((Get-PnPConnection).Url)"
    }
    Catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error }
}

function Disconnect-SharePoint
{
    Disconnect-PnPOnline
}

function Get-AllItems
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$listName
    )
    try { return Get-PnPListItem -List $listName; }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
}

function Get-DataFromExport
{
    [OutputType([System.Data.DataTable])]
    Param (
        [Parameter(Mandatory = $true)]
        [String]$fileToRead,

        [Parameter(Mandatory = $true)]
        [String[]]$columns,

        [Parameter(Mandatory = $false)]
        [String]$delimeter,

        [Parameter(Mandatory = $false)]
        [String]$mergeColumn, #If specified, will merge rows, using this as the unique identifier.

        [Parameter(Mandatory = $false)]
        [String]$mergeCharacter = ";" #Separator character to use for any merged values.
    )
    try
    {

        # Create DataTable and Schema (columns)
        $datatable = New-Object System.Data.DataTable("data")
        $columns| ForEach-Object { $datatable.Columns.Add($_) } | Out-Null;

        # Load Export File to memory.
        Update-Status -Message "Reading file '$(Split-Path $fileToRead -leaf)' to Memory.";
        if (!$delimeter) { $importedData = Import-CSV $fileToRead -Delimiter "`t"; }
        else { $importedData = Import-CSV $fileToRead -Delimiter $delimeter ; }

        foreach ($item in $importedData)
        {
            $today = (Get-Date) 
            $classStartDate = ([DateTime]$item.DATEENROLLED);
            $classEndDate = ([DateTime]$item.DATELEFT);
            $todayIsAfterAccessStartDate = ($today -ge $classStartDate.AddDays(-$config.GrantAccessDaysBeforeStart));
            $todayIsBeforeAccessEndDate = ($today -le $classEndDate.AddDays($config.GrantAccessDaysAfterEnd));
            if ($todayIsAfterAccessStartDate -and $todayIsBeforeAccessEndDate)
            {
                $row = $datatable.NewRow()
                foreach ($column in $columns) { $row[$column] = $item.$column }
                $datatable.Rows.Add($row) | Out-Null
            }
            else {$removedCounter++; }
        }

        Update-Status -Message "$($importedData.count) rows read from '$(Split-Path $fileToRead -leaf)'."
        if ($removedCounter -gt 0)
        { Update-Status -Message "$($removedCounter) rows removed which were outside the Allowed Access Date Range." }

        if (($datatable.Rows.Count -gt 0) -and $mergeColumn)
        {
            # If we should merge multiple rows, Get Unique values of merge column.
            # then use those values to iterate the larger list.  Attempt to merge values found in a given column, when the mergecolumn isn't unique.

            $uniqueEntries = $datatable.DefaultView.ToTable($true, $mergeColumn)
            Update-Status -Message "Merging Values found in $($uniqueEntries.Rows.Count) unique entries."

            # Create a new datatable with the same columns, used for holding the merge results.
            $datatable_merged = $datatable.Clone();
            foreach ($dataRow in $uniqueEntries)
            {
                #get all matching rows for the given unique entry
                $dataview	= New-Object System.Data.DataView($datatable)
                $dataview.RowFilter = "$($mergeColumn) = '$($dataRow.$mergeColumn)'"

                #Perform merge where necessary
                $mergedRow = $datatable_merged.NewRow()
                foreach ($column in $columns)
                { $mergedRow[$column] = ($dataview.$($column)|Select-Object -Unique) -join $mergeCharacter }

                $datatable_merged.Rows.Add($mergedRow) | Out-Null
            }
            return [System.Data.DataTable]$datatable_merged;
        }
        else
        {
            return [System.Data.DataTable]$datatable;
        }
    }
    catch
    {
        Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error
    }

}

function Get-DocumentByFileName
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$ListName,
        [Parameter(Mandatory = $true)]
        [String]$FileName
    )  try
    {
        $result = Get-PnPListItem -List $ListName | Where-Object { $_["FileLeafRef"] -EQ $FileName; }
        if ($result.count -ne 1) # Either nothing or not unique
        { throw [System.IO.FileNotFoundException] "Could not return a single file. Result count: $($result.count)." }
        return $result;
    }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
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

function Grant-ItemPermission
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$listName,

        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.ListItem]$item,

        [Parameter(Mandatory = $true)]
        [String]$userName,

        [Parameter(Mandatory = $false)]
        [Boolean]$isGroup = $false,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Full Control", "Edit", "Contribute", "Read")]
        [string]$role
    )
    try
    {
        if ($isGroup)
        { Set-PnPListItemPermission -list $listName -Identity $item -Group $userName -AddRole $role }
        else
        {
            $user = Get-SharePointUser $userName;
            Set-PnPListItemPermission -list $listName -Identity $item -user $user.Email -AddRole $role
        }
    }
    catch
    { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
}
function Move-DocumentItem
{
    Param (
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.ListItem]$itemToMove,
        [Parameter(Mandatory = $true)]
        [String]$DestinationLibraryNameOrURL,
        [Parameter(Mandatory = $false)]
        [Boolean]$OverWriteIfExists = $false
    )
    $fileName = $itemToMove.FieldValues.FileLeafRef;
    if ($DestinationLibraryNameOrURL -match "/")
    { $targetFile = $DestinationLibraryNameOrURL / $fileName ; }
    else { $targetPath = (Get-PnPList $DestinationLibraryNameOrURL).RootFolder.ServerRelativeUrl; $targetFile = "$($targetPath)/$($fileName)"; }
    Update-Status -Message "Moving '$($fileName)' from '$($itemToMove["FileDirRef"])' to '$($targetFile.Replace($fileName,''))'.";
    try
    {
        if ($OverWriteIfExists)
        { Move-PnPFile -ServerRelativeUrl $itemToMove.FieldValues.FileRef -TargetUrl $targetFile -OverwriteIfAlreadyExists -Force; }
        else { Move-PnPFile -ServerRelativeUrl $itemToMove.FieldValues.FileRef -TargetUrl $targetFile -Force; }
    }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
}

function Rename-Key
{
    [cmdletbinding(SupportsShouldProcess = $True, DefaultParameterSetName = "Pipeline")]

    Param(
        [parameter(Position = 0, Mandatory = $True,
            HelpMessage = "Enter the name of your hash table variable without the `$",
            ParameterSetName = "Name")]
        [ValidateNotNullorEmpty()]
        [string]$Name,
        [parameter(Position = 0, Mandatory = $True,
            ValueFromPipeline = $True, ParameterSetName = "Pipeline")]
        [ValidateNotNullorEmpty()]
        [object]$InputObject,
        [parameter(position = 1, Mandatory = $True, HelpMessage = "Enter the existing key name you want to rename")]
        [ValidateNotNullorEmpty()]
        [string]$Key,
        [parameter(position = 2, Mandatory = $True, HelpMessage = "Enter the NEW key name")]
        [ValidateNotNullorEmpty()]
        [string]$NewKey,
        [switch]$Passthru,
        [ValidateSet("Global", "Local", "Script", "Private", 0, 1, 2, 3)]
        [ValidateNotNullOrEmpty()]
        [string]$Scope = "Global"
    )

    Begin
    {
        Write-Verbose -Message "Starting $($MyInvocation.Mycommand)"
        Write-verbose "using parameter set $($PSCmdlet.ParameterSetName)"
    }

    Process
    {
        #validate Key and NewKey are not the same
        if ($key -eq $NewKey)
        {
            Write-Warning "The values you specified for -Key and -NewKey appear to be the same. Names are NOT case-sensitive"
            Return
        }

        Try
        {
            #validate variable is a hash table
            if ($InputObject)
            {
                $name = "tmpInputHash"
                Set-Variable -Name $name -Scope $scope -value $InputObject
                $Passthru = $True
            }

            Write-Verbose (get-variable -Scope $scope | out-string)
            Write-Verbose "Validating $name as a hashtable in $Scope scope."
            #get the variable
            $var = Get-Variable -Name $name -Scope $Scope -ErrorAction Stop

            if ( $var.Value -is [hashtable])
            {
                #create a temporary copy
                Write-Verbose "Cloning a temporary hashtable"
                <#
                Use the clone method to create a separate copy.
                If you just assign the value to $temphash, the
                two hash tables are linked in memory so changes
                to $tempHash are also applied to the original
                object.
            #>
                $tempHash = $var.Value.Clone()
                #validate key exists
                Write-Verbose "Validating key $key"
                if ($tempHash.Contains($key))
                {
                    #create a key with the new name using the value from the old key
                    Write-Verbose "Adding new key $newKey to the temporary hashtable"
                    $tempHash.Add($NewKey, $tempHash.$Key)
                    #remove the old key
                    Write-Verbose "Removing $key"
                    $tempHash.Remove($Key)
                    #write the new value to the variable
                    Write-Verbose "Writing the new hashtable"
                    Write-Verbose ($tempHash | out-string)
                    Set-Variable -Name $Name -Value $tempHash -Scope $Scope -Force -PassThru:$Passthru |
                        Select-Object -ExpandProperty Value
                }
                else
                {
                    Write-Warning "Can't find a key called $Key in `$$Name"
                }
            }
            else
            {
                Write-Warning "The variable $name does not appear to be a hash table."
            }
        } #Try

        Catch
        {
            Write-Warning "Failed to find a variable with a name of $Name."
        }

        Write-Verbose "Rename complete."
    } #Process

    End
    {
        #clean up any temporary variables
        Get-Variable -Name tmpInputHash -Scope $scope -ErrorAction SilentlyContinue |
            Remove-Variable -Scope $scope
        Write-Verbose -Message "Ending $($MyInvocation.Mycommand)"
    } #end

}

function Reset-ItemPermission
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$listName,

        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.ListItem]$item
    )

    try { Set-PnPListItemPermission -List $listName -Identity $item -InheritPermissions; }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
}

function Resolve-MultiValuedFields
{
    Param (
        [Parameter(Mandatory = $true)]
        [Hashtable]$proposedMetadata,

        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.ListItem]$item
    )
    $keys = $proposedMetadata.Keys
    $keys | ForEach-Object {if (!$item.FieldValues.ContainsKey($_) -and $item.FieldValues.ContainsKey($_ + "0")) {$proposedMetadata = $proposedMetadata | Rename-Key -Key $_ -NewKey $_"0"}}
    return $proposedMetadata
}

function Update-Status
{
    Param (
        [Parameter(Mandatory = $true)]
        [Object]$Message,

        [Parameter(Mandatory = $false)]
        [string]$Path = $config.OutputLog,

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

#endregion Core Utility Functions ==========================================

#region IEP/504 Processing Helper Functions ================================

function Resolve-Bookmark
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$bookmarkFile,
        [Parameter(Mandatory = $true)]
        [Object[]]$recentData,
        [Parameter(Mandatory = $true)]
        [Object[]]$columns
    )

    If ($bookmarkFile -and (Test-Path $bookmarkFile))
    {
        Update-Status  -Message "Using Bookmark file '$($bookmarkFile)' to resolve recent changes."
        $bookmark = Get-DataFromExport -columns $columns -fileToRead $bookmarkFile
        $comparisons = Compare-Object $recentData $bookmark -Property $columns -PassThru| Group-Object -Property STUDENT_NUMBER
        $additions = $comparisons| Where-Object Count -eq 1 | Select-Object -ExpandProperty Group | Select-Object STUDENT_NUMBER, FIRST_NAME, LAST_NAME

        if ($comparisons)
        {
            Update-Status -Message "Bookmark Comparison has identified $($comparisons.Values.count) student records which are new or updated."
            $updates = $comparisons| Where-Object Count -eq 2 |Select-Object $_ | ForEach-Object {Compare-ObjectProperties $_.Group[0] $_.Group[1] -Key STUDENT_NUMBER}

            If ($updates)
            {
                $updateMessage = $updates | Sort-Object STUDENT_NUMBER |Format-Table Property, @{L = 'Current Value'; E = {$_.Current}}, @{L = 'Will be updated to'; E = {$_.Update}} -GroupBy STUDENT_NUMBER;
                Update-Status -Message "The Following Updates will be made:" -Level Warn
                Update-Status -Message $UpdateMessage -Level Data
            }

            If ($additions)
            {
                Update-Status -Message "The Following are new entries which will be inspected:" -Level Warn
                Update-Status -Message $additions -Level Data
            }
            Update-Status -Message "Updating Bookmark."
            $recentData |Export-Csv -Path $config.BookMarkFile -NoTypeInformation -Delimiter "`t" -Force
            return ($recentData | Where-Object STUDENT_NUMBER -in $comparisons.Name);
        }
    }
    else
    {
        Update-Status -Message "No Existing Bookmark was located.  Setting one for next time..." -Level Warn
        $recentData |Export-Csv -Path $config.BookMarkFile -NoTypeInformation -Delimiter "`t" -Force
        return $recentData;
    }

}

function Resolve-Documents
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$documentLibraryName,
        [Parameter(Mandatory = $true)]
        [Object[]]$studentData
    )
    try
    {
        foreach ($student in $studentData)
        {
            $itemQuery = "<View><Query><Where><Eq><FieldRef Name='StudentId'/><Value Type='Text'>$($student.STUDENT_NUMBER)</Value></Eq></Where></Query></View>"
            [Array]$studentDocs = Get-PnPListItem -List $documentLibraryName -Query $itemQuery
            Update-Status -Message "Located $($studentDocs.Count) Document(s) for student $($student.STUDENT_NUMBER)."
            foreach ($document in $studentDocs)
            { Set-MetadataAndPermission -listName $documentLibraryName -Document $document -studentData $student }
        }

    }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }

}

function Resolve-Dropbox
{
    Param (
        [Parameter(Mandatory = $true)]
        [String]$dropBoxName,
        [Parameter(Mandatory = $true)]
        [String]$documentLibraryName,
        [Parameter(Mandatory = $true)]
        [Object[]]$studentData
    )
    try
    {
        Update-Status -Message "Retreiving Items from the Dropbox Library ($($dropBoxName))."
        $allDocuments = Get-AllItems -listName $dropBoxName

        if (!$allDocuments.Count)
        { Update-Status -Message $("No documents were found in the Dropbox Library.") -Level Warn; }
        else { Update-Status -Message $("Found $($allDocuments.Count) document(s) in the Dropbox Library."); }

        foreach ($document in $allDocuments)
        {
            $docName = $document['FileLeafRef'];
            $studentId = $document['StudentId'];
            $docType = $document['SSDocumentType'];
            if (!$studentId -or !$docType) { Update-Status -Message $("Skipping Document '$($docName)' because it has not been assigned a Student ID and/or Document Type."); continue; }
            #Document has been assigned a student ID, Process Intake Procedure.

            Update-Status -Message $("Processing Document: '$($docName)'.");
            $student = $studentData | Where-Object { $_.STUDENT_NUMBER -eq $document["StudentId"] | Select-Object -First 1 };
            if (!$student) { Update-Status -Message $("The Student ID, $($studentId), assigned to document '$($docName)' does not match Valid Student Data. The file has been skipped.") -level Error; continue; }

            if (!$config.SimulationMode)
            {
                Move-DocumentItem -itemToMove $document -DestinationLibraryNameOrURL $documentLibraryName -OverWriteIfExists $true
                #Get Moved Document and set the metadata
                $movedDocument = Get-DocumentByFileName -ListName $documentLibraryName -FileName $document["FileLeafRef"] ;
                if ($movedDocument) { Set-MetadataAndPermission -listName $documentLibraryName -Document $movedDocument -studentData $student}
            }
            else {Update-Status -Message "Simulation Mode: Would have moved file '$($document["FileLeafRef"])' to '$($documentLibraryName)'." }
        }
    }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
}
function Set-MetadataAndPermission
{
    Param (
        [Parameter(Mandatory = $true)]
        [String] $listName,
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.ListItem] $Document,
        [Parameter(Mandatory = $true)]
        [Object[]] $studentData
    )
    Try
    {

        $metadata = @{
            CaseManager      = Coalesce @($studentData.CASE_MANAGER) $null;
            Teachers         = Coalesce @($studentData.TEACHER_EMAIL.TrimEnd(';').Split(';')) $null;
            School           = Coalesce $studentData.SCHOOL_ABBREVIATION $null;
            StudentFirstName = Coalesce $studentData.FIRST_NAME $null;
            StudentLastName  = Coalesce $studentData.LAST_NAME $null;
            StudentWebId     = Coalesce $studentData.STUDENT_WEB_ID $null;
            GradeLevel       = Coalesce $studentData.GRADE_LEVEL $null;
            GraduationYear   = Coalesce $studentData.CLASSOF $null;
            HomeRoom         = Coalesce $studentData.HOME_ROOM $null;
            SchoolTeam       = Coalesce $studentData.TEAM $null;
        }

        Update-Status -Message "Setting Metadata on '$($Document["FileLeafRef"])'."

        If (!$config.SimulationMode)
        {
            try
            {
                #resolve multi-valued field issue
                $metadata = Resolve-MultiValuedFields -proposedMetadata $metadata -item $Document
                if ($Document.FieldValues.CheckoutUser -and $config.ForceCheckIn)
                {
                    Update-Status -Message "File '$($Document["FileLeafRef"])' is Locked for editing by $($Document.FieldValues.CheckoutUser.Email). Will try to force check-in." -Level Warn
                    Set-PNPFileCheckedIn -url $Document["FileRef"] -CheckinType OverwriteCheckIn -Comment "Checked in by Administrator."
                }

                Set-PnPListItem -List $listName -Identity $Document -Values $metadata -ErrorAction Stop |Out-Null
            }
            catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }

            #Set Permissions
            Reset-ItemPermission -listName $listName -item $Document
            #Set Permission for Case Manager
            if ($Document["CaseManager"])
            {
                Update-Status -Message "Granting Permission to Case Manager: $($Document["CaseManager"].LookupValue)"
                Grant-ItemPermission -listName $documentLibraryName -item $Document -userName $Document["CaseManager"].Email -role Contribute
            }
            #Set Permissions for teachers.
            Update-Status -Message "Granting Permission to $($Document["Teachers"].count) teachers."
            foreach ($teacher in $Document["Teachers"])
            {
                Grant-ItemPermission -listName $documentLibraryName -item $Document -userName $teacher.Email -role Read
            }
            #Set Permissions for School Team (if applicable)
            if ($Document["School"])
            {
                Update-Status -Message "Granting Permission to school Team: $("SPStudentServices-" + $Document["School"])."
                Grant-ItemPermission -listName $documentLibraryName -item $Document -userName $("SPStudentServices-" + $Document["School"]) -isGroup $true -role Contribute
            }
        }
        else
        {
            Update-Status -Message "Simulation Mode: Would have set the following Metadata."
            Update-Status -Message $metadata -Level Data;
        }
    }
    catch { Update-Status -Message "$($Error[0].Exception.Message) Stacktrace: $($Error[0].ScriptStackTrace)" -Level Error; }
}

>>>>>>> 77cac6119b1b583aaef6a5803e5f77a558c98e3f
#endregion IEP/504 Processing Helper Functions =============================