<#
Copyright Amazon.com, Inc. or its affiliates. All Rights Reserved.

Permission is hereby granted, free of charge, to any person obtaining a copy of this
software and associated documentation files (the "Software"), to deal in the Software
without restriction, including without limitation the rights to use, copy, modify,
merge, publish, distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

<#
.SYNOPSIS
    This script contains all of the logic for the EUC Toolkit GUI to operate.
.DESCRIPTION
    This script is the main script to initialize the EUC Toolkit GUI. Many actions are offloaded
    to the EUCToolkit-Helper module. For more information, see the link below:
    https://github.com/aws-samples/euc-toolkit
#>
Write-Host "Please wait while the EUC Toolkit Initializes"

Write-Host "Importing helper module"
$env:PSModulePath = "$env:PSModulePath;$($PSScriptRoot+"\Assets\EUCToolkit-Helper.psm1")"
Import-Module -Name $($PSScriptRoot+"\Assets\EUCToolkit-Helper.psm1") -Force
add-type -AssemblyName System.Windows.Forms
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')

# XML Loader for GUI
[xml]$WksMainXaml = Get-Content $($PSScriptRoot+"\Assets\EUCToolkit-MainGUI.xml")
# Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $WksMainXaml) 
try{$WksMainForm=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader"; exit}

# Store Form Objects In PowerShell
$WksMainXaml.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name ($_.Name) -Value $WksMainForm.FindName($_.Name)}

# Checks credentials. If manually set, it will continue. If not, it will check for an instance profile
# using a IMDSv2 token. For more information, see the link below:
# https://docs.aws.amazon.com/AWSEC2/latest/UserGuide/configuring-instance-metadata-service.html
if(Get-AWSCredential){
    $lblPermissions.Content = "Profile Manually Set"
    $global:InstanceProfile = $false
}elseif(Test-Path env:AWS_ACCESS_KEY_ID) {
    $lblPermissions.Content = "Utilizing Environment Variables"
    $global:InstanceProfile = $false										  
}else{
    $global:InstanceProfile = $true
    try{
        # IMDSv2 Method 
        [string]$token = Invoke-RestMethod -Headers @{"X-aws-ec2-metadata-token-ttl-seconds" = "21600"} -Method PUT -Uri http://169.254.169.254/latest/api/token
        $TestInstanceProfile = Invoke-RestMethod -Headers @{"X-aws-ec2-metadata-token" = $token} -Method GET -Uri http://169.254.169.254/latest/meta-data/iam/info
    }catch{
        $global:InstanceProfile = $false
        Show-MessageError -message "No set credentials or Instance Profile detected." -title "Error Retrieving Credentials"
    }
    if($InstanceProfile -eq $true){
        $lblPermissions.Content = "Utilizing Instance Profile"
    }else{
        $lblPermissions.Content = "Undefined Profile (default)"
    }
}
 
#############################
# ! # ! # FUNCTIONS # ! # ! # 
#############################

# This function updates the PowerShell object that acts as the local db and the counter
function Update-WorkSpaceObject(){
    write-host "Getting a list of regions and directories where WorkSpaces are deployed"
    $global:WorkSpacesDirectoryDB = Get-WksDirectories
    write-host "Getting WorkSpaces Service Quotas"
    $global:WorkSpacesServiceQuotaDB = Get-WksServiceQuotasDB -DeployedRegions $global:WorkSpacesDirectoryDB
    write-host "Getting a list of WorkSpaces deployed"
    $global:WorkSpacesDB = Get-LocalWorkSpacesDB -DeployedRegions $global:WorkSpacesDirectoryDB
    $date = (get-date -Format "MM/dd/yyyy HH:mm") | Out-String
    $lblLastDBUpdateBulk.Content = $date
    $lblLastDBUpdate.Content = $date
    Update-Counter
    write-host "WorkSpaces queries are complete"
}

# This function filters the local db object in real time so that your search criteria is immediately applied
function Search-WorkSpaces(){
    $SearchResults.Items.Clear()
    $filtered = $global:WorkSpacesDB
    if($FirstName.Text -ne ""){
        $filtered = $filtered | Where-Object { ($_.FirstName -like ($FirstName.Text + "*")) } | Select-Object WorkSpaceId,Region,UserName,FirstName,LastName,ComputerName,Email,Protocol
    }
    if($LastName.Text -ne ""){
        $filtered = $filtered | Where-Object { ($_.LastName -like ($LastName.Text + "*")) } | Select-Object WorkSpaceId,Region,UserName,FirstName,LastName,ComputerName,Email,Protocol
    }
    if($Email.Text -ne ""){
        $filtered = $filtered | Where-Object { ($_.Email -like ("*"+$Email.Text + "*")) } | Select-Object WorkSpaceId,Region,UserName,FirstName,LastName,ComputerName,Email,Protocol 
    }
    if($txtComputerName.Text -ne ""){
        $filtered = $filtered | Where-Object { ($_.ComputerName -like ("*" + $txtComputerName.Text + "*")) } | Select-Object WorkSpaceId,Region,UserName,FirstName,LastName,ComputerName,Email,Protocol
    }
    if($txtUserName.Text -ne ""){
        $filtered = $filtered | Where-Object { ($_.UserName -like ("*" + $txtUserName.Text + "*")) } | Select-Object WorkSpaceId,Region,UserName,FirstName,LastName,ComputerName,Email,Protocol
    }
    if($cmboProtocol.SelectedItem -ne "All"){
        $filtered = $filtered | Where-Object { ($_.Protocol -like ($cmboProtocol.SelectedItem)) } | Select-Object WorkSpaceId,Region,UserName,FirstName,LastName,ComputerName,Email,Protocol
    }

    foreach($workspace in $filtered){
        if($NULL -ne $workspace.WorkSpaceId){
            $SilenceOutput=$SearchResults.items.Add($workspace)
        }
    }
}

# This function finds all of your registered WorkSpaces directories 
function Get-BulkDirectories(){
    $selectWKSDirectory.Items.Clear()
    $selectWKSDirectory.Items.Add("All Directories")
    $Directories = $global:WorkSpacesDB | Select-Object directoryId, Region -Unique | Where-Object { ($_.Region -eq $selectWKSRegion.SelectedItem)}
    foreach ($Directory in $Directories ){
        $directorySTR=$Directory.directoryId
        $selectWKSDirectory.Items.Add($directorySTR)
    }
    $selectWKSDirectory.SelectedIndex=0
}

# This function filters the local db object to reflect your search criteria in the bulk tab
function Get-ImpactedWS(){
    $lstImpactedWorkSpaces.items.Clear()
    $bulkdirectoryId = $selectWKSDirectory.SelectedItem
    if($NULL -eq $selectWKSBundle.SelectedItem){
        $selectWKSBundle.items.add("Select Bundle")
        $selectWKSBundle.SelectedIndex=0
    }
    $allImpactedWS = $global:WorkSpacesDB | Where-Object { ($_.Region -eq $selectWKSRegion.SelectedItem)}
    if($bulkdirectoryId -like "All Directories"){
        $allImpactedWS = $allImpactedWS | Select-Object directoryId, WorkSpaceId,UserName,Region,FirstName,LastName,ComputerName,Email,RunningMode,State,Protocol,BundleId
    }else{
        $allImpactedWS = $allImpactedWS | Where-Object { ($_.directoryId -eq $bulkdirectoryId) }| Select-Object directoryId, WorkSpaceId,UserName,Region,FirstName,LastName,ComputerName,Email,RunningMode,State,Protocol,BundleId
    }
    if($selectRunningModeFilterCombo.SelectedItem -ne "Select Running Mode"){
        $allImpactedWS = $allImpactedWS | Where-Object { ($_.RunningMode -like ($selectRunningModeFilterCombo.SelectedItem))} | Select-Object directoryId, WorkSpaceId,UserName,Region,FirstName,LastName,ComputerName,Email,RunningMode,State,Protocol,BundleId
    }
    if($selectWKSBundle.SelectedItem -ne "Select Bundle"){
        $bundleSTR = ($selectWKSBundle.SelectedItem.split(" "))[0]
        $allImpactedWS = $allImpactedWS | Where-Object { ($_.BundleId -like ($bundleSTR))} | Select-Object directoryId, WorkSpaceId,UserName,Region,FirstName,LastName,ComputerName,Email,RunningMode,State,Protocol,BundleId
    }
    if($cmboBulkProtocol.SelectedItem -ne "All"){
        $protocolSTR = ($cmboBulkProtocol.SelectedItem.split(" "))[0]
        $allImpactedWS = $allImpactedWS | Where-Object { ($_.Protocol -like ($protocolSTR))} | Select-Object directoryId, WorkSpaceId,UserName,Region,FirstName,LastName,ComputerName,Email,RunningMode,State,Protocol,BundleId
    }
    foreach ($impactedWS in $allImpactedWS){
        if($NULL -ne $impactedWS.WorkSpaceId -and $lstImpactedWorkSpaces.Items.WorkSpaceId -notcontains $impactedWS.WorkSpaceId){
            $lstImpactedWorkSpaces.Items.Add($impactedWS)
        }
    }
}

# This function logs your action in the logging tab (does not persist GUI sessions)
function Write-Logger(){
    param(
        [String]$message
    )
   $logEntry = New-Object -TypeName PSobject
   $logDate = Get-Date -Format "MM/dd/yyyy HH:mm K"
   $logEntry | Add-Member -NotePropertyName "loggingTime" -NotePropertyValue $logDate
   $logEntry | Add-Member -NotePropertyName "loggingMessage" -NotePropertyValue $message
   $SilenceOutput=$lstLogging.items.Add($logEntry)
}

# This function counts and shows your current WorkSpaces counts (Total, Available, and Stopped)
function Update-Counter(){
    $total = $global:WorkSpacesDB.WorkSpaceId.Count
    $available = ($global:WorkSpacesDB | Where-Object { $_.State -eq "AVAILABLE"}).Count
    $stopped = ($global:WorkSpacesDB | Where-Object { $_.State -eq "STOPPED"}).Count
    $PCoIP = ($global:WorkSpacesDB | Where-Object { $_.Protocol -eq "PCOIP"}).Count
    $WSP = ($global:WorkSpacesDB | Where-Object { $_.Protocol -eq "WSP"}).Count
    $BYOP = ($global:WorkSpacesDB | Where-Object { $_.Protocol -eq "BYOP"}).Count
    if($total -eq 0 -or $NULL -eq $total){
        $total = 0
    }
    if($available -eq 0 -or $NULL -eq $available){
        $available = 0
    }
    if($stopped -eq 0 -or $NULL -eq $stopped){
        $stopped = 0
    }
    if($PCoIP -eq 0 -or $NULL -eq $PCoIP){
        $PCoIP = 0
    }
    if($WSP -eq 0 -or $NULL -eq $WSP){
        $WSP = 0
    }
    if($BYOP -eq 0 -or $NULL -eq $BYOP){
        $BYOP = 0
    }
    $TotalWorkSpacesCount.content = $total
    $TotalAvailable_Count.content = $available
    $TotalStopped_Count.content = $stopped
    $lblBulkPCOIPCounter.content = $PCoIP
    $lblBulkWSPCounter.content = $WSP
    $lblBulkBYOPCounter.content = $BYOP
}

###############################################
# ! # ! # WorkSpaces Main GUI Actions # ! # ! # 
###############################################
# This section's actions correspond with a GUI button action. The button objects below are 
# created from objects outlined within the XML.

$FirstName.Add_TextChanged({
    Search-WorkSpaces
})
$LastName.Add_TextChanged({
    Search-WorkSpaces
})
$Email.Add_TextChanged({
    Search-WorkSpaces
})
$txtComputerName.Add_TextChanged({
    Search-WorkSpaces
})
$txtUserName.Add_TextChanged({
    Search-WorkSpaces
})

$btnUpdateData.Add_Click({
	$btnUpdateData.Content="Running..."
    $btnUpdateData.IsEnabled=$false
								   
    Update-WorkSpaceObject
    Search-WorkSpaces

    $btnUpdateData.content="Refresh"
    $btnUpdateData.IsEnabled=$true
})

$cmboProtocol.add_SelectionChanged({
    Search-WorkSpaces
})

$btnRemoteAssist.Add_Click({
    Invoke-RemoteAssist -privateIP $IPValue.Content
    Write-Logger -message "Remote Assist Initiated on " + $IPValue.Content
})

$btnUpdateComputeType.Add_Click({
    $WorkSpaceId = $WorkSpaceIdValue.Content
    $UpdateComputeReq = New-Object -TypeName PSobject
    $UpdateComputeReq | Add-Member -NotePropertyName "WorkSpaceId" -NotePropertyValue $WorkSpaceId
    $UpdateComputeReq | Add-Member -NotePropertyName "CurrentCompute" -NotePropertyValue ($global:WorkSpacesDB | Where-Object {$_.WorkspaceId -eq $WorkSpaceId}).Compute
    $UpdateComputeReq | Add-Member -NotePropertyName "TargetCompute" -NotePropertyValue $cmboComputeValue.SelectedItem
    $UpdateComputeReq | Add-Member -NotePropertyName "Region" -NotePropertyValue $RegionValue.Content 
    $Output = Update-ComputeType -ComputeReq $UpdateComputeReq
    if($Output.ErrorCode){
        $ErrCode = $Output.ErrorCode.ToString()
        $ErrMsg = $Output.ErrorMessage.ToString()
        Write-Logger -message "Compute change failed for WorkSpaceId $WorkSpaceId Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error during Compute change, see log tab for details" -title "Error Changing Compute Type"
    }else{
        Show-MessageSuccess -message "Compute Type update on $WorkSpaceId executed successfully" -title "Successfully Updated Compute Type"
        Write-Logger -message "Compute Type update on $WorkSpaceId executed successfully"
    }
})

$btnUpdateRootVolume.Add_Click({
    $WorkSpaceId = $WorkSpaceIdValue.Content
    $UpdateRootVolReq = New-Object -TypeName PSobject
    $UpdateRootVolReq | Add-Member -NotePropertyName "WorkSpaceId" -NotePropertyValue $WorkSpaceId
    $UpdateRootVolReq | Add-Member -NotePropertyName "CurrentRootStorage" -NotePropertyValue $RootValue.Content
    $UpdateRootVolReq | Add-Member -NotePropertyName "CurrentUserStorage" -NotePropertyValue $UserValue.Content
    $UpdateRootVolReq | Add-Member -NotePropertyName "Region" -NotePropertyValue $RegionValue.Content 

    $Output = Update-RootVolume -WorkSpaceReq $UpdateRootVolReq
    if($Output.ErrorCode){
        $ErrCode = $Output.ErrorCode.ToString()
        $ErrMsg = $Output.ErrorMessage.ToString()
        Write-Logger -message "Root volume extention failed for WorkSpaceId $WorkSpaceId Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error during Root volume extention, see log tab for details" -title "Error Extending Root Volume"
    }else{
        Show-MessageSuccess -message "Root volume extention on $WorkSpaceId executed successfully" -title "Successfully Extended Root Volume"
        Write-Logger -message "Root volume extention on $WorkSpaceId executed successfully"
    }
})

$btnUpdateUserVolume.Add_Click({
    $WorkSpaceId = $WorkSpaceIdValue.Content
    $UpdateUserVolReq = New-Object -TypeName PSobject
    $UpdateUserVolReq | Add-Member -NotePropertyName "WorkSpaceId" -NotePropertyValue $WorkSpaceId
    $UpdateUserVolReq | Add-Member -NotePropertyName "CurrentRootStorage" -NotePropertyValue $RootValue.Content
    $UpdateUserVolReq | Add-Member -NotePropertyName "CurrentUserStorage" -NotePropertyValue $UserValue.Content
    $UpdateUserVolReq | Add-Member -NotePropertyName "Region" -NotePropertyValue $RegionValue.Content 

    $Output = Update-UserVolume -WorkSpaceReq $UpdateUserVolReq
    if($Output.ErrorCode){
        $ErrCode = $Output.ErrorCode.ToString()
        $ErrMsg = $Output.ErrorMessage.ToString()
        Write-Logger -message "User volume extention failed for WorkSpaceId $WorkSpaceId Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error during User volume extention, see log tab for details" -title "Error Extending User Volume"
    }else{
        Show-MessageSuccess -message "User volume extention on $WorkSpaceId executed successfully" -title "Successfully Extended User Volume"
        Write-Logger -message "User volume extention on $WorkSpaceId executed successfully"
    }
})

$btnTerminateWS.Add_Click({
    $WorkSpaceId = $WorkSpaceIdValue.Content
    $wshell = New-Object -ComObject Wscript.Shell
    $response = $wshell.Popup("Are you sure you would like to terminate $WorkSpaceId? This cannot be undone.",0,"Alert",64+4)
    if($response -eq 6){
        Write-Logger -message "Executing Terminate on WorkSpaceId $WorkSpaceId"
        $filteredWS = @()
        $filteredWS += $global:WorkSpacesDB | Where-Object { ($_.WorkSpaceId -like $WorkSpaceId) } | Select-Object WorkSpaceId,Region
        $TermReq = New-Object -TypeName PSobject
        $TermReq | Add-Member -NotePropertyName "TermList" -NotePropertyValue $filteredWS
        $Output = Optimize-APIRequest -requestInfo $TermReq -APICall "Remove-WKSWorkspace"
        if($Output.ErrorCode){
            $ErrCode = $Output.ErrorCode.ToString()
            $ErrMsg = $Output.ErrorMessage.ToString()
            Write-Logger -message "Terminate failed for WorkSpaceId $WorkSpaceId Details below:"
            Write-Logger -message "Error Code: $ErrCode"
            Write-Logger -message "Error Message: $ErrMsg"
            Show-MessageError -message "Error Terminating, see log tab for details" -title "Error Terminating WorkSpace"
        }else{
            Show-MessageSuccess -message "Terminate API on $WorkSpaceId executed successfully" -title "Successfully Terminated WorkSpace"
            Write-Logger -message "Terminate API on $WorkSpaceId executed successfully"
        }
    }
})

$btnRebuildWS.Add_Click({
    $WorkSpaceId = $WorkSpaceIdValue.Content
    Write-Logger -message "Executing Rebuild on WorkSpaceId $WorkSpaceId"
    $ResetReq = New-Object -TypeName PSobject
    $ResetReq | Add-Member -NotePropertyName "WorkSpaceId" -NotePropertyValue $WorkSpaceId
    $ResetReq | Add-Member -NotePropertyName "Region" -NotePropertyValue $RegionValue.Content
    $Output = Optimize-APIRequest -requestInfo $ResetReq -APICall "Reset-WKSWorkspace"
    if($Output.ErrorCode){
        $ErrCode = $Output.ErrorCode.ToString()
        $ErrMsg = $Output.ErrorMessage.ToString()
        Write-Logger -message "Rebuild failed for WorkSpaceId $WorkSpaceId Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error Rebuilding, see log tab for details" -title "Error Rebuilding WorkSpace"
    }else{
        Show-MessageSuccess -message "Rebuild API on $WorkSpaceId executed successfully" -title "Successfully Rebuilt WorkSpace"
        Write-Logger -message "Rebuild API on $WorkSpaceId executed successfully"
    }
})

$btnRestoreWS.Add_Click({
    $WorkSpaceId = $WorkSpaceIdValue.Content
    Write-Logger -message "Executing Restore on WorkSpaceId $WorkSpaceId"
    $ResetReq = New-Object -TypeName PSobject
    $ResetReq | Add-Member -NotePropertyName "WorkSpaceId" -NotePropertyValue $WorkSpaceId
    $ResetReq | Add-Member -NotePropertyName "Region" -NotePropertyValue $RegionValue.Content
    $Output = Optimize-APIRequest -requestInfo $ResetReq -APICall "Restore-WKSWorkspace"
    if($Output.ErrorCode){
        $ErrCode = $Output.ErrorCode.ToString()
        $ErrMsg = $Output.ErrorMessage.ToString()
        Write-Logger -message "Restore failed for WorkSpaceId $WorkSpaceId Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error Restoring, see log tab for details" -title "Error Restoring WorkSpace"
    }else{
        Show-MessageSuccess -message "Restore API on $WorkSpaceId executed successfully" -title "Successfully Restored WorkSpace"
        Write-Logger -message "Restore API on $WorkSpaceId executed successfully"
    }
})

$btnPowerUpWS.Add_Click({
    $WorkSpaceId =$WorkSpaceIdValue.Content
    Write-Logger -message "Executing Start on WorkSpaceId $WorkSpaceId"
    $StartReq = New-Object -TypeName PSobject
    $StartReq | Add-Member -NotePropertyName "WorkSpaceId" -NotePropertyValue $WorkSpaceId
    $StartReq | Add-Member -NotePropertyName "Region" -NotePropertyValue $RegionValue.Content
    $Output = Optimize-APIRequest -requestInfo $StartReq -APICall "Start-WKSWorkspace"
    if($Output.ErrorCode){
        $ErrCode = $Output.ErrorCode.ToString()
        $ErrMsg = $Output.ErrorMessage.ToString()
        Write-Logger -message "Start failed for WorkSpaceId $WorkSpaceId Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error starting, see log tab for details" -title "Error Starting WorkSpace"
    }else{
        Show-MessageSuccess -message "Start API on $WorkSpaceId executed successfully" -title "Successfully Started WorkSpace"
        Write-Logger -message "Start API on $WorkSpaceId executed successfully"
    }
})

$btnPowerDownWS.Add_Click({
    $WorkSpaceId =$WorkSpaceIdValue.Content
    Write-Logger -message "Executing Stop on WorkSpaceId $WorkSpaceId"
    $StopReq = New-Object -TypeName PSobject
    $StopReq | Add-Member -NotePropertyName "WorkSpaceId" -NotePropertyValue $WorkSpaceId
    $StopReq | Add-Member -NotePropertyName "Region" -NotePropertyValue $RegionValue.Content
    $Output = Optimize-APIRequest -requestInfo $StopReq -APICall "Stop-WKSWorkspace"
    if($Output.ErrorCode){
        $ErrCode = $Output.ErrorCode.ToString()
        $ErrMsg = $Output.ErrorMessage.ToString()
        Write-Logger -message "Stop failed for WorkSpaceId $WorkSpaceId Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error stopping, see log tab for details" -title "Error Stopping WorkSpace"
    }else{
        Show-MessageSuccess -message "Stop API on $WorkSpaceId executed successfully" -title "Successfully Stopped WorkSpace"
        Write-Logger -message "Stop API on $WorkSpaceId executed successfully"
    }
})

$btnRebootWS.Add_Click({
    $WorkSpaceId =$WorkSpaceIdValue.Content
    Write-Logger -message "Executing Reboot on WorkSpaceId $WorkSpaceId"
    $StopReq = New-Object -TypeName PSobject
    $StopReq | Add-Member -NotePropertyName "WorkSpaceId" -NotePropertyValue $WorkSpaceId
    $StopReq | Add-Member -NotePropertyName "Region" -NotePropertyValue $RegionValue.Content
    $Output = Optimize-APIRequest -requestInfo $StopReq -APICall "Restart-WKSWorkspace"
    if($Output.ErrorCode){
        $ErrCode = $Output.ErrorCode.ToString()
        $ErrMsg = $Output.ErrorMessage.ToString()
        Write-Logger -message "Reboot failed for WorkSpaceId $WorkSpaceId Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error rebooting, see log tab for details" -title "Error Rebooting WorkSpace"
    }else{
        Show-MessageSuccess -message "Reboot API on $WorkSpaceId executed successfully" -title "Successfully Rebooted WorkSpace"
        Write-Logger -message "Reboot API on $WorkSpaceId executed successfully"
    }
})

$btnChangeRunningMode.Add_Click({
    $WorkSpaceId = $WorkSpaceIdValue.Content
    $RunningModeUpdate = $RunningModeValue.Content
    $UpdateRunningReq = New-Object -TypeName PSobject
    $UpdateRunningReq | Add-Member -NotePropertyName "WorkSpaceId" -NotePropertyValue $WorkSpaceId
    $UpdateRunningReq | Add-Member -NotePropertyName "CurrentRunMode" -NotePropertyValue $RunningModeUpdate
    $UpdateRunningReq | Add-Member -NotePropertyName "Region" -NotePropertyValue $RegionValue.Content 
    $Output = Update-RunningMode -UpdateReq $UpdateRunningReq
    if($RunningModeUpdate -eq "AUTO_STOP"){
        $TargetRun = "ALWAYS_ON"
    }else{
        $TargetRun = "AUTO_STOP"
    }
    if($Output.ErrorCode){
        $ErrCode = $Output.ErrorCode.ToString()
        $ErrMsg = $Output.ErrorMessage.ToString()
        Write-Logger -message "Running Mode change failed for WorkSpaceId $WorkSpaceId. Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error during Running Mode change, see log tab for details" -title "Error Changing Running Mode"
    }else{
        Show-MessageSuccess -message "Running Mode update on $WorkSpaceId executed successfully" -title "Successfully Updated Running Mode"
        Write-Logger -message "Running Mode update to $TargetRun on $WorkSpaceId executed successfully"
    }
})

$btnCopyWSId.Add_Click({
    write-output $WorkSpaceIdValue.Content | Set-clipboard
})

$btnCopyWSIP.Add_Click({
    write-output $IPValue.Content | Set-clipboard
})

$btnCopyWSComputerName.Add_Click({
    write-output $ComputerNameValue.Content | Set-clipboard
})

$btnBackupUserVolume.Add_Click({
    $WorkspaceId = $WorkSpaceIdValue.Content
    $tmpPath = $txtDisk2VHDPath.Text + "Disk2VHD64.exe"
    $filePathDisk2VHD = Convert-Path -path $tmpPath
    $filePathPSExec = $txtPSExecPath.Text
    $backupDestination = $txtBackUpDest.Text + $WorkspaceId + ".vhdx"
    $targetIP = $IPValue.Content
    if($null -ne $filePathPSExec){
        if($null -ne $txtDisk2VHDPath.Text){
            $tmpFileName = $WorkspaceId+".vhdx"
            $arguments = " -s -c \\$targetIP `"$filePathDisk2VHD`" -C D: D:\$tmpFileName -accepteula"
            $exe = $filePathPSExec+"psexec.exe"
            try{
                start-process $exe $arguments -Wait
                $WorkSpaceDisk2VHDLocation = '\\'+$targetIP+'\d$\'+$WorkspaceId+'.VHDX'
                Copy-Item -Path $WorkSpaceDisk2VHDLocation -Destination $backupDestination
            }Catch{
                $msg = $_
                Write-Logger -message "Failed to backup $WorkspaceId. See error details below:"
                Write-Logger -message "$msg"
            }
        }else{
            Show-MessageError -message "Unable to backup WorkSpace, Disk2VHD path not provided in Admin tab." -title "Disk2VHD Backup Failed"
        }
        
    }else{
        Show-MessageError -message "Unable to backup WorkSpace, PSExec path not provided in Admin tab." -title "PSExec Backup Failed"
    }
})
    
$btnGatherLogs.Add_Click({
    $WorkspaceId = $WorkSpaceIdValue.Content
    $filePathPSExec = $txtPSExecPath.Text
    $targetIP = $IPValue.Content
    if($null -ne $filePathPSExec){
        $arguments = ' -i -s \\'+$targetIP+' PowerShell.exe -NoLogo -ExecutionPolicy RemoteSigned -NoProfile -File "C:\Program Files\Amazon\WorkSpacesConfig\Scripts\Get-WorkSpaceLogs.ps1"'
        $exe = $filePathPSExec+"psexec.exe"
        try{
            start-process $exe $arguments -Wait
            $WorkSpaceLogsLocation = '\\'+$targetIP+'\d$\*.zip'
            $filePathLogsOutputh = $folderPathLogsOutput + $WorkSpaceId + "\"
            if (!(Test-Path $filePathLogsOutputh)){
                New-Item -Path $filePathLogsOutputh  -ItemType Directory
            }
            Copy-Item -Path $WorkSpaceLogsLocation -Destination $filePathLogsOutputh
        }Catch{
            $msg = $_
            Write-Logger -message "Failed to gather logs for $WorkspaceId. See error details below:"
            Write-Logger -message "$msg"
        }
    }else{
        Show-MessageError -message "Unable to gather WorkSpace logs, PSExec path not provided in Admin tab." -title "PSExec Log Gather Failed"
    }
})

$btnRDP.Add_Click({
    $WorkSpaceLookup = $global:WorkSpacesDB | Where-Object { ($_.WorkSpaceId -like $WorkSpaceIdValue.Content) } 
    if($WorkSpaceLookup.ComputerName[0] -ne "A"){
        if($WorkSpaceLookup.State -eq 'STOPPED'){
            $wshell = New-Object -ComObject Wscript.Shell
            $response=$wshell.Popup("Warning, WorkSpace is powered off, Would you like it powered on to connect?",0,"",0x1)
        }else{
            $connectionInfo = $WorkSpaceLookup.ComputerName 
            mstsc /v:$connectionInfo
            Write-Logger -message "RDP session Initiated on " + $IPValue.Content
        }
    }else{
        Show-MessageError -message "Unable to RDP to a Linux WorkSpace" -title "Linux WorkSpace Selected"
    }
})

# Button to Modify Protocol (PCoIP to WSP)
$btnModifyProtocol.Add_Click({
    $WorkSpaceId = $WorkSpaceIdValue.Content
    $ProtocolModifyReq = New-Object -TypeName PSobject
    $ProtocolModifyReq | Add-Member -NotePropertyName "WorkSpaceId" -NotePropertyValue $WorkSpaceId
    $ProtocolModifyReq | Add-Member -NotePropertyName "Region" -NotePropertyValue $RegionValue.Content
    $ProtocolModifyReq | Add-Member -NotePropertyName "Protocol" -NotePropertyValue $lblProtocolValue.Content
    $Output = Set-WKSProtocol -ProtocolModifyReq $ProtocolModifyReq
    if($Output.ErrorCode){
        $ErrCode = $Output.ErrorCode
        $ErrMsg = $Output.ErrorMessage
        Write-Logger -message "Protocol modification failed $WorkSpaceId. Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error modifying protocol, see log tab for details" -title "Error Modifying Protocol"
    }else{
        Show-MessageSuccess -message "Protocol modification on $WorkSpaceId executed successfully" -title "Successfully Modified Protocol"
        Write-Logger -message "Protocol modification on $WorkSpaceId executed successfully"
    }

})

$btnGetUserExperience.Add_Click({
    $workingDirectory=($PSScriptRoot +"\\Assets\\CWHelper\\") 
    $global:cloudWatchRun++

    # Clear out last run
    $CloudWatchHistoricalLatency.Source =($PSScriptRoot+"\\Assets\\CWHelper\\WorkSpacesHistoricalLatency-Start.png") 
    $CloudWatchWorkSpaceLatency.Source =($PSScriptRoot+"\\Assets\\CWHelper\\WorkSpacesUDPPacketLoss-Start.png")
    $CloudWatchWorkSpaceLaunch.Source =($PSScriptRoot+"\\Assets\\CWHelper\\WorkSpacesSessionLaunch-Start.png") 
    $CloudWatchWorkSpaceCPU.Source =($PSScriptRoot+"\\Assets\\CWHelper\\WorkSpacesCPU-Start.png") 
    $CloudWatchWorkSpaceMemory.Source =($PSScriptRoot+"\\Assets\\CWHelper\\WorkSpacesMemory-Start.png")
    $CloudWatchWorkSpaceDisk.Source =($PSScriptRoot+"\\Assets\\CWHelper\\WorkSpacesDisk-Start.png")
    $CloudWatchWorkSpaceCPU.Visibility ="Visible"
    $CloudWatchWorkSpaceMemory.Visibility ="Visible"
    $CloudWatchWorkSpaceDisk.Visibility ="Visible"
    $CloudWatchLoginInfo.items.clear()
    $CloudWatchWorkSpaceModifications.items.clear()

    $resources = Get-ChildItem -Path "$workingDirectory\SelectedWSMetrics\"
    if($resources.count -ne 0-and($global:cloudWatchRun -eq 1)){
        foreach($item in $resources){
            $tmpFileName = $item.Name
            Remove-item -path "$workingDirectory\SelectedWSMetrics\$tmpFileName" -force -ErrorAction SilentlyContinue
        }
    }

    # Move to CloudWatch Tab
    $tabCloudWatch.Visibility ="Visible"
    $tabControl.SelectedIndex =5
    
    $CloudWatch_Tick={ 
        $imageRepo = "$PSScriptRoot\Assets\CWHelper\SelectedWSMetrics"
        $resources = Get-ChildItem -Path $imageRepo
        if($resources.count -eq $global:ImagesExpected){
            start-sleep -Seconds 1
            $CloudWatchHistoricalLatency.Source = ($imageRepo+"\WorkSpacesHistoricalLatency"+$global:cloudWatchRun+".png")
            $CloudWatchWorkSpaceLatency.Source = ($imageRepo+"\WorkSpacesUDP"+$global:cloudWatchRun+".png")
            $CloudWatchWorkSpaceLaunch.Source = ($imageRepo+"\WorkSpacesConnectionSummary"+$global:cloudWatchRun+".png")
            $CloudWatchWorkSpaceCPU.Source = ($imageRepo+"\WorkSpacesCPU"+$global:cloudWatchRun+".png") 
            $CloudWatchWorkSpaceCPU.Visibility ="Visible"
            $CloudWatchWorkSpaceMemory.Source = ($imageRepo+"\WorkSpacesMemory"+$global:cloudWatchRun+".png") 
            $CloudWatchWorkSpaceMemory.Visibility ="Visible"
            $CloudWatchWorkSpaceDisk.Source = ($imageRepo+"\WorkSpacesDisk"+$global:cloudWatchRun+".png") 
            $CloudWatchWorkSpaceDisk.Visibility ="Visible"

            $global:CloudWatchTimer.Stop()
            $global:CloudWatchTimer.Dispose()
        }

        if(($CloudWatchLoginInfo.items.Count -eq 0) -and($resources | Where-Object {$_.Name -like "ConnectionHistory"+$global:cloudWatchRun+".csv"})){
            $global:imagesRetrieved++
            $CSV = Import-Csv ($imageRepo+"\ConnectionHistory"+$global:cloudWatchRun+".csv")
            foreach ($record in $CSV){
                $CloudWatchLoginInfo.items.Add($record)
            }   
        }
        if(($CloudWatchWorkSpaceModifications.items.Count -eq 0) -and($resources | Where-Object {$_.Name -like "UserChanges"+$global:cloudWatchRun+".csv"})){
            $global:imagesRetrieved++
            $CSV = Import-Csv ($imageRepo+"\UserChanges"+$global:cloudWatchRun+".csv")
            foreach ($record in $CSV){
                $CloudWatchWorkSpaceModifications.items.Add($record)
            }   
        }
    }    

    # Timer to Watch the Jobs:
    $global:CloudWatchTimer = New-Object 'System.Windows.Forms.Timer' 
    $global:QueryTime=1

    $imageRepo = "$PSScriptRoot\Assets\CWHelper\SelectedWSMetrics"
    $resources = Get-ChildItem -Path $imageRepo
    $global:ImagesExpected=(3+$resources.count)

    if($txtCloudWatchAccessLogs.Text-ne ""){
        $AccessLogsRegion=($txtCloudWatchAccessLogs.Text-split ":")[3]
        $AccessLogsName=($txtCloudWatchAccessLogs.Text-split ":")[6]
        $global:ImagesExpected+= 2													   
								
    }

    if($global:InstanceProfile -eq $true){
        Get-CloudWatchImagesServiceMetrics -workingDirectory $workingDirectory -WSId $WorkSpaceIdValue.content -AWSProfile $false -region $RegionValue.Content -CWRun $global:cloudWatchRun
										
        Get-CloudWatchImagesWorkSpaceMetrics -workingDirectory $workingDirectory -WSId $WorkSpaceIdValue.content -region $RegionValue.Content -CWRun $global:cloudWatchRun
        $global:ImagesExpected += 3 
		 
        Get-CloudWatchStats -workingDirectory $workingDirectory -WorkSpaceAccessLogs $AccessLogsName -CloudTrailLogs $CloudTrailsLogsName -WorkSpaceId $WorkSpaceIdValue.content -AWSProfile $false -CloudTrailRegion $CloudTrailsLogsRegion -AccessLogsRegion $AccessLogsRegion -CWRun $global:cloudWatchRun
    }else{
        $CurrentCred = Get-AWSCredential
        Get-CloudWatchImagesServiceMetrics -workingDirectory $workingDirectory -WSId $WorkSpaceIdValue.content -AWSProfile $CurrentCred -region $RegionValue.Content -CWRun $global:cloudWatchRun
										
        Get-CloudWatchImagesWorkSpaceMetrics -workingDirectory $workingDirectory -WSId $WorkSpaceIdValue.content -AWSProfile $CurrentCred -region $RegionValue.Content -CWRun $global:cloudWatchRun
        $global:ImagesExpected += 3
		 
        Get-CloudWatchStats -workingDirectory $workingDirectory -WorkSpaceAccessLogs $AccessLogsName -CloudTrailLogs $CloudTrailsLogsName -WorkSpaceId $WorkSpaceIdValue.content -CloudTrailRegion $CloudTrailsLogsRegion -AccessLogsRegion $AccessLogsRegion -CWRun $global:cloudWatchRun
    }

    $global:CloudWatchTimer.Enabled = $True 
    $global:CloudWatchTimer.Interval = 1000 
    $global:CloudWatchTimer.add_Tick($CloudWatch_Tick) 
    $global:imagesRetrieved = 0
})

# Updates interactive WorkSpaces Help Desk tab
$SearchResults.add_SelectionChanged({
    if ($SearchResults.SelectedIndex -ne -1){   
        $tabCloudWatch.Visibility ="Hidden"
        $searchedItem = $SearchResults.SelectedItems
        $wsID = $searchedItem.WorkSpaceId
        $WorkSpaceInfo = $global:WorkSpacesDB | Where-Object { ($_.WorkSpaceId -like $wsID) } 
        $UserNameValue.Content = $WorkSpaceInfo.UserName
        $WorkSpaceIdValue.Content = $WorkSpaceInfo.WorkspaceId
        $WorkSpaceComputeType = $WorkSpaceInfo.Compute.Value | Out-String
        $WorkSpaceComputeType = $WorkSpaceComputeType.replace("`n","").replace("`r","")

        if($WorkSpaceComputeType.substring(0,1) -ne "G"){
            $cmboComputeValue.Visibility="Visible"
            $lblComputeValue.Visibility="Hidden"
            $btnUpdateComputeType.Visibility="Visible"
            $cmboComputeValue.items.clear()
            $cmboComputeValue.items.Add("VALUE")
            $cmboComputeValue.items.Add("STANDARD")
            $cmboComputeValue.items.Add("PERFORMANCE")
            $cmboComputeValue.items.Add("POWER")
            $cmboComputeValue.items.Add("POWERPRO")
            $cmboComputeValue.SelectedItem = $WorkSpaceComputeType
        }else{
            $cmboComputeValue.Visibility="Hidden"
            $lblComputeValue.Visibility="Visible"
            $btnUpdateComputeType.Visibility="Hidden"
            $cmboComputeValue.SelectedItem = $WorkSpaceComputeType
        }

        $RootValue.Content = $WorkSpaceInfo.RootVolume
        $UserValue.Content = $WorkSpaceInfo.UserVolume
        $RunningModeValue.Content = $WorkSpaceInfo.RunningMode
        $IPValue.Content = $WorkSpaceInfo.IPAddress
        $RegionValue.Content = $WorkSpaceInfo.Region
        $ComputerNameValue.Content = $WorkSpaceInfo.ComputerName
        $RegCodeValue.Content = $WorkSpaceInfo.RegCode
        $lblStateValue.Content = $WorkSpaceInfo.State
        $lblProtocolValue.Content = $WorkSpaceInfo.Protocol

        try{
            $snapDate = Get-Date -Date ((Get-WKSWorkspaceSnapshot -WorkspaceId $WorkSpaceIdValue.Content -Region $RegionValue.Content).RebuildSnapshots.SnapshotTime | Out-String) -Format "MM/dd/yyyy HH:mm"  
            }catch{
            $snapDate = "Not Available"
        }
        $lblSnapShotValue.Content = $snapDate
        if($WorkSpaceInfo.LastUpdated){
            $lblLastDBUpdate.Content = $WorkSpaceInfo.LastUpdated
        }else{
            $date = (get-date -Format "MM/dd/yyyy HH:mm") | Out-String
            $lblLastDBUpdate.Content = $date
        }
        if(($ComputerNameValue.Content.substring(0,1) -eq "A")-or($ComputerNameValue.Content.substring(0,1) -eq "U")){
            $btnRemoteAssist.Visibility="Hidden"
            $btnRDP.Visibility="Hidden"
            $btnGatherLogs.Visibility="Hidden"
            $btnBackupUserVolume.Visibility="Hidden"
        }
        else{
        # Windows Machine
            $btnRemoteAssist.Visibility="Visible"
            $btnRDP.Visibility="Visible"
            $btnGatherLogs.Visibility="Visible"
            $btnBackupUserVolume.Visibility="Visible"
        }
    }
})

###############################################
# ! # ! # WorkSpaces Bulk GUI Actions # ! # ! # 
###############################################
# This section's bulk actions correspond with a GUI button action. The button objects below are 
# created from objects outlined within the XML.

$selectWKSRegion.add_SelectionChanged({
    $selectWKSDirectory.Items.Clear()
    Get-BulkDirectories
    $migrateBundleCombo.items.Clear()
    $selectWKSBundle.items.Clear()
    $migrateBundleCombo.items.add("Select Bundle")
    $bundles = $global:WorkSpacesDB | Where-Object { ($_.Region -eq $selectWKSRegion.SelectedItem)} | Select-Object BundleId, Protocol, BundleName -Unique
    if($bundles.count -eq 0){
        $selectWKSBundle.items.add("No Custom Bundles Found")
        $migrateBundleCombo.items.add("No Custom Bundles Found")
    }else{
        foreach($bundle in $bundles){
            $BundleConcat = $bundle.BundleId + " (" + $bundle.BundleName +")"
            $selectWKSBundle.items.add("$BundleConcat")
            $migrateBundleCombo.items.add("$BundleConcat")
        }
    }
    $migrateBundleCombo.SelectedIndex=0
})

# Bulk Buttons 
$selectWKSDirectory.add_SelectionChanged({
    Get-ImpactedWS
})

$cmboBulkProtocol.add_SelectionChanged({
    Get-ImpactedWS
})

$selectWKSBundle.add_SelectionChanged({
    Get-ImpactedWS
})

$selectRunningModeFilterCombo.add_SelectionChanged({
    Get-ImpactedWS
})

$btnPowerOn.Add_Click({
    if($lstImpactedWorkSpaces.SelectedItems.Count -eq 0){
        $ImpactedList = $lstImpactedWorkSpaces.Items 
    }else{
        $ImpactedList = $lstImpactedWorkSpaces.SelectedItems 
    }

    $filteredList = @()
    foreach($WS in $ImpactedList){
        if($WS.RunningMode -eq "ALWAYS_ON"){
            $WS = $WS.WorkSpaceId
            Write-logger -message "$WS is AlwaysOn, skipping in API request."
        }else{
            $filteredList += $WS
        }
    }
    $logs = Optimize-APIRequest -requestInfo $filteredList -APICall "Start-WKSWorkspace"
    if($logs.ErrorCode){
        $ErrCode = $logs.ErrorCode
        $ErrMsg = $logs.ErrorMessage
        Write-Logger -message "Start failed for some or all of the selected WorkSpaces. Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error starting, see log tab for details" -title "Error Stopping WorkSpaces"
    }else{
        Show-MessageSuccess -message "Start API on selected WorkSpaces executed successfully" -title "Successfully Started WorkSpace"
        Write-Logger -message "Start API on selected WorkSpaces executed successfully"
    }
})

$btnPowerDown.Add_Click({
    if($lstImpactedWorkSpaces.SelectedItems.Count -eq 0){
        $ImpactedList = $lstImpactedWorkSpaces.Items 
    }else{
        $ImpactedList = $lstImpactedWorkSpaces.SelectedItems 
    }

    $filteredList = @()
    foreach($WS in $ImpactedList){
        if($WS.RunningMode -eq "ALWAYS_ON"){
            $WS = $WS.WorkSpaceId
            Write-logger -message "$WS is AlwaysOn, skipping in API request."
        }else{
            $filteredList += $WS
        }
    }
    $logs = Optimize-APIRequest -requestInfo $filteredList -APICall "Stop-WKSWorkspace"
    if($logs.ErrorCode){
        $ErrCode = $logs.ErrorCode
        $ErrMsg = $logs.ErrorMessage
        Write-Logger -message "Stop failed for some or all of the selected WorkSpaces. Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error stopping, see log tab for details" -title "Error Stopping WorkSpaces"
    }else{
        Show-MessageSuccess -message "Stop API on selected WorkSpaces executed successfully" -title "Successfully Started WorkSpace"
        Write-Logger -message "Stop API on selected WorkSpaces executed successfully"
    }
})

$btnMigrate.Add_Click({
    if(($migrateBundleTxt.Text -like "Select Bundle") -or($migrateBundleTxt.Text -like "No Custom Bundles Found")){
        Show-MessageError -message "Please select a valid bundle from the dropdown before initiating." -title "Error Migrating WorkSpace"
    }else{
        if($lstImpactedWorkSpaces.SelectedItems.Count -eq 0){
            $ImpactedList = $lstImpactedWorkSpaces.Items 
        }else{
            $ImpactedList = $lstImpactedWorkSpaces.SelectedItems 
        }
        foreach ($wks in $ImpactedList){
            $bundleSTR = ($migrateBundleCombo.SelectedItem.split(" "))[0]
            $migrateReq = New-Object -TypeName PSobject
            $migrateReq | Add-Member -NotePropertyName "WorkSpaceId" -NotePropertyValue $Wks.WorkSpaceId
            $migrateReq | Add-Member -NotePropertyName "BundleId" -NotePropertyValue $bundleSTR
            $migrateReq | Add-Member -NotePropertyName "Region" -NotePropertyValue $selectWKSRegion.selectedItem

            $Output = Initialize-MigrateWorkSpace -requestInfo $migrateReq
            if($NULL -eq $Output -or $Output.SourceWorkspaceId.Count -ne 1){
                $errMsg = "Error Message: Was unable to migrate " + $Wks.WorkSpaceId
                Write-Logger -message $errMsg
                Show-MessageError -message "Error Migrating, see log tab for details" -title "Error Migrating WorkSpace"
            }else{
                $WksId = $Output.SourceWorkspaceId.ToString()
                Write-Logger -message "Migrate successfully initiated for WorkSpaceId $WksId"
                Show-MessageSuccess -message "Migrate successfully initiated for WorkSpaceId $WksId" -title "Successfully Migrated WorkSpace"
            }
        }
    }
})

$btnTerminate.Add_Click({
    if($lstImpactedWorkSpaces.SelectedItems.Count -eq 0){
        $ImpactedList = $lstImpactedWorkSpaces.Items
    }else{
        $ImpactedList = $lstImpactedWorkSpaces.SelectedItems
    }
    $filteredList = @()
    foreach($WS in $ImpactedList){
        $filteredList += $WS
    }
    $impactCount = $filteredList.Count
    $wshell = New-Object -ComObject Wscript.Shell
    $response = $wshell.Popup("Are you sure you would like to terminate the selected $impactCount WorkSpaces? This cannot be undone.",0,"Alert",64+4)
    if($response -eq 6){
        Write-Logger -message "Executing Terminate on selected WorkSpaceId(s)"
        $TermReq = New-Object -TypeName PSobject
        $TermReq | Add-Member -NotePropertyName "TermList" -NotePropertyValue $filteredList
        $Output = Optimize-APIRequest -requestInfo $TermReq -APICall "Remove-WKSWorkspace"
        if($Output.ErrorCode){
            $ErrCode = $Output.ErrorCode.ToString()
            $ErrMsg = $Output.ErrorMessage.ToString()
            Write-Logger -message "Terminate failed for selected WorkSpaces. Details below:"
            Write-Logger -message "Error Code: $ErrCode"
            Write-Logger -message "Error Message: $ErrMsg"
            Show-MessageError -message "Error Terminating, see log tab for details" -title "Error Terminating WorkSpace"
        }else{
            Show-MessageSuccess -message "Terminate API on selected WorkSpaces executed successfully" -title "Successfully Terminated WorkSpace"
            Write-Logger -message "Terminate API on selected WorkSpaces executed successfully"
        }
    }
})

$btnEnableMain.Add_Click({
    if($lstImpactedWorkSpaces.SelectedItems.Count -eq 0){
        foreach($wks in $lstImpactedWorkSpaces.Items){
            $ImpactedList += $wks
        }
    }else{
        foreach($wks in $lstImpactedWorkSpaces.SelectedItems){
            $ImpactedList += $wks
        }
    }
    $Output = Optimize-APIRequest -requestInfo $ImpactedList -APICall "Enable-Main"
    if($Output.ErrorCode){
        $ErrCode = $Output.ErrorCode.ToString()
        $ErrMsg = $Output.ErrorMessage.ToString()
        Write-Logger -message "Enabling Admin Maintenance failed. Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error enabling Admin Maintenance, see log tab for details" -title "Error Enabling Admin Maintenance"
    }else{
        Show-MessageSuccess -message "Terminate API on $WorkSpaceId executed successfully" -title "Successfully Enabled Admin Maintenance"
        Write-Logger -message "Enabling Admin Maintenance API ran successfully"
    }
})

$btnDisableMain.Add_Click({
    if($lstImpactedWorkSpaces.SelectedItems.Count -eq 0){
        $ImpactedList = $lstImpactedWorkSpaces.Items 
    }else{
        $ImpactedList = $lstImpactedWorkSpaces.SelectedItems 
    }
    $Output = Optimize-APIRequest -requestInfo $ImpactedList -APICall "Disable-Main"
    if($Output.ErrorCode){
        $ErrCode = $Output.ErrorCode.ToString()
        $ErrMsg = $Output.ErrorMessage.ToString()
        Write-Logger -message "Terminate failed for WorkSpaceId $WorkSpaceId Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error Terminating, see log tab for details" -title "Error Terminating WorkSpace"
    }else{
        Show-MessageSuccess -message "Terminate API on $WorkSpaceId executed successfully" -title "Successfully Terminated WorkSpace"
    }
})

$btnUpdateDB.Add_Click({
    Update-WorkSpaceObject
    Get-ImpactedWS
})

$btnExportBulk.Add_Click({
    if($null -ne $txtReporting.Text){
        $ImpactedList = $lstImpactedWorkSpaces.Items 
        $filteredList = @()
        foreach($WS in $ImpactedList){
            $filteredList += $WS
        }
        $output = $txtReporting.text + "BulkWorkSpaces.csv"
        try{
            $filteredList | Export-Csv -Path ($output) -NoTypeInformation
            Show-MessageSuccess -message "Successfully exported CSV to provided path" -title "Successfully Exported"
        }catch{
            Show-MessageError -message "Export failed to write CSV to provided path." -title "Export Failed"
        }
    }else{
        Show-MessageError -message "Export failed due to export path not being set on Admin tab." -title "Export Failed"
    }
})

#Button to Modify Protocol
$btnBulkModifyProtocol.Add_Click({
    if($lstImpactedWorkSpaces.SelectedItems.Count -eq 0){
        $ImpactedList = $lstImpactedWorkSpaces.Items 
    }else{
        $ImpactedList = $lstImpactedWorkSpaces.SelectedItems 
    }
    $BulkProtocolModifyReq = @()
    foreach($WS in $ImpactedList){
        $BulkProtocolModifyReq += $WS
    }
    $Output = Set-WKSProtocol -ProtocolModifyReq $BulkProtocolModifyReq
    if($Output.ErrorCode){
        $ErrCode = $Output.ErrorCode
        $ErrMsg = $Output.ErrorMessage
        Write-Logger -message "Protocol modification failed $WorkSpaceId. Details below:"
        Write-Logger -message "Error Code: $ErrCode"
        Write-Logger -message "Error Message: $ErrMsg"
        Show-MessageError -message "Error modifying protocol, see log tab for details" -title "Error Modifying Protocol"
    }else{
        Show-MessageSuccess -message "Protocol modification on $WorkSpaceId executed successfully" -title "Successfully Modified Protocol"
        Write-Logger -message "Protocol modification on $WorkSpaceId executed successfully"
    }
})

###############################################
# ! # ! # AppStream GUI # ! # ! # 
###############################################

# This function creates a object to hold all of your AppStream sessions to act as a local db 
# for the GUI
function Get-AppStreamSessions(){
    #Stop new queries, while data is populated
    $global:DataPullInProgress = $true
    $cmboAppStreamHelpDesk.items.Clear()
    $cmboAppStreamHelpDesk.items.add("*")
    $cmboAppStreamHelpDesk.SelectedIndex = 0

    $cmboAppStreamHelpDeskUserStateConnectedState.items.Clear()
    $cmboAppStreamHelpDeskUserStateConnectedState.items.add("*")
    $cmboAppStreamHelpDeskUserStateConnectedState.items.add("CONNECTED")
    $cmboAppStreamHelpDeskUserStateConnectedState.items.add("NOT_CONNECTED")
    $cmboAppStreamHelpDeskUserStateConnectedState.SelectedIndex = 0

    $cmboAppStreamHelpDeskUserState.items.Clear()
    $cmboAppStreamHelpDeskUserState.items.add("*")
    $cmboAppStreamHelpDeskUserState.items.add("ACTIVE")
    $cmboAppStreamHelpDeskUserState.SelectedIndex = 0

    $global:AppStreamDB = @()

    # Get a list of all Stacks
    $listAppStreamSessions.Items.Clear()
    if($cmboAppStreamHelpDeskRegion.SelectedIndex -ne -1 -and $cmboAppStreamHelpDeskRegion.SelectedIndex -ne 0){
        $filteredStacks = $global:TotalStacks 
        $region = $cmboAppStreamHelpDeskRegion.SelectedItem.ToString()
        $filteredStacks = $global:TotalStacks | Where-Object { ($_[0].arn -split ":")[3] -eq $region }
        foreach ($stack in $filteredStacks){
            $cmboAppStreamHelpDesk.items.add($stack.Name)
            $fleetName = $null
            $fleetName = Get-APSAssociatedFleetList -StackName $stack.Name -Region $region
            if($null -ne $fleetName){
                $SessionList = Get-APSSessionList -StackName $stack.Name -FleetName $fleetName -Region $region -limit 50
                foreach ($session in $SessionList){
                    $AS2Session = New-Object -TypeName PSobject
                    $AS2Session | Add-Member -NotePropertyName "UserId" -NotePropertyValue $session.UserId
                    $AS2Session | Add-Member -NotePropertyName "Stack" -NotePropertyValue $stack.Name
                    $AS2Session | Add-Member -NotePropertyName "State" -NotePropertyValue $session.State
                    $AS2Session | Add-Member -NotePropertyName "ConnectedState" -NotePropertyValue $session.ConnectionState
                    $AS2Session | Add-Member -NotePropertyName "StartTime" -NotePropertyValue $session.StartTime.ToString()
                    $AS2Session | Add-Member -NotePropertyName "PrivateIP" -NotePropertyValue $session.NetworkAccessConfiguration.EniPrivateIpAddress
                    $AS2Session | Add-Member -NotePropertyName "Id" -NotePropertyValue $session.Id
                    $global:AppStreamDB  += $AS2Session
                }
            }
        }
    }
    $global:DataPullInProgress=$false
    Search-AppStreamSession
}

# This functions filters your AppStream db object to only display information matching your search criteria 
function Search-AppStreamSession(){
    $listAppStreamSessions.Items.Clear()

    $filtered = $global:AppStreamDB 
    if($txtAppStreamHelpDeskUserId.Text -ne ""){
        $filtered = $filtered | Where-Object { ($_.UserId -like ($txtAppStreamHelpDeskUserId.Text + "*")) } | Select-Object UserId, Stack, State, ConnectedState, StartTime, PrivateIP, Id
    }
    if($txtAppStreamHelpDeskUserSessionId.Text -ne ""){
        $filtered = $filtered | Where-Object { ($_.SessionId -like ($txtAppStreamHelpDeskUserSessionId.Text + "*")) } | Select-Object UserId, Stack, State, ConnectedState, StartTime, PrivateIP, Id
    }
    if($txtAppStreamHelpDeskUserIP.Text -ne ""){
        $filtered = $filtered | Where-Object { ($_.PrivateIP -like ($txtAppStreamHelpDeskUserIP.Text + "*")) } | Select-Object UserId, Stack, State, ConnectedState, StartTime, PrivateIP, Id
    }
    if($cmboAppStreamHelpDesk.SelectedItem.ToString() -ne ""){
        $filtered = $filtered | Where-Object { ($_.Stack -like ($cmboAppStreamHelpDesk.SelectedItem.ToString())) } | Select-Object UserId, Stack, State, ConnectedState, StartTime, PrivateIP, Id
    }
    if($cmboAppStreamHelpDeskUserState.SelectedItem.ToString() -ne ""){
        $filtered = $filtered | Where-Object { ($_.State -like ($cmboAppStreamHelpDeskUserState.SelectedItem.ToString())) } | Select-Object UserId, Stack, State, ConnectedState, StartTime, PrivateIP, Id
    }
    if($cmboAppStreamHelpDeskUserState.SelectedItem.ToString() -ne ""){
        $filtered = $filtered | Where-Object { ($_.ConnectedState -like ($cmboAppStreamHelpDeskUserStateConnectedState.SelectedItem.ToString())) } | Select-Object UserId, Stack, State, ConnectedState, StartTime, PrivateIP, Id
    }

    foreach($session in $filtered){
        $listAppStreamSessions.items.Add($session)
    }
}

# This function pulls all regions that you have deployed AppStream 
function Get-AppStreamRegions(){
    $cmboAppStreamHelpDeskRegion.items.clear()
    $SilenceOutput=$cmboAppStreamHelpDeskRegion.items.add("Select Region")
    $global:TotalStacks = @()
    $regions = @('us-east-1', 'us-east-2', 'us-west-2', 'ap-south-1', 'ap-northeast-2', 'ap-southeast-1', 'ap-southeast-2', 'ap-northeast-1', 'ca-central-1', 'eu-central-1','eu-west-1', 'eu-west-2')
    foreach ($region in $regions){
        $TempStacks = $null
        $TempStacks = Get-APSStackList -Region $region
        if($null -ne $TempStacks){
            $global:TotalStacks += $TempStacks
            if($cmboAppStreamHelpDeskRegion.Items -notcontains $region){
                $SilenceOutput=$cmboAppStreamHelpDeskRegion.items.add("$region")
            }
        }
    }
    $cmboAppStreamHelpDeskRegion.SelectedIndex=0
}

# This section's AppStream actions correspond with a GUI button action. The button objects below are 
# created from objects outlined within the XML.

$btnAppStreamHelpDeskSessionDisconnect.Add_Click({
    $sessionId=$listAppStreamSessions.SelectedItems.Id
    Revoke-APSSession -SessionId $sessionId
    Start-Sleep -Seconds 2
    Get-AppStreamSessions
})

$btnAppStreamHelpDeskRemoteAssist.Add_Click({
    $privateIP=$listAppStreamSessions.SelectedItems.PrivateIP
    $parameters="/offerRA $privateIP"
    $exe="msra.exe"
    start-process $exe $parameters -Wait    
})

$txtAppStreamHelpDeskUserId.Add_TextChanged({
    Search-AppStreamSession
})
$txtAppStreamHelpDeskUserSessionId.Add_TextChanged({
    Search-AppStreamSession
})
$txtAppStreamHelpDeskUserIP.Add_TextChanged({
    Search-AppStreamSession
})

$cmboAppStreamHelpDesk.add_SelectionChanged({
    if(!$global:DataPullInProgress){
        Search-AppStreamSession
    }
})

$cmboAppStreamHelpDeskUserState.add_SelectionChanged({
    if(!$global:DataPullInProgress){
        Search-AppStreamSession
    }
})
$cmboAppStreamHelpDeskUserStateConnectedState.add_SelectionChanged({
    if(!$global:DataPullInProgress){
        Search-AppStreamSession
    }
})

$cmboAppStreamHelpDeskRegion.add_SelectionChanged({
    Get-AppStreamSessions
})

$btnAppStreamHelpDeskSessionExport.Add_Click({
    if($null -ne $txtReporting.Text){
        $ImpactedList = $listAppStreamSessions.Items 
        $filteredList = @()
        foreach($WS in $ImpactedList){
            $filteredList += $WS
        }
        $output = $txtReporting.text + "AppStreamSessions.csv"
        try{
            $filteredList | Export-Csv -Path ($output) -NoTypeInformation
            Show-MessageSuccess -message "Successfully exported CSV to provided path" -title "Successfully Exported"
        }catch{
            Show-MessageError -message "Export failed to write CSV to provided path." -title "Export Failed"
        }
    }else{
        Show-MessageError -message "Export failed due to export path not being set on Admin tab." -title "Export Failed"
    }
})

###############################################
# ! # ! # WorkSpaces Admin GUI Actions # ! # ! # 
###############################################

$btnQueryWSUpdateDB.Add_Click({
    Update-WorkSpaceObject
})

$cmboAdminSelectRegionValue.add_SelectionChanged({
    if(($cmboAdminSelectRegionValue.SelectedValue) -ne 'Select Region'){
        Update-WorkSpaceAdminDirectories
        Update-ServiceQuotas
    }
})

$cmboAdminWSDirectory.add_SelectionChanged({
    if(($cmboDeploymentWSDirectory.SelectedValue) -ne 'Select a Directory'){
        Update-WorkSpaceAdminDirectoryDetails
    }
})

#Function loads all of the directories in a region.
function Update-WorkSpaceAdminDirectories(){    
    $cmboAdminWSDirectory.Items.Clear()
    $cmboAdminWSDirectory.Items.Add("Select a Directory")
    $Directories = $global:WorkSpacesDirectoryDB | Where-Object { ($_.Region -eq $cmboAdminSelectRegionValue.SelectedItem)} | Select-Object directoryId, Region -Unique 
    foreach ($Directory in $Directories ){
        $directorySTR = $Directory.directoryId
        $cmboAdminWSDirectory.Items.Add($directorySTR)
    }
    $cmboAdminWSDirectory.SelectedIndex=0
}

function Update-WorkSpaceAdminDirectoryDetails(){
    $directoryInfo = $global:WorkSpacesDirectoryDB | Where-Object { ($_.directoryId -eq $cmboAdminWSDirectory.SelectedValue)} | Get-Unique
    $lblAdminDirectoryAliasContent.Content = $directoryInfo.directoryAlias
    $lblAdminDirectoryNameContent.Content = $directoryInfo.directoryName
    $lblAdminDirectoryIdContent.Content = $directoryInfo.directoryId
    $lblAdminDirectoryRegCodeContent.Content = $directoryInfo.RegistrationCode
    $lblAdminDirectoryStateContent.Content = $directoryInfo.directoryState
    $lblAdminDirectoryLocalAdinContent.Content = $directoryInfo.directoryUserEnabledAsLocalAdministrator
    $lblAdminDirectoryTenancyContent.Content = $directoryInfo.directoryTenancy
    $lblAdminDirectoryIPContent.Content = $directoryInfo.directoryAvailableIPs
}

function Update-ServiceQuotas(){
    $targetDirectory = $global:WorkSpacesServiceQuotaDB | Where-Object { ($_.Region -eq $cmboAdminSelectRegionValue.SelectedValue)} | Get-Unique

    $lblAdminWSServiceQuotaCurrent.Content = (($global:WorkSpacesDB | Where-Object { ($_.Region -eq $cmboAdminSelectRegionValue.SelectedValue)}).count)
    $lblAdminWSServiceQuotaCurrentMax.Content = $targetDirectory.quotaWks

    $lblAdminWSStandByServiceQuotaCurrent.Content = "N/A"
    $lblAdmintWSStandByServiceQuotaMax.Content = $targetDirectory.quotaStandby

    $lblAdminWSGraphicsServiceQuotaCurrent.Content = (($global:WorkSpacesDB | Where-Object { ($_.Region -eq $cmboAdminSelectRegionValue.SelectedValue)} | Where-Object {($_.WorkspaceProperties.ComputeTypeName).ToUpper -eq "GRAPHICS"}).count)
    $lblAdminWSGraphicsServiceQuotaMax.Content = $targetDirectory.quotaGraphics

    $lblAdminWSGraphicsProServiceQuotaCurrent.Content = (($global:WorkSpacesDB | Where-Object { ($_.Region -eq $cmboAdminSelectRegionValue.SelectedValue)} | Where-Object {($_.WorkspaceProperties.ComputeTypeName).ToUpper -eq "GRAPHICSPRO"}).count)
    $lblAdminWSGraphicsProServiceQuotaMax.Content = $targetDirectory.quotaGraphicsPro

    $lblAdminWSGraphicsg4dnServiceQuotaCurrent.Content = (($global:WorkSpacesDB | Where-Object { ($_.Region -eq $cmboAdminSelectRegionValue.SelectedValue)} | Where-Object {($_.WorkspaceProperties.ComputeTypeName).ToUpper -eq "GRAPHICS.G4DN"}).count)
    $lblAdminWSGraphicsg4dnProServiceQuotaMax.Content = $targetDirectory.quotaG4dn

    $lblAdminWSGraphicsg4dnProServiceQuotaCurrent.Content = (($global:WorkSpacesDB | Where-Object { ($_.Region -eq $cmboAdminSelectRegionValue.SelectedValue)} | Where-Object {($_.WorkspaceProperties.ComputeTypeName).ToUpper -eq "GRAPHICSPRO.G4DN"}).count)
    $lblAdminWSGraphicsg4dnServiceQuotaMax.Content = $targetDirectory.quotaG4dnPro
}


# Populate protocol dropown options
$SilenceOutput = $cmboProtocol.items.Add("All")
$SilenceOutput = $cmboProtocol.items.Add("BYOP")
$SilenceOutput = $cmboProtocol.items.Add("WSP")
$SilenceOutput = $cmboProtocol.items.Add("PCoIP")
$SilenceOutput = $cmboProtocol.SelectedIndex=0
$SilenceOutput = $cmboBulkProtocol.items.Add("All")
$SilenceOutput = $cmboBulkProtocol.items.Add("BYOP")
$SilenceOutput = $cmboBulkProtocol.items.Add("WSP")
$SilenceOutput = $cmboBulkProtocol.items.Add("PCoIP")
$cmboBulkProtocol.SelectedIndex=0
# Initialization

Update-WorkSpaceObject

Search-WorkSpaces

$SilenceOutput = Get-AppStreamSessions
$regions = $global:WorkSpacesDB.Region | Select-Object -Unique
foreach($region in $regions){
    if(-not(($selectWKSRegion.Items).Contains($region))){
        $SilenceOutput=$selectWKSRegion.Items.Add($region)
        $SilenceOutput=$cmboAdminSelectRegionValue.Items.Add($region)
    }
}

# Preset admin tab textboxes
$tmpPath = $($PSScriptRoot)+"\Assets\"
	   
# Preset admin tab textboxes
$tmpPath = $($PSScriptRoot)+"\Assets\"

# if there is a settings CSV with Values
$settingsCSVFile=$tmpPath+"Settings.csv"
if ([System.IO.File]::Exists($settingsCSVFile)) {
    $loadSettings =Import-Csv $settingsCSVFile
    $loadSettings |ForEach{
        $txtServerSideLogs.Text = $_.WorkSpaceSideLogs
        $txtBackUpDest.Text = $_.Backups
        $txtPSExecPath.Text = $_.PSEXEC 
        $txtDisk2VHDPath.Text = $_.Disk2VHD
        $txtReporting.Text = $_.Reporting
        $txtCloudWatchAccessLogs.Text= $_.CloudWatchAccessLogs
    }
} else {
    $txtServerSideLogs.Text = $tmpPath
    $txtBackUpDest.Text = $tmpPath
    $txtPSExecPath.Text = $tmpPath
    $txtDisk2VHDPath.Text = $tmpPath
    $txtReporting.Text = $tmpPath
}

# Populate Compute Dropdown in Main
$selectRunningModeFilterCombo.items.Clear()
$SilenceOutput=$selectRunningModeFilterCombo.items.Add("Select Running Mode")
$SilenceOutput=$selectRunningModeFilterCombo.items.Add("ALWAYS_ON")
$SilenceOutput=$selectRunningModeFilterCombo.items.Add("AUTO_STOP")
$selectRunningModeFilterCombo.SelectedIndex=0
Get-AppStreamRegions

# Populate Bundle Combos
$SilenceOutput=$SilenceOutput=$migrateBundleCombo.items.add("Select Bundle")
$cmboAdminSelectRegionValue.SelectedIndex=0

# Set index on Main items 
$SearchResults.SelectedIndex=0

# Set index on Region Bulk selection
$selectWKSRegion.SelectedIndex=0

$global:cloudWatchRun=0

Write-Logger -message "EUC Toolkit Launched"
write-host "Powershell GUI Loaded"

#Show Form
$WksMainForm.ShowDialog() | out-null