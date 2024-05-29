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
    This script module provides helper functions for the EUC Toolkit.
.DESCRIPTION
    This script module performs the heavy lifting for the EUC toolkit. This includes building the local db object, optimiziing API calls,
    returning error and success messages, generating dhasboard images, and othe functionality part of the EUC toolkit. For more information,
    see the link below:
    https://github.com/aws-samples/euc-toolkit
#>


function Get-LocalWorkSpacesDB(){
    # This function build a PSObject that contains all of your WorkSpaces information. The object will act as a local DB for the GUI.
    # If you need to have object persistence to save API calls, this function can be replaced with a function that calls your persistent store.
    $WorkSpacesDDB = @()
    $DeployedRegions = Get-WksDirectories
    
    # Finds all current WorkSpaces. If Actice Directory cannot be reached, those attributes are omitted.
    foreach($DeployedRegion in $DeployedRegions){
        $RegionalWks = Get-WKSWorkSpaces -Region $DeployedRegion.Region -DirectoryId $DeployedRegion.DirectoryId
        $RegionalBundles = Get-AllBundles -Region $DeployedRegion.Region -Custom $true
        foreach ($Wks in $RegionalWks){
            $adErr = $false
            $entry = New-Object -TypeName PSobject
            $entry | Add-Member -NotePropertyName "WorkSpaceId" -NotePropertyValue $Wks.WorkspaceId
            $entry | Add-Member -NotePropertyName "Region" -NotePropertyValue $DeployedRegion.Region
            $entry | Add-Member -NotePropertyName "UserName" -NotePropertyValue $Wks.UserName 
            $entry | Add-Member -NotePropertyName "ComputerName" -NotePropertyValue $Wks.ComputerName
            $entry | Add-Member -NotePropertyName "Compute" -NotePropertyValue $Wks.WorkspaceProperties.ComputeTypeName | Out-String
            $entry | Add-Member -NotePropertyName "RootVolume" -NotePropertyValue $Wks.WorkspaceProperties.RootVolumeSizeGib 
            $entry | Add-Member -NotePropertyName "UserVolume" -NotePropertyValue $Wks.WorkspaceProperties.UserVolumeSizeGib
            $entry | Add-Member -NotePropertyName "RunningMode" -NotePropertyValue $Wks.WorkspaceProperties.RunningMode
            #Update for Protocol
            if($Wks.WorkspaceProperties.Protocols -like "WSP"){$wsProto = 'WSP'}elseif ($Wks.WorkspaceProperties.Protocols -like "PCOIP"){$wsProto = 'PCoIP'} else{$wsProto = 'BYOP'}
            $entry | Add-Member -NotePropertyName "Protocol" -NotePropertyValue $wsProto
            $entry | Add-Member -NotePropertyName "IPAddress" -NotePropertyValue $Wks.IPAddress
            $entry | Add-Member -NotePropertyName "RegCode" -NotePropertyValue ($DeployedRegion | Where-Object {$_.directoryId -eq $Wks.directoryId}).RegistrationCode
            $entry | Add-Member -NotePropertyName "directoryId" -NotePropertyValue $Wks.directoryId
            $entry | Add-Member -NotePropertyName "directoryName" -NotePropertyValue ($DeployedRegion | Where-Object {$_.directoryId -eq $Wks.directoryId}).DirectoryName
            $entry | Add-Member -NotePropertyName "directoryAlias" -NotePropertyValue ($DeployedRegion | Where-Object {$_.directoryId -eq $Wks.directoryId}).DirectoryAlias
            $entry | Add-Member -NotePropertyName "directoryType" -NotePropertyValue ($DeployedRegion | Where-Object {$_.directoryId -eq $Wks.directoryId}).DirectoryType
            $entry | Add-Member -NotePropertyName "directoryState" -NotePropertyValue ($DeployedRegion | Where-Object {$_.directoryId -eq $Wks.directoryId}).DirectoryState
            $entry | Add-Member -NotePropertyName "directoryUserEnabledAsLocalAdministrator" -NotePropertyValue ($DeployedRegion | Where-Object {$_.directoryId -eq $Wks.directoryId}).DirectoryUserEnabledAsLocalAdministrator
            $entry | Add-Member -NotePropertyName "directoryTenancy" -NotePropertyValue ($DeployedRegion | Where-Object {$_.directoryId -eq $Wks.directoryId}).DirectoryTenancy
            $entry | Add-Member -NotePropertyName "directoryAvailableIPs" -NotePropertyValue ($DeployedRegion | Where-Object {$_.directoryId -eq $Wks.directoryId}).DirectoryAvailableIPs
            $entry | Add-Member -NotePropertyName "State" -NotePropertyValue $Wks.State
            $entry | Add-Member -NotePropertyName "BundleId" -NotePropertyValue $Wks.BundleId
            if((($RegionalBundles | Where-Object {$_.BundleId -like $Wks.BundleId}).Name) -eq ''){
                $unavailablePubBundle = Get-WKSWorkspaceBundle -BundleId $Wks.BundleId -Region $DeployedRegion.Region
                $entry | Add-Member -NotePropertyName "BundleName" -NotePropertyValue $unavailablePubBundle.Name
            }else{
                $entry | Add-Member -NotePropertyName "BundleName" -NotePropertyValue ($RegionalBundles | Where-Object {$_.BundleId -eq $Wks.BundleId}).Name
            }
            try{
                $ADUser = Get-ADUser -Identity $Wks.UserName -Properties "EmailAddress"
            }catch{
                $adErr = $true
            }
            if($adErr -eq $false){
                $entry | Add-Member -NotePropertyName "FirstName" -NotePropertyValue ($ADUser.GivenName)
                $entry | Add-Member -NotePropertyName "LastName" -NotePropertyValue ($ADUser.Surname)
                $entry | Add-Member -NotePropertyName "Email" -NotePropertyValue ($ADUser.EmailAddress)
                $WorkSpacesDDB += $entry
            }else{
                $entry | Add-Member -NotePropertyName "FirstName" -NotePropertyValue "AD Info Not Available"
                $entry | Add-Member -NotePropertyName "LastName" -NotePropertyValue "AD Info Not Available"
                $entry | Add-Member -NotePropertyName "Email" -NotePropertyValue "AD Info Not Available"
                $WorkSpacesDDB += $entry
            }
        }
    }
    return $WorkSpacesDDB
}

function Get-WksDirectories(){
    $DeployedRegions = @()

    # Current WorkSpaces Regions. See link below for current WorkSpaces availbility
    # https://docs.aws.amazon.com/workspaces/latest/adminguide/azs-workspaces.html
    $regions = @('us-east-1','us-west-2', 'ap-south-1', 'ap-northeast-2', 'ap-southeast-1', 'ap-southeast-2', 'ap-northeast-1', 'ca-central-1', 'eu-central-1','eu-west-1', 'eu-west-2', 'sa-east-1')

    # Find regions that have WorkSpaces deployments
    foreach($region in $regions){
        $RegionsCall = Get-WKSWorkspaceDirectories -Region $region
        if($RegionsCall){
            foreach($WksRegion in $RegionsCall){
                $DeployedRegionsTemp = New-Object -TypeName PSobject
                $DeployedRegionsTemp | Add-Member -NotePropertyName "Region" -NotePropertyValue $region
                $DeployedRegionsTemp | Add-Member -NotePropertyName "RegistrationCode" -NotePropertyValue $WksRegion.RegistrationCode
                $DeployedRegionsTemp | Add-Member -NotePropertyName "DirectoryId" -NotePropertyValue $WksRegion.DirectoryId
                $DeployedRegionsTemp | Add-Member -NotePropertyName "DirectoryName" -NotePropertyValue $WksRegion.DirectoryName
                $DeployedRegionsTemp | Add-Member -NotePropertyName "DirectoryAlias" -NotePropertyValue $WksRegion.Alias
                $DeployedRegionsTemp | Add-Member -NotePropertyName "DirectoryType" -NotePropertyValue $WksRegion.Type
                $DeployedRegionsTemp | Add-Member -NotePropertyName "DirectoryState" -NotePropertyValue $WksRegion.State
                $DeployedRegionsTemp | Add-Member -NotePropertyName "DirectoryUserEnabledAsLocalAdministrator" -NotePropertyValue $WksRegion.WorkspaceCreationProperties.UserEnabledAsLocalAdministrator
                $DeployedRegionsTemp | Add-Member -NotePropertyName "DirectoryTenancy" -NotePropertyValue $WksRegion.Tenancy
                $subnetA = Get-EC2Subnet -SubnetId $WksRegion.SubnetIds[0] -Region $region
                $subnetB = Get-EC2Subnet -SubnetId $WksRegion.SubnetIds[1] -Region $region
                $dirAvailableIPs = $subnetA.AvailableIpAddressCount + $subnetB.AvailableIpAddressCount
                $DeployedRegionsTemp | Add-Member -NotePropertyName "DirectoryAvailableIPs" -NotePropertyValue $dirAvailableIPs
                #Overall WorkSpaces
                $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaWks" -NotePropertyValue (Get-SQServiceQuota -ServiceCode workspaces -QuotaCode "L-34278094" -Region $region)
                #StandBy WorkSpaces
                try{
                    $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaStandby" -NotePropertyValue (Get-SQServiceQuota -ServiceCode workspaces -QuotaCode "L-9A67B5CB" -Region $region)
                }catch{
                    $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaStandby" -NotePropertyValue "n/a"
                }
                #Graphics WorkSpaces
                try{
                    $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaG" -NotePropertyValue (Get-SQServiceQuota -ServiceCode workspaces -QuotaCode "L-84611756" -Region $region )
                }catch{
                    $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaG" -NotePropertyValue "n/a"
                }
                #GraphicsPro WorkSpaces
                try{
                    $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaGPro" -NotePropertyValue (Get-SQServiceQuota -ServiceCode workspaces -QuotaCode "L-254B485B" -Region $region)
                }catch{
                    $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaGPro" -NotePropertyValue "n/a"
                }
                #Graphics.g4dn WorkSpaces
                $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaG4dn" -NotePropertyValue (Get-SQServiceQuota -ServiceCode workspaces -QuotaCode "L-BCACAEBC" -Region $region)
                #Graphics.g4dn WorkSpaces Pro
                $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaG4dnPro" -NotePropertyValue (Get-SQServiceQuota -ServiceCode workspaces -QuotaCode "L-BE9A8466" -Region $region)
                $DeployedRegions += $DeployedRegionsTemp
            }
        }
    }

    return $DeployedRegions
}

function Get-AllBundles(){
    param(
        $Region,
        $Custom
    )
    if($custom -eq $true){
        $Bundles = Get-WKSWorkspaceBundle -Region $Region
        $Bundles += Get-WKSWorkspaceBundle -Region $Region -Owner 'AMAZON'
    }else{
        $Bundles = Get-WKSWorkspaceBundle -Region $Region -Owner 'AMAZON'
    }
    return $Bundles
}

function Update-RunningMode(){
    # This function updates the running mode on a selected WorkSpace. For more information see the link below:
    # https://docs.aws.amazon.com/powershell/latest/reference/items/Edit-WKSWorkspaceProperty.html
    param(
        $UpdateReq
    )
    $WorkSpaceId = $UpdateReq.WorkSpaceId
    $region = $UpdateReq.Region

    if($UpdateReq.CurrentRunMode -eq "AUTO_STOP"){    
        $callBlock = "Edit-WKSWorkspaceProperty -WorkspaceId $WorkSpaceId -Region $region -WorkspaceProperties_RunningMode ALWAYS_ON"
    }
    elseif($UpdateReq.CurrentRunMode -eq "ALWAYS_ON"){
        $callBlock = "Edit-WKSWorkspaceProperty -WorkspaceId $WorkSpaceId -Region $region -WorkspaceProperties_RunningMode AUTO_STOP"
    }
    try{
        $scriptblock = [Scriptblock]::Create($callBlock)
        $logging = Invoke-Command -scriptblock $scriptblock
    }Catch{
        $msg = $_
        $logging = New-Object -TypeName PSobject
        $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error During Running Mode Change"
        $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
    }
    return $logging 
}

function Update-ComputeType(){
    # This function updates the compute type associated with a targeted WorkSpace. For more information see the link below:
    # https://docs.aws.amazon.com/powershell/latest/reference/items/Edit-WKSWorkspaceProperty.html 
    param(
        $ComputeReq
    )
    $WorkSpaceId = $ComputeReq.WorkSpaceId
    $region = $ComputeReq.Region 
    $TargetCompute = $ComputeReq.TargetCompute
    $callBlock = "Edit-WKSWorkspaceProperty -WorkspaceId $WorkSpaceId -Region $Region -WorkspaceProperties_ComputeTypeName $TargetCompute"
    $scriptblock = [Scriptblock]::Create($callBlock)
    try{
        $logging = Invoke-Command -scriptblock $scriptblock
    }Catch{
        $msg = $_
        $logging = New-Object -TypeName PSobject
        $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error During Running Mode Change"
        $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
    }
    return $logging 
}

function Invoke-RemoteAssist{
    # This function invokes the remote assist functionality. Refer to the blog and README for more info.
    param(
        [String]$privateIP
    )

    $parameters="/offerRA " +$privateIP
    $exe="msra.exe"
    start-process $exe $parameters -Wait
    
} 

function Update-RootVolume{
    # This function increases the available capacity on the WorkSpaces' Root volume. For more information see the links below:
    # https://docs.aws.amazon.com/workspaces/latest/adminguide/modify-workspaces.html#change_volume_sizes
    # https://docs.aws.amazon.com/powershell/latest/reference/items/Edit-WKSWorkspaceProperty.html
    param(
        $WorkSpaceReq
    )
    $CurrentUser = $WorkSpaceReq.CurrentUserStorage
    $CurrentRoot = $WorkSpaceReq.CurrentRootStorage
    $WorkSpaceId = $WorkSpaceReq.WorkSpaceId
    $Region = $WorkSpaceReq.Region
    if($CurrentUser -lt 100 -and $CurrentRoot -lt 175){
        $msg = "Unable to extend $WorkSpaceId's Root volume until you have increased the User volume to 100GB (Currently User is $CurrentUser GB)."
        #Write-Logger -message $msg
        $logging = New-Object -TypeName PSobject
        $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error During Root Increase"
        $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
    }elseif($CurrentUser -ge 100){
        $msg = "Enter new Root Volume size in GB. Note, it must be larger or equal to 175 GB."

        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $title = 'Root Volume Increase'
        $requestedSize = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
        $requestedSize = $requestedSize -as [int]
    
        if ($requestedSize -ge 175 -and $requestedSize -gt $CurrentRoot){
            try{
                $callBlock = "Edit-WKSWorkspaceProperty -WorkspaceId $WorkSpaceId -Region $region -WorkspaceProperties_RootVolumeSizeGib $requestedSize"
                $scriptblock = [Scriptblock]::Create($callBlock)
                $logging += Invoke-Command -scriptblock $scriptblock
            }Catch{
                $msg = $_
                $logging = New-Object -TypeName PSobject
                $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error During User Increase"
                $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
            }
            
        }else{
            $msg = "$WorkSpaceId was unable to extend its Root Volume to $requested GB since its not greater than or equal to 175."
            $logging = New-Object -TypeName PSobject
            $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error During Root Increase"
            $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
        }
    }else{
        $msg = "$WorkSpaceId was unable to extend its Root Volume to $requested GB since your User Volume is not equal to 100GB."
        $logging = New-Object -TypeName PSobject
        $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error During Root Increase"
        $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
    }
    return $logging
}  

function Set-WKSProtocol{
    param(
        $ProtocolModifyReq
    )
    $logging = @()

    foreach($req in $ProtocolModifyReq){
        $WorkSpaceId = $req.WorkSpaceId 
        $Region = $req.Region
        if($req.Protocol -eq 'PCOIP'){
            $TargetProtocol = 'WSP'
        }
        elseif($req.Protocol -eq 'WSP'){
            $TargetProtocol = 'PCOIP'
        }
        $callBlock = "Edit-WKSWorkspaceProperty -WorkspaceId $WorkSpaceId -Region $Region -WorkspaceProperties_Protocols $TargetProtocol"
        $scriptblock = [Scriptblock]::Create($callBlock)
        try{
            $logging += Invoke-Command -scriptblock $scriptblock
        }Catch{
            $msg = $_
            $tmplogging = New-Object -TypeName PSobject
            $tmplogging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error Terminating WorkSpaces"
            $tmplogging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
            $logging += $tmplogging
        }
    }

    return $logging 
} 

function Update-UserVolume{
    # This function increases the available capacity on the WorkSpaces' User volume. For more information see the links below:
    # https://docs.aws.amazon.com/workspaces/latest/adminguide/modify-workspaces.html#change_volume_sizes
    # https://docs.aws.amazon.com/powershell/latest/reference/items/Edit-WKSWorkspaceProperty.html    
    param(
        $WorkSpaceReq
    )
    $CurrentUser = $WorkSpaceReq.CurrentUserStorage
    $CurrentRoot = $WorkSpaceReq.CurrentRootStorage
    $WorkSpaceId = $WorkSpaceReq.WorkSpaceId
    $Region = $WorkSpaceReq.Region
    if($CurrentRoot -lt 175){
        $msg = "Enter new User Volume size in GB. Note, it must be larger than $CurrentUser GB and, to go above 100GB, the Root volume will need to be atleast 175GB."
    }
    elseif($CurrentRoot -ge 175){
        $msg = "Enter new User Volume size in GB. Note, it must be larger than $CurrentUser GB."
    }
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $title = 'User Volume Increase'
    $requestedSize = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    $requestedSize = $requestedSize -as [int]

    if ($requestedSize -le $CurrentUser){
        $msg = "$WorkSpaceId was unable to extend its User Volume to $requested GB since it currently is $CurrentUser GB."
        $logging = New-Object -TypeName PSobject
        $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error During User Increase"
        $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
    }
    elseif($requestedSize -gt 100 -and $CurrentRoot -lt 175){
        $msg = "$WorkSpaceId was unable to extend its User Volume to $requested GB since Root is not at least 175GB (Currently $CurrentRoot GB)."
        $logging = New-Object -TypeName PSobject
        $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error During User Increase"
        $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
    }
    else{
        $callBlock = "Edit-WKSWorkspaceProperty -WorkspaceId $WorkSpaceId -Region $Region -WorkspaceProperties_UserVolumeSizeGib $requestedSize"
        $scriptblock = [Scriptblock]::Create($callBlock)
        try{
            $logging = Invoke-Command -scriptblock $scriptblock
        }Catch{
            $msg = $_
            $logging = New-Object -TypeName PSobject
            $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error During User Increase"
            $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
        }
    }

    return $logging
}   

function Optimize-APIRequest{
    # This function optimizes several API calls so that the GUI can pass in a large list of targets, and APIs
    # can be optimized to use as few calls as possible. Some APIs can be built out to call 25 targets at a time
    # while some cannot. See the link below for more information:
    # https://docs.aws.amazon.com/powershell/latest/reference/items/WorkSpaces_cmdlets.html
    [CmdletBinding()]
    param(
        $requestInfo,
        $APICall
    )
    $callBlock = ""
    $counter = 0
    $logging = @()
    $builder = @()
    
    if($APICall.split("-")[0] -like "Start" -or $APICall.split("-")[0] -like "Stop" -or $APICall.split("-")[0] -like "Restart"){
        foreach($call in $requestInfo){
            $counter++
            if($counter -eq $requestInfo.WorkSpaceId.count){
                $builder += $call.WorkSpaceId
                $region = $call.Region
                $callBlock = "$APICall -Region $region -WorkSpaceId $builder"
                $scriptblock = [Scriptblock]::Create($callBlock) 
                $logging += Invoke-Command -scriptblock $scriptblock
            }
            elseif($counter % 25 -ne 0){ 
                $builder += $call.WorkSpaceId + ","
            }else{
                $builder += $call.WorkSpaceId
                $region = $call.Region
                $callBlock = "$APICall -Region $region -WorkSpaceId $builder"
                $scriptblock = [Scriptblock]::Create($callBlock) 
                $logging += Invoke-Command -scriptblock $scriptblock
                $builder = @()
            }
        }
    }
    elseif($APICall.split("-")[0] -like "Remove"){
        foreach($call in $requestInfo.TermList){
            $counter++
            if($counter -eq $requestInfo.TermList.count){
                $builder += $call.WorkSpaceId
                $region = $call.Region
                $callBlock = "$APICall -Region $region -WorkSpaceId $builder -force"
                $scriptblock = [Scriptblock]::Create($callBlock) 
                try{
                    $logging += Invoke-Command -scriptblock $scriptblock
                }Catch{
                    $msg = $_
                    $tmplogging = New-Object -TypeName PSobject
                    $tmplogging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error Terminating WorkSpaces"
                    $tmplogging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
                    $logging += $tmplogging
                }
            }
            elseif($counter % 25 -ne 0){
                $builder += $call.WorkSpaceId + ","
            }else{
                $builder += $call.WorkSpaceId
                $region = $call.Region
                $callBlock = "$APICall -Region $region -WorkSpaceId $builder -force"
                $scriptblock = [Scriptblock]::Create($callBlock) 
                try{
                    $logging += Invoke-Command -scriptblock $scriptblock
                }Catch{
                    $msg = $_
                    $tmplogging = New-Object -TypeName PSobject
                    $tmplogging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error Terminating WorkSpaces"
                    $tmplogging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
                    $logging += $tmplogging
                }
                $builder = @()
            }
        }
    }
    elseif($APICall.split("-")[0] -like "Reset" -or $APICall.split("-")[0] -like "Restore"){
        foreach($call in $requestInfo){
            $builder = $call.WorkSpaceId
            $region = $call.Region
            $callBlock = "$APICall -Region $region -WorkSpaceId $builder"
            $scriptblock = [Scriptblock]::Create($callBlock) 
            try{
                $logging += Invoke-Command -scriptblock $scriptblock
            }Catch{
                $msg = $_
                $logging = New-Object -TypeName PSobject
                $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error During WorkSpaces $APICall"
                $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
            }
        }
    }
    elseif($APICall.split("-")[0] -like "Enable"){
        foreach($call in $requestInfo){
            $callBlock = "Edit-WKSWorkspaceState -WorkspaceId " + $call.WorkSpaceId + " -WorkspaceState ADMIN_MAINTENANCE -Region " + $call.Region
            $scriptblock = [Scriptblock]::Create($callBlock) 
            try{
                $logging += Invoke-Command -scriptblock $scriptblock
            }Catch{
                $msg = $_
                $logging = New-Object -TypeName PSobject
                $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error Enabling Admin Maintenance"
                $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
            }
        }
    }
    elseif($APICall.split("-")[0] -like "Disable"){
        foreach($call in $requestInfo){
            $callBlock = "Edit-WKSWorkspaceState -WorkspaceId " + $call.WorkSpaceId + " -WorkspaceState AVAILABLE -Region " + $call.Region
            $scriptblock = [Scriptblock]::Create($callBlock) 
            try{
                $logging += Invoke-Command -scriptblock $scriptblock
            }Catch{
                $msg = $_
                $logging = New-Object -TypeName PSobject
                $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error Disabling Admin Maintenance"
                $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
            }
        }
    }
    return $logging
}

function Initialize-MigrateWorkSpace(){
    # This function will migrate a WorkSpace to a new bundle. See the links below for more information:
    # https://docs.aws.amazon.com/workspaces/latest/adminguide/migrate-workspaces.html
    # https://docs.aws.amazon.com/powershell/latest/reference/items/Start-WKSWorkspaceMigration.html
    [CmdletBinding()]
    param(
        $requestInfo
    )

    $logging = $null
    $callBlock = "Start-WKSWorkspaceMigration -SourceWorkspaceId " + $requestInfo.WorkSpaceId + " -BundleId " + $requestInfo.BundleId + " -Region " + $requestInfo.Region
    $scriptblock = [Scriptblock]::Create($callBlock) 
    try{
        $logging = Invoke-Command -scriptblock $scriptblock
    }Catch{
        $msg = $_
        $logging = New-Object -TypeName PSobject
        $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error During Migrate"
        $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
    }
    
    return $logging
}

function Show-MessageError(){
    # This function provides visual error messages for the GUI.
    param(
        [String]$message,
        [String]$title
    )
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    [System.Windows.MessageBox]::Show($message,$title,'OK','Error')
}
function Show-MessageSuccess(){
    # This function provides visual success messages for the GUI.
    param(
        [String]$message,
        [String]$title
    )
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    [System.Windows.MessageBox]::Show($message,$title,'OK')
}

function Get-CloudWatchImagesServiceMetrics(){
    # This function will generate dashboard images for your metrics portal. It targets the deafult metrics for WorkSpaces. 
    param(
        [String]$workingDirectory,
        [String]$WSId,
        [String]$AWSProfile,
        [String]$region,
        [String]$CWRun
    )
    $path=$workingDirectory
        
    #Latency Average last 24 Hours
    $json = Get-Content ($path+"WorkSpacesHistoricalLatencyTemplate.json") -Raw
    $jsonobj = $json | ConvertFrom-Json
    $data = @("AWS/WorkSpaces","InSessionLatency","WorkspaceId",$WSId)
    $jsonobj.metrics[0]=$data
    $jsonobj = $jsonobj | ConvertTo-Json -depth 6
    set-content -Path  ($path+"WorkSpacesHistoricalLatency.json") -Value $jsonobj
    $JsonFile = Get-Content -Raw -Path ($path+"WorkSpacesHistoricalLatency.json")
    $image= get-CWMetricWidgetImage -MetricWidget $JsonFile -Region $region
    [byte[]]$bytes = $image.ToArray()
    Set-Content -Path ($path+"SelectedWSMetrics\WorkSpacesHistoricalLatency"+$CWRun+".png") -Value $bytes -Encoding Byte

    #Now UDP Packet Loss
    $json = Get-Content ($path+"\WorkSpacesUDPTemplate.json") -Raw
    $jsonobj = $json | ConvertFrom-Json
    $data = @("AWS/WorkSpaces","WorkSpacesUDPPacketLossRate","WorkspaceId",$WSId)
    $jsonobj.metrics[0]=$data
    $jsonobj = $jsonobj | ConvertTo-Json -depth 6
    set-content -Path  ($path+"\WorkSpacesUDP.json") -Value $jsonobj
    $JsonFile = Get-Content -Raw -Path ($path+"\WorkSpacesUDP.json")
    $image= get-CWMetricWidgetImage -MetricWidget $JsonFile -Region $region
    [byte[]]$bytes = $image.ToArray()
    Set-Content -Path ($path+"SelectedWSMetrics\WorkSpacesUDP"+$CWRun+".png") -Value $bytes -Encoding Byte

    #Connection Summary
    $json = Get-Content ($path+"\WorkSpacesConnectionSummaryTemplate.json") -Raw
    $jsonobj = $json | ConvertFrom-Json
    $data = @("AWS/WorkSpaces","ConnectionSuccess","WorkspaceId",$WSId)
    $jsonobj.metrics[0]=$data
    $jsonobj = $jsonobj | ConvertTo-Json -depth 6
    set-content -Path  ($path+"\WorkSpacesConnectionSummary.json") -Value $jsonobj
    $JsonFile = Get-Content -Raw -Path ($path+"\WorkSpacesConnectionSummary.json") 
    $image= get-CWMetricWidgetImage -MetricWidget $JsonFile -Region $region
    [byte[]]$bytes = $image.ToArray()
    Set-Content -Path ($path+"SelectedWSMetrics\WorkSpacesConnectionSummary"+$CWRun+".png") -Value $bytes -Encoding Byte

    $global:CloudWatchImageProcessComplete=$true
    
}
    
function Get-CloudWatchImagesWorkSpaceMetrics(){
    # This function will generate dashboard images for your metrics portal. It targets the additional metrics for WorkSpaces through CloudWatch Agent.  
    param(
        [String]$workingDirectory,
        [String]$WSId,
        [String]$AWSProfile,
        [String]$region,
        [String]$CWRun
    )     

    $path=$workingDirectory
    #Process to get the CPU PNG File
    $json = Get-Content ($path+"\WorkSpacesCPUTemplate.json") -Raw
    $jsonobj = $json | ConvertFrom-Json
    $data = @("AWS/WorkSpaces","CPUUsage","WorkspaceId",$WSId)
    $jsonobj.metrics[0]=$data
    $jsonobj = $jsonobj | ConvertTo-Json -depth 6
    set-content -Path  ($path+"\WorkSpacesCPU.json") -Value $jsonobj
    $JsonFile = Get-Content -Raw -Path ($path+"\WorkSpacesCPU.json")
    $image= get-CWMetricWidgetImage -MetricWidget $JsonFile -Region $region
    [byte[]]$bytes = $image.ToArray()
    Set-Content -Path ($path+"\SelectedWSMetrics\WorkSpacesCPU"+$CWRun+".png") -Value $bytes -Encoding Byte

    #Disk
    $json = Get-Content ($path+"\WorkSpacesDiskTemplate.json") -Raw
    $jsonobj = $json | ConvertFrom-Json
    $data = @("AWS/WorkSpaces","RootVolumeDiskUsage","WorkspaceId",$WSId)
    $jsonobj.metrics[0]=$data
    $jsonobj = $jsonobj | ConvertTo-Json -depth 6
    set-content -Path  ($path+"\WorkSpacesDisk.json") -Value $jsonobj
    $JsonFile = Get-Content -Raw -Path ($path+"\WorkSpacesDisk.json")
    $image= get-CWMetricWidgetImage -MetricWidget $JsonFile -Region $region
    [byte[]]$bytes = $image.ToArray()
    Set-Content -Path ($path+"\SelectedWSMetrics\WorkSpacesDisk"+$CWRun+".png") -Value $bytes -Encoding Byte

    #Memory
    $json = Get-Content ($path+"\WorkSpacesMemoryTemplate.json") -Raw
    $jsonobj = $json | ConvertFrom-Json
    $data = @("AWS/WorkSpaces","MemoryUsage","WorkspaceId",$WSId)
    $jsonobj.metrics[0]=$data
    $jsonobj = $jsonobj | ConvertTo-Json -depth 6
    set-content -Path  ($path+"\WorkSpacesMemory.json") -Value $jsonobj
    $JsonFile = Get-Content -Raw -Path ($path+"\WorkSpacesMemory.json")
    $image= get-CWMetricWidgetImage -MetricWidget $JsonFile -Region $region
    [byte[]]$bytes = $image.ToArray()
    Set-Content -Path ($path+"\SelectedWSMetrics\WorkSpacesMemory"+$CWRun+".png") -Value $bytes -Encoding Byte

    $global:CloudWatchImageProcessCompleteWSMetrics=$true

}
    
function Get-CloudWatchStats(){
    # This function will read your access metrics from CloudTrail and then return object for the GUI to present. 
    param(
        [String]$workingDirectory,
        [String]$WorkSpaceAccessLogs,
        [String]$CloudTrailLogs,
        [String]$WorkSpaceId,
        [String]$AccessLogsRegion,
        [String]$CloudTrailRegion,
        [String]$CWRun
    )

    #Query Logins for the User
    $Startdate = (New-TimeSpan -Start (Get-Date "01/01/1970") -End (Get-Date)).TotalSeconds
    $Startdate = $Startdate - 604800
    $Enddate = Get-Date
    $Enddate=$Enddate.ToUniversalTime()
    $Enddate = (New-TimeSpan -Start (Get-Date "01/01/1970") -End ($Enddate)).TotalSeconds

    #Initiate Query if Access Logs are configured
    if ($WorkSpaceAccessLogs -ne ""){
        $queryString=('fields @message |filter detail.workspaceId="'+$WorkSpaceId+'"') 
        $queryResultUserAccess=Start-CWLQuery -QueryString $queryString  -LogGroupName $WorkSpaceAccessLogs -StartTime $Startdate -EndTime $Enddate -Region $AccessLogsRegion
        $queryResultUserAccessComplete=$false
        $queryString=('fields @message |filter `detail.eventName`="ModifyWorkspaceProperties" |filter `detail.requestParameters.workspaceId`="'+$WorkSpaceId+'"') 
        $queryResultWorkSpaceChanges=Start-CWLQuery -QueryString ('fields @message |filter `detail.eventName`="ModifyWorkspaceProperties"  |filter `detail.requestParameters.workspaceId`="'+$WorkSpaceId+'"') -LogGroupName $WorkSpaceAccessLogs -StartTime $Startdate -EndTime $Enddate -Region $AccessLogsRegion
        $queryResultCloudTrailComplete = $false
    }
    else{
        $queryResultUserAccessComplete = $true
        $queryResultCloudTrailComplete = $true
        
    }

    # Wait for the WorkSpace Access query to complete
    while ($queryResultUserAccessComplete -eq $false){
        foreach ($completeQuerry in $queryStatus){
            if ($completeQuerry.QueryId -eq $queryResultUserAccess){
                $queryResultUserAccessComplete = $true
                # Loop through the records and output to a CSV
                $loadResults=Get-CWLQueryResult -QueryId $queryResultUserAccess -Region $AccessLogsRegion
                $x=0
                $OutputObj = @()
                while ($x-lt $loadResults.Statistics.RecordsMatched){
                    $userAccessLog = $loadResults.Results[$x][0].Value | ConvertFrom-Json
                    $NewObj = New-Object -TypeName PSobject
                    $NewObj | Add-Member -NotePropertyName "Time" -NotePropertyValue $userAccessLog.time
                    $NewObj | Add-Member -NotePropertyName "ClientVersion" -NotePropertyValue $userAccessLog.detail.clientVersion
                    $NewObj | Add-Member -NotePropertyName "IP" -NotePropertyValue $userAccessLog.detail.clientIpAddress
                    $NewObj | Add-Member -NotePropertyName "Platform" -NotePropertyValue $userAccessLog.detail.clientPlatform
                    $OutputObj += $NewObj
                    $x=$x+1
                }
                # Now load the object into a CSV to load into the Form
                $OutputObj | Export-Csv -Path ($workingDirectory+"\SelectedWSMetrics\ConnectionHistory"+$CWRun+".csv") -NoTypeInformation
            }
        }
        sleep -Seconds 1
        $queryStatus =Get-CWLQuery -Status Complete -Region $AccessLogsRegion
    }

    
    $queryStatus =Get-CWLQuery -Status Complete -Region $AccessLogsRegion
    # Wait for the CloudTrail query to complete
    $queryString=('fields @message |filter `detail.eventName`="ModifyWorkspaceProperties" |filter `detail.requestParameters.workspaceId`="'+$WorkSpaceId+'"') 
    $queryResultWorkSpaceChanges=Start-CWLQuery -QueryString ('fields @message |filter `detail.eventName`="ModifyWorkspaceProperties"  |filter `detail.requestParameters.workspaceId`="'+$WorkSpaceId+'"') -LogGroupName $WorkSpaceAccessLogs -StartTime $Startdate -EndTime $Enddate -Region $AccessLogsRegion
    $queryResultCloudTrailComplete = $false
    while ($queryResultCloudTrailComplete -eq $false){
        foreach ($completeQuerry in $queryStatus){
            if ($completeQuerry.QueryId -eq $queryResultWorkSpaceChanges){
                $queryResultCloudTrailComplete = $true
                $x=0
                $OutputObj = @()
                $loadResults=Get-CWLQueryResult -QueryId $queryResultWorkSpaceChanges -Region $AccessLogsRegion
                while ($x-lt $loadResults.Statistics.RecordsMatched){
                    $CloudTrailLog = $loadResults.Results[$x][0].Value | ConvertFrom-Json
                    write-host $CloudTrailLog.detail.requestParameters
                    $NewObj = New-Object -TypeName PSobject
                    $NewObj | Add-Member -NotePropertyName "Time" -NotePropertyValue $CloudTrailLog.detail.eventTime
                    $NewObj | Add-Member -NotePropertyName "Region" -NotePropertyValue $CloudTrailLog.detail.awsRegion
                    $NewObj | Add-Member -NotePropertyName "WorkSpaceId" -NotePropertyValue $CloudTrailLog.detail.requestParameters.workspaceId
                    $NewObj | Add-Member -NotePropertyName "ComputeType" -NotePropertyValue $CloudTrailLog.detail.requestParameters.workspaceProperties.computeTypeName
                    $NewObj | Add-Member -NotePropertyName "RunningMode" -NotePropertyValue $CloudTrailLog.detail.requestParameters.workspaceProperties.runningMode
                    $NewObj | Add-Member -NotePropertyName "UserVolume" -NotePropertyValue $CloudTrailLog.detail.requestParameters.workspaceProperties.userVolumeSizeGib
                    $NewObj | Add-Member -NotePropertyName "RootVolume" -NotePropertyValue $CloudTrailLog.detail.requestParameters.workspaceProperties.rootVolumeSizeGib
                    $OutputObj += $NewObj
                    $x=$x+1
                }
                # Now load the object into a CSV to load into the Form
                $OutputObj | Export-Csv -Path ($workingDirectory+"\SelectedWSMetrics\UserChanges"+$CWRun+".csv") -NoTypeInformation
    
            }
        }
        sleep -Seconds 1
        $queryStatus =Get-CWLQuery -Status Complete -Region $AccessLogsRegion
    }
}  
