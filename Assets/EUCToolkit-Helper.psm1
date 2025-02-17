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
    This script module performs the heavy lifting for the EUC toolkit. This includes building the local db object, optimizing API calls,
    returning error and success messages, generating dashboard images, and other functionality part of the EUC toolkit. For more information,
    see the link below:
    https://github.com/aws-samples/euc-toolkit
#>

# Current WorkSpaces Regions. See link below for current WorkSpaces availability
# https://docs.aws.amazon.com/workspaces/latest/adminguide/azs-workspaces.html


function Get-LocalWorkSpacesDB(){
    param(
        $DeployedRegions,
        $throttleControl
    )
    # This function build a PSObject that contains all of your WorkSpaces information. The object will act as a local DB for the GUI.
    # If you need to have object persistence to save API calls, this function can be replaced with a function that calls your persistent store.
    $WorkSpacesDDB = @()
    # Finds all current WorkSpaces. If Active Directory cannot be reached, those attributes are omitted.
    foreach($DeployedRegion in $DeployedRegions){
        $wksResponse = Get-WKSWorkSpaces -Region $DeployedRegion.Region -DirectoryId $DeployedRegion.DirectoryId -limit 25 -NoAutoIteration -select * -NextToken $null
        $RegionalWks = $wksResponse.Workspaces
        $token = $wksResponse.NextToken
        while ($null -ne $token) {
            $wksResponse = Get-WKSWorkSpaces -Region $DeployedRegion.Region -DirectoryId $DeployedRegion.DirectoryId -limit 25 -NoAutoIteration -select * -NextToken $token
            $RegionalWks += $wksResponse.Workspaces
            $token = $wksResponse.NextToken
            if($throttleControl){
                Start-Sleep -Milliseconds 200
            }
        }
        foreach ($Wks in $RegionalWks){
            $adErr = $false
            $wks | Add-Member -NotePropertyName "Region" -NotePropertyValue $DeployedRegion.Region
            if($Wks.WorkspaceProperties.Protocols -like "WSP"){$wsProto = 'DCV'}elseif ($Wks.WorkspaceProperties.Protocols -like "PCOIP"){$wsProto = 'PCoIP'} else{$wsProto = 'BYOP'}
            $wks | Add-Member -NotePropertyName "Protocol" -NotePropertyValue $wsProto
            $wks | Add-Member -NotePropertyName "RegCode" -NotePropertyValue ($DeployedRegion | Where-Object {$_.directoryId -eq $Wks.directoryId}).RegistrationCode
            try{
                $ADUser = Get-ADUser -Identity $Wks.UserName -Properties "EmailAddress"
            }catch{
                $adErr = $true
            }
            if($adErr -eq $false){
                $wks | Add-Member -NotePropertyName "FirstName" -NotePropertyValue ($ADUser.GivenName)
                $wks | Add-Member -NotePropertyName "LastName" -NotePropertyValue ($ADUser.Surname)
                $wks | Add-Member -NotePropertyName "Email" -NotePropertyValue ($ADUser.EmailAddress)
            }else{
                $wks | Add-Member -NotePropertyName "FirstName" -NotePropertyValue "AD Info Not Available"
                $wks | Add-Member -NotePropertyName "LastName" -NotePropertyValue "AD Info Not Available"
                $wks | Add-Member -NotePropertyName "Email" -NotePropertyValue "AD Info Not Available"
            }
            $WorkSpacesDDB += $wks
        }
    }
    return $WorkSpacesDDB
}

function Get-WksServiceQuotasDB(){
    param(
        $DeployedRegions
    )
    $WSServiceQuota = @()
    $UniqueRegions = $DeployedRegions | Select-Object Region -Unique
    foreach($WksRegion in $UniqueRegions){
        $region = $WksRegion.Region
        $DeployedRegionsTemp = New-Object -TypeName PSobject
        $DeployedRegionsTemp | Add-Member -NotePropertyName "Region" -NotePropertyValue $region
        # For more information on WorkSpaces Quotas, see https://docs.aws.amazon.com/workspaces/latest/adminguide/workspaces-limits.html.
        #Total Regional WorkSpaces
        try{
            $TotalQuota = (Get-SQServiceQuota -ServiceCode workspaces -QuotaCode "L-34278094" -Region $region).Value
        }catch{
            $TotalQuota = "N/A"
        }
        $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaWks" -NotePropertyValue $TotalQuota
        #Total Regional General Purpose 4XL WorkSpaces
        try{
            $Gp4xlQuota = (Get-SQServiceQuota -ServiceCode workspaces -QuotaCode "L-465DA8AF" -Region $region).Value
        }catch{
            $Gp4xlQuota = "N/A"
        }
        $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaWks4xl" -NotePropertyValue $Gp4xlQuota
        #Total Regional General Purpose 8XL WorkSpaces
        try{
            $Gp8xlQuota = (Get-SQServiceQuota -ServiceCode workspaces -QuotaCode "L-C266A5F4" -Region $region).Value
        }catch{
            $Gp8xlQuota = "N/A"
        }
        $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaWks8xl" -NotePropertyValue $Gp8xlQuota
        # StandBy WorkSpaces
        try{
            $StandbyQuota = (Get-SQServiceQuota -ServiceCode workspaces -QuotaCode "L-9A67B5CB" -Region $region).Value
        }catch{
            $StandbyQuota = "N/A"
        }
        $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaStandby" -NotePropertyValue $StandbyQuota
        # GraphicsPro WorkSpaces
        try{
            $GraphicsProQuota = (Get-SQServiceQuota -ServiceCode workspaces -QuotaCode "L-254B485B" -Region $region).Value
        }catch{
            $GraphicsProQuota = "N/A"
        }
        $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaGraphicsPro" -NotePropertyValue $GraphicsProQuota
        # Graphics.g4dn WorkSpaces
        try{
            $GraphicsG4Quota = (Get-SQServiceQuota -ServiceCode workspaces -QuotaCode "L-BCACAEBC" -Region $region).Value
        }catch{
            $GraphicsG4Quota = "N/A"
        }
        $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaG4dn" -NotePropertyValue $GraphicsG4Quota
        # Graphics.g4dn WorkSpaces Pro
        try{
            $GraphicsG4ProQuota = (Get-SQServiceQuota -ServiceCode workspaces -QuotaCode "L-BE9A8466" -Region $region).Value
        }catch{
            $GraphicsG4ProQuota = "N/A"
        }
        $DeployedRegionsTemp | Add-Member -NotePropertyName "quotaG4dnPro" -NotePropertyValue $GraphicsG4ProQuota

        $WSServiceQuota += $DeployedRegionsTemp
    }
    return $WSServiceQuota
}

function Get-WksDirectories(){
    param(
        $throttleControl
    )
    $regions = @('us-east-1','us-west-2', 'ap-south-1', 'ap-northeast-2', 'ap-southeast-1', 'ap-southeast-2', 'ap-northeast-1', 'ca-central-1', 'eu-central-1','eu-west-1', 'eu-west-2', 'sa-east-1')
    $DeployedDirectories = @()
    $personalDirectory = New-Object -TypeName Amazon.WorkSpaces.Model.DescribeWorkspaceDirectoriesFilter
    $personalDirectory.Name = "WORKSPACE_TYPE"
    $personalDirectory.Values += "PERSONAL"
    # Find regions that have WorkSpaces deployments
    foreach($region in $regions){
        $directoryResponse = Get-WKSWorkspaceDirectories -Region $region -limit 25 -Filter $personalDirectory -NoAutoIteration -select * -NextToken $null
        $RegionDirectories = $directoryResponse.Directories
        $token = $directoryResponse.NextToken
        while ($null -ne $token) {
            $directoryResponse = Get-WKSWorkspaceDirectories -Region $region -limit 25 -Filter $personalDirectory -NoAutoIteration -select * -NextToken $token
            $RegionDirectories += $directoryResponse.Directories
            $token = $directoryResponse.NextToken
            if($throttleControl){
                Start-Sleep -Milliseconds 200
            }
        }
        if($RegionDirectories){
            foreach($WksDirectory in $RegionDirectories){
                $WksDirectory | Add-Member -NotePropertyName "Region" -NotePropertyValue $region
                $subnetA = Get-EC2Subnet -SubnetId $WksDirectory.SubnetIds[0] -Region $region
                $subnetB = Get-EC2Subnet -SubnetId $WksDirectory.SubnetIds[1] -Region $region
                $dirAvailableIPs = $subnetA.AvailableIpAddressCount + $subnetB.AvailableIpAddressCount
                $WksDirectory | Add-Member -NotePropertyName "DirectoryAvailableIPs" -NotePropertyValue $dirAvailableIPs
                $DeployedDirectories += $WksDirectory
            }
        }else{
            Write-Host "Skipping $region"
        }
    }
    return $DeployedDirectories
}

function Get-AllBundles(){
    param(
        $Regions,
        $Custom,
        $throttleControl
    )
    $Bundles = @()
    foreach($Region in $Regions.Region){
        if($custom -eq $true){
            $bundleResponse = Get-WKSWorkspaceBundle -Region $Region -NoAutoIteration -select * -NextToken $null
            $Bundles += $bundleResponse.Bundles
            $token = $bundleResponse.NextToken
            while ($null -ne $token) {
                $bundleResponse = Get-WKSWorkspaceBundle -Region $Region -NoAutoIteration -select * -NextToken $token
                $Bundles += $bundleResponse.Bundles
                $token = $bundleResponse.NextToken
                if($throttleControl){
                    Start-Sleep -Milliseconds 200
                }
            }
            $bundleResponse = Get-WKSWorkspaceBundle -Region $Region -Owner 'AMAZON' -NoAutoIteration -select * -NextToken $null
            $Bundles += $bundleResponse.Bundles
            $token = $bundleResponse.NextToken
            while ($null -ne $token) {
                $bundleResponse = Get-WKSWorkspaceBundle -Region $Region -Owner 'AMAZON' -NoAutoIteration -select * -NextToken $token
                $Bundles += $bundleResponse.Bundles
                $token = $bundleResponse.NextToken
                if($throttleControl){
                    Start-Sleep -Milliseconds 200
                }
            }
        }else{
            $bundleResponse = Get-WKSWorkspaceBundle -Region $Region -Owner 'AMAZON' -NoAutoIteration -select * -NextToken $null
            $Bundles += $bundleResponse.Bundles
            $token = $bundleResponse.NextToken
            while ($null -ne $token) {
                $bundleResponse = Get-WKSWorkspaceBundle -Region $Region -Owner 'AMAZON' -NoAutoIteration -select * -NextToken $token
                $Bundles += $bundleResponse.Bundles
                $token = $bundleResponse.NextToken
                Start-Sleep -Milliseconds 200
            }
        }
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
        if ( -not ([string]::IsNullOrEmpty( $requestedSize ))){
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
                $msg = "$WorkSpaceId was unable to extend its Root Volume to $requestedSize GB since its not greater than or equal to 175."
                $logging = New-Object -TypeName PSobject
                $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error During Root Increase"
                $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
            }
        }
    }else{
        if ( -not ([string]::IsNullOrEmpty( $requestedSize ))){
            $msg = "$WorkSpaceId was unable to extend its Root Volume to $requestedSize GB since your User Volume is not equal to 100GB."
            $logging = New-Object -TypeName PSobject
            $logging | Add-Member -NotePropertyName "ErrorCode" -NotePropertyValue "Error During Root Increase"
            $logging | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $msg
        }
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
    $Enddate = $Enddate.ToUniversalTime()
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
        $queryStatus = Get-CWLQuery -Status Complete -Region $AccessLogsRegion
    }

    $queryStatus = Get-CWLQuery -Status Complete -Region $AccessLogsRegion
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

###############################################
# ! # ! # WorkSpaces Pools Helper # ! # ! # 
###############################################

function Get-WksPoolsDirectories(){
    param(
        $throttleControl
    )
    $regions = @('us-east-1','us-west-2', 'ap-south-1', 'ap-northeast-2', 'ap-southeast-1', 'ap-southeast-2', 'ap-northeast-1', 'ca-central-1', 'eu-central-1','eu-west-1', 'eu-west-2', 'sa-east-1')
    $poolsDirectory = New-Object -TypeName Amazon.WorkSpaces.Model.DescribeWorkspaceDirectoriesFilter
    $poolsDirectory.Name = "WORKSPACE_TYPE"
    $poolsDirectory.Values += "POOLS"
    $DeployedPoolsDirectories = @()
    # Find regions that have WorkSpaces deployments
    foreach($region in $regions){
        $directoryResponse = Get-WKSWorkspaceDirectories -Region $region -limit 25 -Filter $poolsDirectory -NoAutoIteration -select * -NextToken $null
        $RegionsCall = $directoryResponse.Directories
        $token = $directoryResponse.NextToken
        while ($null -ne $token) {
            $directoryResponse = Get-WKSWorkspaceDirectories -Region $region -limit 25 -Filter $poolsDirectory -NoAutoIteration -select * -NextToken $token
            $RegionsCall += $directoryResponse.Directories
            $token = $directoryResponse.NextToken
            if($throttleControl){
                Start-Sleep -Milliseconds 200
            }
        }
        if($RegionsCall){
            foreach($PoolsRegion in $RegionsCall){
                $DeployedDirectoriesTemp = New-Object -TypeName PSobject
                $DeployedDirectoriesTemp | Add-Member -NotePropertyName "Region" -NotePropertyValue $region
                $DeployedDirectoriesTemp | Add-Member -NotePropertyName "RegistrationCode" -NotePropertyValue $PoolsRegion.RegistrationCode
                $DeployedDirectoriesTemp | Add-Member -NotePropertyName "DirectoryId" -NotePropertyValue $PoolsRegion.DirectoryId
                $DeployedDirectoriesTemp | Add-Member -NotePropertyName "DirectoryName" -NotePropertyValue $PoolsRegion.DirectoryName
                $DeployedDirectoriesTemp | Add-Member -NotePropertyName "DirectoryAlias" -NotePropertyValue $PoolsRegion.Alias
                $DeployedDirectoriesTemp | Add-Member -NotePropertyName "DirectoryType" -NotePropertyValue $PoolsRegion.Type
                $DeployedDirectoriesTemp | Add-Member -NotePropertyName "DirectoryState" -NotePropertyValue $PoolsRegion.State
                $DeployedDirectoriesTemp | Add-Member -NotePropertyName "DirectoryUserEnabledAsLocalAdministrator" -NotePropertyValue $PoolsRegion.WorkspaceCreationProperties.UserEnabledAsLocalAdministrator
                $subnetA = Get-EC2Subnet -SubnetId $PoolsRegion.SubnetIds[0] -Region $region
                $subnetB = Get-EC2Subnet -SubnetId $PoolsRegion.SubnetIds[1] -Region $region
                $dirAvailableIPs = $subnetA.AvailableIpAddressCount + $subnetB.AvailableIpAddressCount
                $DeployedDirectoriesTemp | Add-Member -NotePropertyName "DirectoryAvailableIPs" -NotePropertyValue $dirAvailableIPs
                $DeployedPoolsDirectories += $DeployedDirectoriesTemp
            }
        }else{
            Write-Host "Skipping $region"
        }
    }
    return $DeployedPoolsDirectories
}

function Get-WksPools(){
    param(
        $DeployedRegions,
        $bundles,
        $throttleControl
    )

    $poolsDB = @()
    foreach($region in $DeployedRegions.Region){
        $poolsResponse = Get-WKSWorkspacesPool -Region $region -limit 25 -NoAutoIteration -select * -NextToken $null
        $pools = $poolsResponse.WorkspacesPools
        $token = $poolsResponse.NextToken
        while ($null -ne $token) {
            $poolsResponse = Get-WKSWorkspacesPool -Region $region -limit 25 -NoAutoIteration -select * -NextToken $token
            $pools += $poolsResponse.WorkspacesPools
            $token = $poolsResponse.NextToken
            if($throttleControl){
                Start-Sleep -Milliseconds 200
            }
        }
        $poolsDB += $pools
    }
    return $poolsDB
}


function Import-WksPoolsSessions(){
    # Description
    param(
        [String]$poolId,
        [String]$region,
        $throttleControl
    )

    $poolsSessions = @()
    $poolsResponse = Get-WKSWorkspacesPoolSession -PoolId $poolId -Region $region -limit 25 -NoAutoIteration -select * -NextToken $null
    $poolsSessions = $poolsResponse.Sessions
    $token = $poolsResponse.NextToken
    while ($null -ne $token) {
        $poolsResponse = Get-WKSWorkspacesPoolSession -PoolId $poolId -Region $region -limit 25 -NoAutoIteration -select * -NextToken $token
        $poolsSessions += $poolsResponse.Sessions
        $token = $poolsResponse.NextToken
        if($throttleControl){
            Start-Sleep -Milliseconds 200
        }
    }
    return $poolsSessions
}


###############################################
# ! # ! # AppStream Helper # ! # ! # 
###############################################

function Import-AppStreamRegions(){
    param(
        $throttleControl
    )
    #description
    $responseStacks = @()
    $regions = @('us-east-1', 'us-east-2', 'us-west-2', 'ap-south-1', 'ap-northeast-2', 'ap-southeast-1', 'ap-southeast-2', 'ap-northeast-1', 'ca-central-1', 'eu-central-1','eu-west-1', 'eu-west-2')
    foreach ($region in $regions){
        $tempStacks = $null
        $tempStacks = Get-APSStackList -Region $region -NoAutoIteration -select * -NextToken $null
        if(($tempStacks.Stacks).Count -ne 0){
            $responseStacks += $tempStacks.Stacks
            $token = $tempStacks.NextToken
            while ($null -ne $token) {
                $tempStacks = Get-APSStackList -Region $region -NoAutoIteration -select * -NextToken $token
                $responseStacks += $tempStacks.Stacks
                $token = $tempStacks.NextToken
                if($throttleControl){
                    Start-Sleep -Milliseconds 200
                }
            }
        }else{
            Write-Host "Skipping $region"
        }
    }
    return $responseStacks
}

function Import-AppStreamSessions(){
    # Description
    param(
        [String]$stackName,
        [String]$fleetName,
        [String]$region,
        $capacityInfo,
        $throttleControl
    )

    $sessionList = @()
    # Get fleet and check capacity
    $sessionResponse = Get-APSSessionList -StackName $stackName -FleetName $fleetName -Region $region -limit 50 -NoAutoIteration -select * -NextToken $null
    if (($sessionResponse.Sessions).Count -ne 0){
        $sessionList = $sessionResponse.Sessions
        $token = $sessionResponse.NextToken
        while ($null -ne $token) {
            $sessionResponse += Get-APSSessionList -StackName $stackName -FleetName $fleetName -Region $region -limit 50 -NoAutoIteration -select * -NextToken $sessionResponse.NextToken
            $sessionList += $sessionResponse.Sessions
            $token = $sessionResponse.NextToken
            if($throttleControl){
                Start-Sleep -Milliseconds 200
            }
        }
    }
    return $sessionList
}