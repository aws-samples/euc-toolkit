## Table of contents

- [Solution Overview](#solution-overview)
- [Getting Started](#getting-started)
- [Customizing the Solution](#customizing-the-solution)
- [File Structure](#file-structure)
- [License](#license)

<a name="solution-overview"></a>
# Solution Overview
The EUC Toolkit was created to provide additional features for admins managing [Amazon WorkSpaces Personal](https://docs.aws.amazon.com/workspaces/latest/adminguide/amazon-workspaces.html), [Amazon WorkSpaces Pools](https://docs.aws.amazon.com/workspaces/latest/adminguide/managing-wsp-pools.html), and [Amazon AppStream 2.0](https://docs.aws.amazon.com/appstream2/latest/developerguide/what-is-appstream.html) environments. The current build offers the listed functionality below. For additional information, please see our [launch announcement](https://aws.amazon.com/blogs/desktop-and-application-streaming/euc-toolkit/).

### Amazon WorkSpaces Personal
- Search by any attribute
    - First name, last name, computer name, WorkSpace ID, bundle ID, running mode, email, username, Region, and/or directory ID
- Bulk or single calls for start, stop, migrate, rebuild, restore, enable and disable admin maintenance (APIs optimized).
- Global WorkSpaces inventory visibility 
- Export WorkSpaces report (CSV)
- Optional functionality:
    - Amazon CloudWatch metrics (service and OS level metrics)
    - AWS CloudTrail modification history
    - WorkSpaces access history
    - Windows Remote Assistance
    - Remote backup
    - Remote server-side log gathering

### Amazon WorkSpaces Pools
- Query and display active sessions
- Filter active sessions by:
    - Stack, connected state, userId, session state, IP address, and/or Region
- View in-use IP of active sessions
- Terminate active sessions
- Export report (CSV)
- Active session and host inventory visibility
- Optional functionality:
    - Windows Remote Assistance

### Amazon AppStream
- Query and display active sessions
- Filter active sessions by:
    - Stack, connected state, userId, session state, IP address, and/or Region
- View in-use IP of active sessions
- Terminate active sessions
- Export report (CSV)
- Active session and fleet inventory visibility
- Optional functionality:
    - Windows Remote Assistance

### Overall Toolkit
- API logging
- Source permissions identifier (supports instance profiles) 
- Regional service quotas visibility


<a name="getting-started"></a>
# Getting Started
For information on getting the solution setup, along with the steps for optional features, see our [launch announcement](https://aws.amazon.com/blogs/desktop-and-application-streaming/euc-toolkit/).

# Installing AWS Tools
The EUC Toolkit relies on [AWS Tools for PowerShell version 4](https://docs.aws.amazon.com/powershell/latest/userguide/v4migration.html#migrate-select) to use the `-Select` attribute. To install the required modules, you can utilize the following command:

`Install-AWSToolsModule AWS.Tools.Common,AWS.Tools.EC2,AWS.Tools.Workspaces,AWS.Tools.Appstream,AWS.Tools.Cloudwatch,AWS.Tools.CloudwatchLogs,AWS.Tools.ServiceQuotas -CleanUp
`

To upgrade your installed AWS Tools for PowerShell modules, you can utilize the following command:

`
Update-AWSToolsModule -CleanUp -Force
`

<a name="aws-solutions-constructs"></a><a name="customizing-the-solution"></a>
# Customizing the Solution
## Updating supported Regions
By default, the EUC Toolkit will call all commercial regions to build its local database. It is recommended that you update the toolkit to call only the regions you are deployed in. The regional calls are in three functions within `EUCToolkit-helper.psm1`:
- `Get-WksDirectories`
- `Get-WksPoolsDirectories`
- `Import-AppStreamRegions`

### Example
**Default Get-WksDirectories Regions call**

`$regions = @('us-east-1','us-west-2', 'ap-south-1', 'ap-northeast-2', 'ap-southeast-1', 'ap-southeast-2', 'ap-northeast-1', 'ca-central-1', 'eu-central-1','eu-west-1', 'eu-west-2', 'sa-east-1')`

**Updated Get-WksDirectories to call only us-east-1 and us-west-2**

`$regions = @('us-east-1','us-west-2')`

## Customizing Functionality
The solution can be completely customized to meet business needs. The EUC toolkit is built on PowerShell using the Windows Presentation Framework ([WPF](https://learn.microsoft.com/en-us/visualstudio/designers/getting-started-with-wpf?view=vs-2022)) to display a graphical user interface (GUI) on Windows machines. In addition, the solution has been modularized to to allow for changes and customizations. As is, the toolkit is licensed as MIT-0, meaning it is an 'as-is' example. Any changes made to the project are owned by the modifier. 

**Customizing the PowerShell GUI**

The GUI for the application is built using XML. To add additional buttons, labels, etc., open up the EUCToolkit-MainGUI.xml and make modifications there. Below is a sample button that is defined for changing the running mode of a WorkSpace. From there, the button can be referenced in your PowerShell script and have invocation actions configured. 

`
"Button Name="btnChangeRunningMode" Content="Change RunningMode" HorizontalAlignment="Left" Height="24" Margin="41,575,0,0" VerticalAlignment="Top" Width="124" RenderTransformOrigin="0.671,0.467" Grid.Column="1" Grid.ColumnSpan="2"
`

**Creating / Customizing Functions**

The Powershell script is divided into 2 files, both can be customized to add additional functionality or used as a reference for other automation.

**Start-EUCToolkit.ps1**

Contains all of the code that interacts with the GUI. The functions in this script will call actions in the EUCToolkit-Helper.psm1 to perform calls against WorkSpaces and on AppStream.

**EUCToolkit-Helper.psm1**

Contains all of the functions that interact with WorkSpaces and AppStream.

**Settings.csv**

Contains paths for reporting, logs collection, tools, and CloudWatch Log groups.

**Updating CloudWatch Images**

Included in the EUC toolkit are several JSON files that are used to generate images from CloudWatch. These can be customized, see this [documentation](https://docs.aws.amazon.com/AmazonCloudWatch/latest/monitoring/CloudWatch-metric-streams-formats-json.html) for additional information.

**AWS Identity and Access Management (IAM) permissions**

You must have IAM permissions to call the service APIs. It is a best practice to follow the principle of least privilege. The following policy provides access to APIs needed by to the toolkit. If you do not plan to use the WorkSpaces CloudWatch functionality, you may remove WorkSpacesCloudWatchImages and WorkSpacesCloudWatchMetrics from the policy

<a name="Required-permissions"></a>
# Required permissions
- Get-WKSWorkspaceBundle
- Get-WKSWorkspace
- Get-WKSWorkspaceDirectories
- Get-SQServiceQuota
- Get-WKSWorkspaceSnapshot
- Start-WKSWorkspaceMigration
- Edit-WKSWorkspaceProperty
- Edit-WKSWorkspaceState
- Restart-WKSWorkspace
- Restore-WKSWorkspace
- Reset-WKSWorkspace
- Start-WKSWorkspace
- Stop-WKSWorkspace 
- Remove-WKSWorkspace
- Get-WKSWorkspacesPool 
- Get-WKSWorkspacesPoolSession
- Remove-WKSWorkspacesPoolSession
- Get-APSFleetList
- Get-APSStackList
- Revoke-APSSession
- Start-CWLQuery
- Get-CWLQueryResult
- Get-CWMetricWidgetImage

<a name="file-structure"></a>
# File structure

<pre>
|-Assets/
  |-CWHelper/
    |-SelectedWSMetrics/
    |-WorkSpacesConnectionSummaryTemplate.json
    |-WorkSpacesUDPPacketLoss-Start.png
    |-WorkSpacesUDPTemplate.JSON
    |-WorkSpacesHistoricalLatency-Start.png
    |-WorkSpacesHistoricalLatencyTemplate.JSON
    |-WorkSpacesCPU-Start.png
    |-WorkSpacesCPUTemplate.JSON
    |-WorkSpacesDisk-Start.png
    |-WorkSpacesDiskTemplate.JSON
    |-WorkSpacesMemory-Start.png
    |-WorkSpacesMemoryTemplate.JSON
    |-WorkSpacesSessionLaunch-Start.png
    |-WorkSpacesSessionLaunchTemplate.JSON
  |-EUCToolkit-Helper.psm1
  |-EUCToolkit-MainGUI.xml
|-Start-EUCToolkit.ps1
|-CONTRIBUTING.md
|-LICENSE.txt
|-NOTICE.txt
|-README.md
</pre>


################################################

<a name="license"></a>
# License

This library is licensed under the MIT-0 License. See the LICENSE file.
See license [here](https://github.com/aws-samples/euc-toolkit/blob/main/LICENSE).