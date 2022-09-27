## Table of contents

- [Solution Overview](#solution-overview)
- [Getting Started](#getting-started)
- [Customizing the Solution](#customizing-the-solution)
- [File Structure](#file-structure)
- [License](#license)

<a name="solution-overview"></a>
# Solution Overview
The EUC Toolkit was created to provide additional features for admins managing [Amazon WorkSpaces](https://docs.aws.amazon.com/workspaces/latest/adminguide/amazon-workspaces.html) and [Amazon AppStream 2.0](https://docs.aws.amazon.com/appstream2/latest/developerguide/what-is-appstream.html) environments. The current build offers the listed functionality below. For additional information, please see our [launch announcement](https://aws.amazon.com/blogs/desktop-and-application-streaming/euc-toolkit/).

### Amazon WorkSpaces
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

### Amazon AppStream 2.0
- Query and display active sessions
- Filter active sessions by:
    - Stack, connected state, userId, session state, IP address, and/or Region
- View in-use IP of active sessions
- Terminate active sessions
- Export AppStream 2.0 report (CSV)
- Optional functionality:
    - Windows Remote Assistance

### Overall Toolkit
- API logging
- Source permissions identifier (supports instance profiles) 


<a name="getting-started"></a>
# Getting Started
For information on getting the solution setup, along with the steps for optional features, see our [launch announcement](https://aws.amazon.com/blogs/desktop-and-application-streaming/euc-toolkit/).


<a name="aws-solutions-constructs"></a><a name="customizing-the-solution"></a>
# Customizing the Solution
The solution can be completely customized to meet business needs. The EUC toolkit is built on PowerShell using the Windows Presentation Framework ([WPF](https://learn.microsoft.com/en-us/visualstudio/designers/getting-started-with-wpf?view=vs-2022)) to display a graphical user interface (GUI). In addition, the solution has been modularized to to allow for changes and customizations. As is, the toolkit is licensed as MIT-0, meaning it is an 'as-is' example. Any changes made to the project are owned by the modifier. 

**Customizing the PowerShell GUI**

The GUI for the application is built using XML. To add additional buttons, labels, etc., open up the EUCToolkit-MainGUI.xml and make modifications there. Below is a sample button that is defined for changing the running mode of a WorkSpace. From there, the button can be referenced in your PowerShell script and have invocation actions configured. 

`
"Button Name="btnChangeRunningMode" Content="Change RunningMode" HorizontalAlignment="Left" Height="24" Margin="41,575,0,0" VerticalAlignment="Top" Width="124" RenderTransformOrigin="0.671,0.467" Grid.Column="1" Grid.ColumnSpan="2"
`

**Creating / Customizing Functions**

The Powershell script is divided into 2 files, both can be customized to add additional functionality or used as a reference for other automation.

**EUCToolkit-Main.ps1**

Contains all of the code that interacts with the GUI. The functions in this script will call actions in the EUCToolkit-Helper.psm1 to perform calls against WorkSpaces and on AppStream 2.0.

**EUCToolkit-Helper.psm1**

Contains all of the functions that interact with WorkSpaces and AppStream 2.0.

**Updating CloudWatch Images**

Included in the EUC toolkit are several JSON files that are used to generate images from CloudWatch. These can be customized, see this [documentation](https://docs.aws.amazon.com/AmazonCloudWatch/latest/monitoring/CloudWatch-metric-streams-formats-json.html) for additional information.

<a name="file-structure"></a>
# File structure

<pre>
|-Assets/
  |-CWHelper/
    |-SelectedWSMetrics/
    |-WorkSpaceHistoricalLatency-Start.png
    |-WorkSpaceHistoricalLatencyTemplate.JSON
    |-WorkSpaceLatency-Start.png
    |-WorkSpaceLatencyTemplate.JSON
    |-WorkSpacesSessionLaunch.png
    |-WorkSpacesSessionLaunchTemplate.JSON
    |-WorkSpacesCPU-Start.png
    |-WorkSpacesCPUTemplate.JSON
    |-WorkSpacesDisk.png
    |-WorkSpacesDiskTemplate.json
    |-WorkSpacesMemory.png
    |-WorkSpacesMemoryTemplate.json
  |-EUCToolkit-Helper.psm1
  |-EUCToolkit-MainGUI.xml
|-EUCToolkit-Main.ps1
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