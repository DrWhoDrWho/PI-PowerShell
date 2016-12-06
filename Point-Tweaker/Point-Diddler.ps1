<#
.SYNOPSIS
Point-Diddler: Edit Points to test Point-Tweaker PowerShell Script

.DESCRIPTION
Do not even THINK about running this in a producion enviornment
    without evaluating it in development and test environments first.

This script edits points to have PointSource=Base
So that the Point-Tweaker PowerShell Script has something to do
.NOTES
Copyright 2016 Paul H. Gusciora - OSIsoft, LLC 

Licensed under the Apache License, Version 2.0 (the "License"); 
you may not use this file except in compliance with the License. 
You may obtain a copy of the License at 
    <http://www.apache.org/licenses/LICENSE-2.0> 
Unless required by applicable law or agreed to in writing, software 
distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. 
See the License for the specific language governing permissions and 
limitations under the License. 

Additional restrictions: 
1. If this code works for you, great! If it does not, Oh Well!
2. Do not even THINK about running this in a producion enviornment
     without evaluating it in development and test environments first.

#>
<#
Change Log

Version 0.80 2016-10-10
Version 0.90 2016-11-23 improve output format
Version 0.91 2016-12-05 Reorganize into multi-line quick-help comments
Version 0.92 2016-12-05 add license
#>
param (
    [Parameter(Position=0, Mandatory=$false)]
    # default PI server for testing
    [string]$PIServerName = "DellDSCsPI", 
    # Maximum delay between loop executions [seconds]
    # Average time will be half of that
    [Parameter(Position=1,Mandatory=$false)]
    [ValidateRange(2,1800)][int]$Wait = 15
)

[string]$t_name      = "Point-Tweaker"

# early versions of OSISoft.PowerShell can not edit PointType
$h_PointEdit = If ( $false ) { @{PointClass="Base";PointType="Float64"} } else { @{PointClass="Base"} }
# $h_PointEdit | get-member
[string[]]$pt_extension = @(".test00", ".test01", ".test02", ".test03", ".test04", ".test05", ".test06", ".test07", ".test08", ".test09", ".test10", ".test11", ".test12", ".test13", ".test14"  )


    # Get the PI Server object
    $PIServer = Get-PIDataArchiveConnectionConfiguration $PIServerName -ErrorAction Stop
    # connect to PI Server
    $con = Connect-PIDataArchive -PIDataArchiveConnectionConfiguration $PIServer -ErrorAction Stop

Try
    {
    do
        {
        # Construct a randomly selected test point name
        # point names consist of a base and an extension.
        $k = random -Minimum 0 -Maximum ($pt_extension.Length - 1)
        [string]$pt_name = $t_name + $pt_extension[$k]
        # sleep a random amount of time between 0 and the maximum selected.
        $sleep_time = random -Minimum 0 -Maximum $Wait
        #
        Write-Host ("{0} {1} {2} {3} {4}" -f "Edit: ", $pt_name, "; Sleep ", $sleep_time, " s" )
        #
        Set-PIPoint -connection $con -Name ($pt_name) -Attributes $h_PointEdit
        Start-Sleep -seconds $sleep_time
        }
    while($true)
    }
finally
    {
    # allow for a graceful exit to Stop, Control-C, or Stop Service.
    }