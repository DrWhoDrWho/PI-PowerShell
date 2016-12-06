<#
.SYNOPSIS
Point-Tweaker: Tweak Points created by Connectors

Monitors the PI message log for PI Point creation
checks if PI PointPource matches one of a set of strings,
if PI PointClass = Base, changes PI PointClass to Classic.

.DESCRIPTION
Do not even THINK about running this in a producion enviornment
    without evaluating it in development and test environments first.

The current version of PI Connectors set:
    PointSource = Prefix.DataSourceName
    for instance, UFL connector uses PointSource = "UFL.Example1"
    PointClass = Base
    PointType  = Float64 for fractional numbers

This script reads PI Message logs based on: PowerShell sample code ListPIMessages.ps1
and incorporates some ideas from PowerShell sample code Connections.ps1
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
Disclaimer: 
All sample code is provided by the author for illustrative purposes only.
These examples have not been thoroughly tested under all conditions.
The author provides no guarantee nor implies any reliability, 
serviceability, or function of these programs.
All programs contained herein are provided to you "As Is" 
without any warranties of any kind. All warranties including 
the implied warranties of non-infringement, merchantability
and fitness for a particular purpose are expressly disclaimed.

*******
Environment setup required:

From a PowerShell script or PowerShell execution window with elevated (Administrator) priviledge,
Use the following command to create the event log and source:
    new-eventlog -LogName $t_logname -Source $t_logsource
where $t_logname and $t_logsource have the values used in the script

To have Windows task scheduler run this script,
configure a task with an action to run the 32-bit version of PowerShell
    Program/Script: C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe
    Arguments:      –NoProfile -NoBanner -NonInteractive -ExecutionPolicy Bypass –File path-to-this-script\Point-Tweaker.ps1
#>
<#
 Change Log

 Version 0.10 2016-08-31
 Version 0.20 2016-09-01
 Version 0.30 2016-09-19
 Version 0.40 2016-09-21 Change names of points
 Version 0.50 2016-09-23
 Version 0.55 2016-09
 Version 0.60 2016-09-28
 Version 0.62 2016-09-29 multiple point-sources
 Version 0.70 2016-10-06 
 Version 0.71 2016-10-28 clean up comments
 Version 0.80 2016-10-28 handle chunks of messages instead of everything since last run
 Version 0.81 2016-10-28 correct RegEx error that did not match created points
 Version 0.82 2016-11-08 Add Start-Trascript
 Version 0.83 2016-11-23 Add $MatchEdited
 Version 0.84 2016-12-02 Reorganize into multi-line quick-help comments
 Version 0.85 2016-12-02 Add license text
#>
param (
    # default PI server for testing
    [Parameter(Position=0, Mandatory=$false)]
    [string]$PIServerName = "DellDSCsPI", 

    # delay between loop executions [seconds]
    [Parameter(Position=1,Mandatory=$false)]
    [ValidateRange(2,1800)][int]$Wait = 3,


    # Time interval message retrieval chuncks [seconds]
    [Parameter(Position=2,Mandatory=$false)]
    [ValidateRange(600,86400)][int]$MessageChunk = 600

)
start-transcript
<#
There is no way to detect the running OSIsoft.PowerShell Library version,
    so $OSI_PS_Ver is set here according to the environment.
Later statements in the code check this variable to use alternate paths
    depending on capability available in that version

OSI_PS_Ver is set assuming that each element of the version can vary between 0 and 99
For Version m.n.p.q :
    $OSI_PS_Ver = 100* (100 * ( 100 * m + n ) + P ) + q
For Version 2.1.0.5	[int]$OSI_PS_Ver = 2010005
For Version 3.0.0.1	[int]$OSI_PS_Ver = 3000001
#>
[int]$OSI_PS_Ver = 2010005
<#
set $ExitNormally to $true for normal processing
set $ExitNormally to $false to test AF Analytic monitoring of this script acting as a Watch-Dog by causing the State to NOT be set to Stopped
Then, when this script exits at a state other than Stopped, the AF Analytic will eventually set the Status to TimeOut and then Failed.
#>
$ExitNormally = $true
<#
set $MatchEdited to $true to match messages that indicate editing points in addition to creating points (useful for testing)
set $MatchEdited to $false to match messages that indicate creating points only
#>
$MatchEdited = $true
# 
# Hashtable containing the PointSource names to monitor
$h_PointSource = @{ "Point-Tweaker.data" = 1 ; "Point-Tweaker" = 2 }
#
# Script name used in messages and to build PI Point tagnames
[string]$t_name      = "Point-Tweaker"
# Windows Log Name to use.
# This script could write to the Windows Application Log, but using a different log makes it easier to view messages.
[string]$t_logname   = "OSIsoft.PowerShell"
# Windows Log Source could be this script name, or could be something else...
[string]$t_logsource = $t_name
# Constants representing PI Point TagNames
[string]$t_State     = $t_name + ".state"    # Processing state in state-transition diagram. Time-stamp checked by AF analytic to make certain the script is running.
[string]$t_Messages  = $t_name + ".messages" # Messages processed by this script. Time-stamp indicates last time-stamp of messages inspected
[string]$t_Matched   = $t_name + ".matched"  # Messages that matched regular expression for Create/Delete/Edit
[string]$t_Changed   = $t_name + ".changed"  # Point edits that succeded
[string]$t_Failed    = $t_name + ".failed"   # Point edits that failed
#
[string]$t_Status    = $t_name + ".status"   # witten by AF Analytic that monitors running of this script
#
# There is no easy way to write to PI Message Log at the time that this script was being created.
#  So this script writes to standard output & the Windows Event log
#
Write-Eventlog -LogName $t_logname -Source $t_logsource -EntryType Information -EventId 3 -Message "Starting"
Write-Host ("{0} {1:dd-MMM-yyyy HH:mm:ss} {2}:{3} {4}" -f "I", (Get-Date), $t_name, "", "Starting")
# 
# The big loop continues until interrupted by Control-C from the command line or IDE, or by stopping the task
try
    {
    #
    # Get the PI Server object
    $PIServer = Get-PIDataArchiveConnectionConfiguration $PIServerName -ErrorAction Stop
    # connect to PI Server
    $con = Connect-PIDataArchive -PIDataArchiveConnectionConfiguration $PIServer -ErrorAction Stop
    #
    # get PointIDs of PI points for status
    $pid_State    = (Get-PIPoint -name $t_State    -connection $con).Point.ID
    # get PI snapshot values of status
    $v_State      = ( Get-PIValue -connection $con -PointID $pid_State -Count 1 -Reverse -StartTime (Get-Date).AddDays(1) )
    # An alternate version might do something different depending on the value of the state.
    #
    # set PI value of status to starting
    Add-PIValue -connection $con -PointID $pid_State -Time (Get-Date) -Value "Starting"
    # get PointIDs of PI points for counters
    $pid_Messages = (Get-PIPoint -name $t_Messages -connection $con).Point.ID
    $pid_Matched  = (Get-PIPoint -name $t_Matched  -connection $con).Point.ID
    $pid_Changed  = (Get-PIPoint -name $t_Changed  -connection $con).Point.ID
    $pid_Failed   = (Get-PIPoint -name $t_Failed   -connection $con).Point.ID
    #
    # Retrieve starting values of integral counters from the SnapShot
    # Account for the possibility that the value is a digital state (likely PTCreated or Shutdown) and treat as 0
    # There is no direct way to specify retrieving the SnapShot.
    # The following expressions return the first value prior to 1 day in the future which should be the SnapShot
    #
    # $pid_Matched
    $v = ( Get-PIValue -connection $con -PointID $pid_Matched  -Count 1 -Reverse -StartTime (Get-Date).AddDays(1) )
    [int64]$v_Matched  = if ($v.Value.GetType().Name -eq "EventState") { 0 } else {$v.Value}
    # $pid_Changed
    $v = ( Get-PIValue -connection $con -PointID $pid_Changed  -Count 1 -Reverse -StartTime (Get-Date).AddDays(1) )
    [int64]$v_Changed  = if ($v.Value.GetType().Name -eq "EventState") { 0 } else {$v.Value}
    # $pid_Failed
    $v = ( Get-PIValue -connection $con -PointID $pid_Failed   -Count 1 -Reverse -StartTime (Get-Date).AddDays(1) )
    [int64]$v_Failed   = if ($v.Value.GetType().Name -eq "EventState") { 0 } else {$v.Value}
    # $pid_Messages
    $v = ( Get-PIValue -connection $con -PointID $pid_Messages -Count 1 -Reverse -StartTime (Get-Date).AddDays(1) )
    [int64]$v_Messages = if ($v.Value.GetType().Name -eq "EventState") { 0 } else {$v.Value}
    #
    # The SnapShot for Messages counter represents the last end time that this script collected messages
    [DateTime]$t_Messages = $v.TimeStamp
    # set the end time to the last time messages were collected so that this script can pick up from the last time it ran.
    [DateTime]$et = $t_messages 
    #
    Write-Host ("{0} {1:dd-MMM-yyyy HH:mm:ss} {2}:{3} {4}" -f "I", (Get-Date), $t_name, "", "Entering Processing Loop. Last ran $t_messages, starting from $et")
    Write-Eventlog -LogName $t_logname -Source $t_logsource -EntryType Information -EventId 4 -Message ("Entering Processing Loop. Last ran $t_messages, starting from $et")
    while ($true)
        {
        # set PI value of status to processing
        Add-PIValue -connection $con -PointID $pid_State -Time (Get-Date) -Value "Processing"
        #
	    # Update the start time to be 1 ms after the previous end time
	    $st = $et.AddMilliseconds(1)
        #
        # During startup after a long time, the following will limit retrieval by $MessageChunk s at a time.
        $et1 = $et.AddSeconds($MessageChunk)
        # While running continously, the end time should be 1 second before current time with no sub-second time
        #  This gives the PIMsgSS time to receive all of the messages
	    $et2 = (Get-Date -millisecond 0).addSeconds(-1)
        # get the earlier of $et1 and $et2
        if ( $et1.CompareTo($et2) -le 0 )
            {
            # startup mode. Chunked retrieval of messages applies. No wait in the while loop
            $et = $et1 
            [int]$Wait_this = 0
            }
        else
            { 
            # normal running. Retrieve messages since last time through. Wait after each pass through the while loop.
            $et = $et2 
            [int]$Wait_this = $Wait
            }
        #
        # Assume less than 2*31 messages for each pass through the loop, so Int32 counter will work for the number of messages processed each pass
        [int]$c_Messages = 0
        [int]$c_Matched  = 0
        [int]$c_Changed  = 0
        [int]$c_Failed   = 0
        #
	    # Retrieve messages from the server.
        #
        # The applicable Message ID for this script
        # 6079 - Create/Edit Point, Trusts, ... PIBaseSS
        #
        # Some other Message ID's from PI:
        # 7039 - Begin connection       PINetMgr
        # 7080 - Connection information
        # 7096 - End connection
        # 7121 - End connection
        # 7133 - Connection Statistics
        #
        # Get-PIMessage filter:  -Program PIBaseSS -ID 6079
        #
        # If there are no messages, Get-PIMessage will return an empty set of objects and an error. Silently ignore the error.
        # If there are very many messages, this could run out of memory in a 32-bit environment.
        #
	    $messages = Get-PIMessage -Connection $con -StartTime $st -EndTime $et -ID 6079 -Program PIBaseSS -ErrorAction SilentlyContinue
        # On restart, or with a long time interval, that might have taken some time, so update the status
        Add-PIValue -connection $con -PointID $pid_State -Time (Get-Date) -Value "Processing"
        #
        ForEach($m in $messages)
            {
            # WI #166089 GetPIMessage -ID filter returns messages that do not match the filter.
            # However, at least the filter reduces the volume of messages significantly
            $c_Messages++
            #
            # The following RegEx filter will only match the messages corresponding to Point Create/Delete/Edit
            # Note that the RegEx has escaped meta-characters "[]()" which appear in the PI Messages.
            # The RegEx is anchored to the beginning of the message to speed up rejection of non-matching messages
            if ($m.Message -match "^Point \[Name: (.*), ID: (.*)\] - (....?ted) by user (.*) \(userid: (.*), cnxnid: (.*)\)" -eq $true)
                {
                # The message matched a Point Created/Deleted/Edited
                $c_Matched++
                # In this block:
			    #  $Matches[1] Point TagName
			    #  $Matches[2] PointID
			    #  $Matches[3] Action (Created, Deleted, Edited)
			    #  $Matches[4] OSUser
			    #  $Matches[5] UserID
			    #  $Matches[6] ConnectionID
                $pm_TagName = $Matches[1]
                $pm_PointID = $Matches[2]
                $pm_Action  = $Matches[3]
                Write-Host ("{0} {1:dd-MMM-yyyy HH:mm:ss} {2}: {3} {4} {5}" -f "D", ($m.TimeStamp), $t_name, $pm_TagName, $pm_PointID, $pm_Action )
                # For testing, match Created or Edited. For production match Created
                if (($pm_Action -eq "Created") -or ( $MatchEdited -and ($pm_Action -eq "Edited")) )
                    {
                    # Get-PIPoint can only retrieve a Point by TagName not PointID.
                    # Renaming a PI Point makes TagName mutable relative to PointID, so will have to check PointID later...
                    $pt = Get-PIPoint -Connection $con -Name $pm_TagName -Attributes tag,pointid,pointsource,pointclass,pointtype
                    # if $pt is null, a Point with the TagName does not exist. Either it was renamed or deleted.
                    if ($pt -ne $null)
                        {
                        $pt_PointID     = $pt.Point.ID
                        If ($OSI_PS_Ver -GT [int]3000000)
                            {
                            # In OSIsoft.PowerShell version 3.n and later   Get-PIPoint returns a HashTable. WI 167404
                            $pt_PointSource = $pt.Attributes["PointSource"]
                            $pt_PointClass  = $pt.Attributes["PtClassName"]
                            $pt_PointType   = $pt.Attributes["PointType"]
                            }
                        else
                            {
                            # In OSIsoft.PowerShell version 2.1.0.5 and earlier Get-PIPoint returns a Dictionary instead of a HashTable.     WI 167404
                            # Use a workaround to get the PointSource, PointClass, and PointType from the Dictionary
                            $pt_PointSource = $pt.Attributes.GetEnumerator() | ForEach-Object { if ($_.Key.NativeAttributeName -eq "PointSource") {$_.Value } }
                            $pt_PointClass  = $pt.Attributes.GetEnumerator() | ForEach-Object { if ($_.Key.NativeAttributeName -eq "PtClassName") {$_.Value } }
                            $pt_PointType   = $pt.Attributes.GetEnumerator() | ForEach-Object { if ($_.Key.NativeAttributeName -eq "PointType"  ) {$_.Value } }
                            }
                        #
                        Write-Host ("{0} {1:dd-MMM-yyyy HH:mm:ss} {2}: {3} PointSource = {4}, PointClass = {5}, PointType = {6}" -f "D", ($m.TimeStamp), $t_name, $pn_TagName, $pt_PointSource, $pt_PointClass, $pt_PointType )
                        #
                        If ( ($pt_PointID -eq $pm_PointID) -and ( $h_PointSource.Contains($pt_PointSource) ) -and ( $pt_PointClass -eq "base" ) )
                            {
                            # This Point matches the TagName and PointID in the message.
                            # It has one of the correct PointSource(s) and PointClase = Base
                            #
                            # Build a hashtable with the Point attributes to change.
                            # Change PointType from Float64 to Float32 otherwise leave alone
                            # MDA WI #112059 PointType is not recognized when passed to Set-PIPoint, so exclude PointType in earlier versions 
                            $h_PointEdit = If (($pt_PointType -eq "Float64") -and ($OSI_PS_Ver -gt [int]3000000) ) { @{PointClass="Classic";PointType="Float32"} } else { @{PointClass="Classic"} }
                            # $h_PointEdit
                            #
                            # Point Edit might fail most likely because PI Point permissions do not allow it
                            try
                                {
                                Set-PIPoint -connection $con -Name ($pt.Point.Name) -Attributes $h_PointEdit
                                # if the script gets here, the Edit completed
                                $c_Changed++
                                Write-Host ("{0} {1:dd-MMM-yyyy HH:mm:ss} {2}: {3} {4} {5}" -f "D", (Get-Date), $t_name, $pm_TagName, $pt_PointID, " was edited")
                                Add-PIValue -connection $con -PointID $pid_State -Time (Get-Date) -Value "Processing"
                                Write-Eventlog -LogName $t_logname -Source $t_logsource -EntryType Information -EventId 11 -Message ("TagName: " + $pm_TagName + ", PointID: " + $pt_PointID + " was edited: " + ( $h_PointEdit.getenumerator() | ForEach-Object { $_.Key + " = " + $_.Value } ) )
                                }
                            catch
                                {
                                # if the script gets here, the Edit failed
                                $c_Failed++
                                $ErrorMessage = $_.Exception.Message
                                $FailedItem = $_.Exception.ItemName
                                Write-Host ("{0} {1:dd-MMM-yyyy HH:mm:ss} {2}:{3} {4}" -f "E", (Get-Date), $t_name, $ErrorMessage, $FailedItem)
                                Add-PIValue -connection $con -PointID $pid_State -Time (Get-Date) -Value "Error-Edit"
                                Write-Eventlog -LogName $t_logname -Source $t_logsource -EntryType Error -EventId 12 -Message ("TagName: " + $pm_TagName + ", PointID: " + $pt_PointID + "; " + $ErrorMessage + "; " + $FailedItem)
                                }
                            }
                        else
                            {
                            # write a value to the state point to protect against timeout when processing large batches of messages that do not match PointID, PointSource and PointClass Base
                            Add-PIValue -connection $con -PointID $pid_State -Time (Get-Date) -Value "Processing"
                            }
                        }
                    else
                        {
                        Write-Host ("{0} {1:dd-MMM-yyyy HH:mm:ss} {2}: {3} {4} {5}" -f "D", (Get-Date), $t_name, $pm_TagName, $pt_PointID, " was already Deleted or Renamed" )
                        Add-PIValue -connection $con -PointID $pid_State -Time (Get-Date) -Value "Processing"
                        Write-Eventlog -LogName $t_logname -Source $t_logsource -EntryType Information -EventId 10 -Message ("TagName: " + $pm_TagName + ", PointID: " + $pt_PointID + " was already deleted")
                        }
                    }
                }
            if ($true)
                {
                # Output any messages returned by PIGet-Msg filter
		        # To format this the same as pigetmsg, we first just want the first character of the severity
		        $sev = $m.Severity.ToString().SubString(0,1)
		        #
		        # Set the color of the message based on its severity
		        switch ($m.Severity)
		            {
			        Critical { $newForegroundColor = "Red" }
			        Error { $newForegroundColor = "Red" }
			        Warning { $newForegroundColor = "White" }
			        Informational { $newForegroundColor = "Gray" }
			        Debug { $newForegroundColor = "DarkGray" }
			        default { $newForegroundColor = -1 }
		            }
		        # If Source1 is empty, we want to put a colon after the Program name to display Source1
		        if ([string]::IsNullOrEmpty($m.Source1) -eq $true)
		            {
			        # To right justify the ID, get the size of the screen, and subtract the length of
			        # the items displayed
			        $width = $Host.UI.RawUI.WindowSize.Width - (27 + $m.ProgramName.Length + $m.ID.ToString().Length)
                    if ($newForegroundColor -eq -1)
                        {
    			        Write-Host ("{0} {1:dd-MMM-yyyy HH:mm:ss} {2} {4,$width}({3})" -f $sev, $m.TimeStamp, $m.ProgramName, $m.ID, "")
                        }
                    else
                        {
    			        Write-Host ("{0} {1:dd-MMM-yyyy HH:mm:ss} {2} {4,$width}({3})" -f $sev, $m.TimeStamp, $m.ProgramName, $m.ID, "") -ForegroundColor $newForegroundColor
                        }
		            }
		        else
		            {
			        $width = $Host.UI.RawUI.WindowSize.Width - (28 + $m.ProgramName.Length + $m.Source1.Length + $m.ID.ToString().Length)
                    if ($newForegroundColor -eq -1)
                        {
			            Write-Host ("{0} {1:dd-MMM-yyyy HH:mm:ss} {2}:{3} {5,$width}({4})" -f $sev, $m.TimeStamp, $m.ProgramName, $m.Source1, $m.ID, "")
                        }
                    else
                        {
			            Write-Host ("{0} {1:dd-MMM-yyyy HH:mm:ss} {2}:{3} {5,$width}({4})" -f $sev, $m.TimeStamp, $m.ProgramName, $_.Source1, $m.ID, "") -ForegroundColor $newForegroundColor
                        }
		            }

                if ($newForegroundColor -eq -1)
                    {
    		        Write-Host (" >> {0}" -f $m.Message)
                    }
                else
                    {
    		        Write-Host (" >> {0}" -f $m.Message) -ForegroundColor $newForegroundColor
                    }
	            Write-Host
                }
            }
        # add the results of the current ForEach loop to the totals
        $v_Messages += $c_Messages
        $v_Matched  += $c_Matched
        $v_Changed  += $c_Changed
        $v_Failed   += $c_Failed
        # Write-Host ("Messages {0} {1}; Matched {2} {3}; Changed {4} {5}; Failed {6} {7}" -f $v_Messages, $c_Messages, $v_Matched, $c_Matched, $v_Changed, $c_Changed, $v_Failed, $c_Failed )
        #
        # Update the counters stored in Int32 PI Points which can store up to (2**31-1). So truncate the Int64 to 31 bits
        # Need to cast the value to int, else Add-PIValue silently rejects the value
        # use timestamp of end time retrieved
        Add-PIValue -connection $con -PointID $pid_Messages -Time $et -Value ( [int]($v_Messages-bAND 0x07fffffff) )
        Add-PIValue -connection $con -PointID $pid_Matched  -Time $et -Value ( [int]($v_Matched -bAND 0x07fffffff) )
        Add-PIValue -connection $con -PointID $pid_Changed  -Time $et -Value ( [int]($v_Changed -bAND 0x07fffffff) )
        Add-PIValue -connection $con -PointID $pid_Failed   -Time $et -Value ( [int]($v_Failed  -bAND 0x07fffffff) )
        # update status for good measure
        Add-PIValue -connection $con -PointID $pid_State    -Time (Get-Date) -Value "Processing"
        #
	    # Wait for the specified time in seconds.
        # Note that if processing takes a long time, the start time will drift 
        Start-Sleep -seconds $Wait_this
        #
        }
    }
catch
    {
    # catch and report errors that might occur
    $ErrorMessage = $_.Exception.Message
    $FailedItem   = $_.Exception.ItemName
    Write-Host ("{0} {1:dd-MMM-yyyy HH:mm:ss} {2}:{3} {4}" -f "E", (Get-Date), $t_name, $ErrorMessage, $FailedItem)
    Add-PIValue -connection $con -PointID $pid_State -Time (Get-Date) -Value "Error"
    Write-Eventlog -LogName $t_logname -Source $t_logsource -EntryType Error -EventId 5 -Message ($ErrorMessage + "; " + $FailedItem)
    }
finally
    {
    # allow for a graceful exit after Stop, Control-C, or Stop Service.
    Write-Host ("{0} {1:dd-MMM-yyyy HH:mm:ss} {2}:{3} {4}" -f "I", (Get-Date), $t_name, "", "Stopping")
    Add-PIValue -connection $con -PointID $pid_State -Time (Get-Date) -Value "Stopping"
    Write-Eventlog -LogName $t_logname -Source $t_logsource -EntryType Information -EventId 2 -Message "Stopping"
    #
    If ($ExitNormally)
        {
        Write-Host ("{0} {1:dd-MMM-yyyy HH:mm:ss} {2}:{3} {4}" -f "I", (Get-Date), $t_name, "", "Stopped")
        Add-PIValue -connection $con -PointID $pid_State -Time (Get-Date) -Value "Stopped"
        Write-Eventlog -LogName $t_logname -Source $t_logsource -EntryType Information -EventId 1 -Message "Stopped"
        }
    }