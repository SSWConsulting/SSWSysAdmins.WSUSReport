﻿﻿<#
.SYNOPSIS
    PowerShell WSUS report builder.
.DESCRIPTION
    PowerShell WSUS report builder.
    It fetches data from on-premises WSUS server, builds an email and sends it to the target email.
.EXAMPLE
    This script runs on a schedule and accepts no parameters, only reads the parameteres from a configuration file in the same folder.
.INPUTS
    Configuration file: Config.psd1
.OUTPUTS
    Email to target email.
.NOTES
    Created by Steven Andrews and Kaique "Kiki" Biancatti for SSW.
#>

# Importing the configuration file
$config = Import-PowerShellDataFile $PSScriptRoot\Config.PSD1

# Importing the SSW Write-Log module
Import-Module -Name $config.LogModuleLocation

# Creating variables to determine magic strings and getting them from the configuration file
$UserKey = $config.UserKey
$User = $config.User
$Pass = $config.Pass
$LogFile = $config.LogFile
$OriginEmail = $config.OriginEmail
$TargetEmail = $config.TargetEmail
$WSUSServer = $config.WSUSServer
$ScriptServer = $config.ScriptServer
$GitHubLink = $config.GitHubLink

# Get the encrypted username and password files
$password = $Pass | ConvertTo-SecureString -Key (get-content $UserKey)
$cred = New-Object System.Management.Automation.PsCredential($User, $password)

Write-Log -File $LogFile -Message "Initiating connection to $WSUSServer to fetch WSUS data..."
Invoke-Command -ComputerName $WSUSServer -Credential $cred -ArgumentList $WSUSServer, $OriginEmail, $TargetEmail, $ScriptServer, $GitHubLink -ScriptBlock {

    #region User Specified WSUS Information
    $WSUSServer = $args[0]

    #Accepted values are "80","443","8530" and "8531"
    $Port = 8530
    $UseSSL = $False

    #Specify when a computer is considered stale
    $DaysComputerStale = 30 

    #Send email of report
    [bool]$SendEmail = $FALSE
    #Display HTML file
    [bool]$ShowFile = $TRUE
    #endregion User Specified WSUS Information

    #region Helper Functions
    Function Set-AlternatingCSSClass {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
            [string]$HTMLFragment,
            [Parameter(Mandatory = $True)]
            [string]$CSSEvenClass,
            [Parameter(Mandatory = $True)]
            [string]$CssOddClass
        )
        [xml]$xml = $HTMLFragment
        $table = $xml.SelectSingleNode('table')
        $classname = $CSSOddClass
        foreach ($tr in $table.tr) {
            if ($classname -eq $CSSEvenClass) {
                $classname = $CssOddClass
            }
            else {
                $classname = $CSSEvenClass
            }
            $class = $xml.CreateAttribute('class')
            $class.value = $classname
            $tr.attributes.append($class) | Out-null
        }
        $xml.innerxml | out-string
    }
    Function Convert-Size {
        [cmdletbinding()]
        Param (
            [parameter(ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
            [Alias("Length")]
            [int64]$Size
        )
        Begin {
            If (-Not $ConvertSize) {
                Write-Verbose ("Creating signature from Win32API")
                $Signature = @"
                 [DllImport("Shlwapi.dll", CharSet = CharSet.Auto)]
                 public static extern long StrFormatByteSize( long fileSize, System.Text.StringBuilder buffer, int bufferSize );
"@
                $Global:ConvertSize = Add-Type -Name SizeConverter -MemberDefinition $Signature -PassThru
            }
            Write-Verbose ("Building buffer for string")
            $stringBuilder = New-Object Text.StringBuilder 1024
        }
        Process {
            Write-Verbose ("Converting {0} to upper most size" -f $Size)
            $ConvertSize::StrFormatByteSize( $Size, $stringBuilder, $stringBuilder.Capacity ) | Out-Null
            $stringBuilder.ToString()
        }
    }
    #endregion Helper Functions

    #region Load WSUS Required Assembly
    If (-Not (Get-Module -ListAvailable -Name UpdateServices)) {
        #Add-Type "$Env:ProgramFiles\Update Services\Api\Microsoft.UpdateServices.Administration.dll"
        $Null = [reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
    }
    Else {
        Import-Module -Name UpdateServices
    }
    #endregion Load WSUS Required Assembly

    #region CSS Layout
    $head = @"
    <style> 
        h1 {
            text-align:center;
            border-bottom:1px solid #666666;
            color:#009933;
        }
		TABLE {
			TABLE-LAYOUT: fixed; 
			FONT-SIZE: 75%; 
			WIDTH: 100%
		}
		* {
			margin:0
		}

		.pageholder {
			margin: 0px auto;
		}
					
		td {
			VERTICAL-ALIGN: TOP; 
			FONT-FAMILY: Tahoma
		}
					
		th {
			VERTICAL-ALIGN: TOP; 
			COLOR: #018AC0; 
			TEXT-ALIGN: left;
            background-color:DarkGrey;
            color:Black;
		}
        body {
            text-align:left;
            font-smoothing:always;
            width:100%;
        }
        .odd { background-color:#ffffff; }
        .even { background-color:#dddddd; }               
    </style>
"@
    #endregion CSS Layout

    #region Initial WSUS Connection
    $ErrorActionPreference = 'Stop'
    Try {
        $Wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WSUSServer, $UseSSL, $Port)
    }
    Catch {
        Write-warning "$($WSUSServer)<$($Port)>: $($_)"
        Break
    }
    $ErrorActionPreference = 'Continue'
    #endregion Initial WSUS Connection

    #region Pre-Stage -- Used in more than one location
    $htmlFragment = ''
    $WSUSConfig = $Wsus.GetConfiguration()
    $WSUSStats = $Wsus.GetStatus()
    $TargetGroups = $Wsus.GetComputerTargetGroups()
    $EmptyTargetGroups = $TargetGroups | Where {
        $_.GetComputerTargets().Count -eq 0 -AND $_.Name -ne 'Unassigned Computers'

        #Stale Computers
        $computerscope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
        $computerscope.ToLastReportedStatusTime = (Get-Date).AddDays(-$DaysComputerStale)
        $StaleComputers = $wsus.GetComputerTargets($computerscope) | ForEach {
            [pscustomobject]@{
                Computername = $_.FullDomainName
                ID           = $_.Id
                IPAddress    = $_.IPAddress
                LastReported = $_.LastReportedStatusTime
                LastSync     = $_.LastSyncTime
                TargetGroups = ($_.GetComputerTargetGroups() | Select -Expand Name) -join ', '
            }
        }
    }
    #Pending Reboots
    $updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
    $updateScope.IncludedInstallationStates = 'InstalledPendingReboot'
    $computerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
    $computerScope.IncludedInstallationStates = 'InstalledPendingReboot'
    $GroupRebootHash = @{ }
    $ComputerPendingReboot = $wsus.GetComputerTargets($computerScope) | ForEach {
        $Update = ($_.GetUpdateInstallationInfoPerUpdate($updateScope) | ForEach {
                $Update = $_.GetUpdate()
                $Update.title
            }) -join ', '
        If ($Update) {
            $TempTargetGroups = ($_.GetComputerTargetGroups() | Select -Expand Name)
            $TempTargetGroups | ForEach {
                $GroupRebootHash[$_]++
            }
            [pscustomobject] @{
                Computername = $_.FullDomainName
                #            ID = $_.Id
                IPAddress    = $_.IPAddress
                TargetGroups = $TempTargetGroups -join ', '
                #Updates = $Update
            }
        }
    } | Sort Computername

    #Failed Installations
    $updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
    $updateScope.IncludedInstallationStates = 'Failed'
    $computerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
    $computerScope.IncludedInstallationStates = 'Failed'
    $GroupFailHash = @{ }
    $ComputerHash = @{ }
    $UpdateHash = @{ }
    $ComputerFailInstall = $wsus.GetComputerTargets($computerScope) | ForEach {
        $Computername = $_.FullDomainName
        $Update = ($_.GetUpdateInstallationInfoPerUpdate($updateScope) | ForEach {
                $Update = $_.GetUpdate()
                $Update.title
                $ComputerHash[$Computername] += , $Update.title
                $UpdateHash[$Update.title] += , $Computername
            }) -join ', '
        If ($Update) {
            $TempTargetGroups = ($_.GetComputerTargetGroups() | Select -Expand Name)
            $TempTargetGroups | ForEach {
                $GroupFailHash[$_]++
            }
            [pscustomobject] @{
                Computername = $_.FullDomainName
                #            ID = $_.Id
                IPAddress    = $_.IPAddress
                TargetGroups = $TempTargetGroups -join ', '
                Updates      = $Update
            }
        }
    } | Sort Computername
    #endregion Pre-Stage -- Used in more than one location

    #startregion Automatic and Manual Update

    $StagingCount = $wsus.GetComputerTargetGroups() | Where { $_.Name -eq 'Staging' } | Select-Object -ExpandProperty Id
    $StagingCount = $wsus.GetComputerTargetGroup($StagingCount)
    $StagingTargets = ($StagingCount).GetComputerTargets()
    $StagingTargets.Count

    $DevelopmentCount = $wsus.GetComputerTargetGroups() | Where { $_.Name -eq 'Development' } | Select-Object -ExpandProperty Id
    $DevelopmentCount = $wsus.GetComputerTargetGroup($DevelopmentCount)
    $DevelopmentTargets = ($DevelopmentCount).GetComputerTargets()
    $DevelopmentTargets.Count

    $DeveloperVMCount = $wsus.GetComputerTargetGroups() | Where { $_.Name -eq 'Developer VMs' } | Select-Object -ExpandProperty Id
    $DeveloperVMCount = $wsus.GetComputerTargetGroup($DeveloperVMCount)
    $DeveloperVMTargets = ($DeveloperVMCount).GetComputerTargets()

    $DeveloperVMTargets.Count

    $AutoUpdateCount = $StagingTargets.Count + $DevelopmentTargets.Count + $DeveloperVMTargets.Count

    $ManualUpdateCount = $wsus.GetComputerTargets().Count - $AutoUpdateCount

    #endregion Automatic and Manual Update

    #region CLIENT INFORMATION
    $Pre = @"
<div style='margin: 0px auto; BACKGROUND-COLOR:#cb463c;Color:White;font-weight:bold;FONT-SIZE: 16pt;'>
    WSUS Client Information
</div>
"@
    #region Computer Statistics
    $WSUSComputerStats = [pscustomobject]@{
        TotalComputers                   = [int]$WSUSStats.ComputerTargetCount    
        "Stale($DaysComputerStale Days)" = ($StaleComputers | Measure-Object).count
        NeedingUpdates                   = [int]$WSUSStats.ComputerTargetsNeedingUpdatesCount
        FailedInstall                    = [int]$WSUSStats.ComputerTargetsWithUpdateErrorsCount
        PendingReboot                    = ($ComputerPendingReboot | Measure-Object).Count
        AutomaticUpdates                 = $AutoUpdateCount
        ManualUpdateCount                = $ManualUpdateCount 
    }

    $Pre += @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:	#414141;Color:White;font-weight:bold;FONT-SIZE: 14pt;'>
            Computer Statistics
        </div>

"@
    $Body = $WSUSComputerStats | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $htmlFragment += $Pre, $Body, $Post
    #endregion Computer Statistics

    #region Stale Computers
    $Pre = @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:	#414141;Color:White;font-weight:bold;FONT-SIZE: 14pt;'>
            Stale Computers ($DaysComputerStale Days)
        </div>

"@
    $Body = $StaleComputers | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $htmlFragment += $Pre, $Body, $Post
    #endregion Stale Computers

    #region Unassigned Computers
    $Unassigned = ($TargetGroups | Where {
            $_.Name -eq 'Unassigned Computers'
        }).GetComputerTargets() | ForEach {
        [pscustomobject]@{
            Computername    = $_.FullDomainName
            OperatingSystem = $_.OSDescription
            #            ID=  $_.Id
            IPAddress       = $_.IPAddress
            LastReported    = $_.LastReportedStatusTime
            LastSync        = $_.LastSyncTime
        }    
    }
    $Pre = @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:	#414141;Color:White;font-weight:bold;FONT-SIZE: 14pt;'>
            Unassigned Computers (in Unassigned Target Group)
        </div>

"@
    $Body = $Unassigned | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $htmlFragment += $Pre, $Body, $Post
    #endregion Unassigned Computers

    #region Failed Update Install
    $Pre = @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:	#414141;Color:White;font-weight:bold;FONT-SIZE: 14pt;'>
            Failed Update Installations By Computer
        </div>

"@
    $Body = $ComputerFailInstall | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $htmlFragment += $Pre, $Body, $Post
    #endregion Failed Update Install

    #region Pending Reboot 
    $Pre = @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:	#414141;Color:White;font-weight:bold;FONT-SIZE: 14pt;'>
            Computers with Pending Reboot
        </div>

"@
    $Body = $ComputerPendingReboot | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $htmlFragment += $Pre, $Body, $Post
    #endregion Pending Reboot

    #endregion CLIENT INFORMATION

    #region UPDATE INFORMATION
    $Pre = @"
<div style='margin: 0px auto; BACKGROUND-COLOR:#cb463c;Color:White;font-weight:bold;FONT-SIZE: 16pt;'>
    WSUS Update Information
</div>
"@
    #region Update Statistics
    $WSUSUpdateStats = [pscustomobject]@{
        TotalUpdates        = [int]$WSUSStats.UpdateCount    
        Needed              = [int]$WSUSStats.UpdatesNeededByComputersCount
        Approved            = [int]$WSUSStats.ApprovedUpdateCount
        Declined            = [int]$WSUSStats.DeclinedUpdateCount
        ClientInstallError  = [int]$WSUSStats.UpdatesWithClientErrorsCount
        UpdatesNeedingFiles = [int]$WSUSStats.ExpiredUpdateCount    
    }
    $Pre += @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:	#414141;Color:White;font-weight:bold;FONT-SIZE: 14pt;'>
            Update Statistics
        </div>

"@
    $Body = $WSUSUpdateStats | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $htmlFragment += $Pre, $Body, $Post
    #endregion Update Statistics

    #endregion UPDATE INFORMATION

    #region TARGET GROUP INFORMATION
    $Pre = @"
<div style='margin: 0px auto; BACKGROUND-COLOR:#cb463c;Color:White;font-weight:bold;FONT-SIZE: 16pt;'>
    WSUS Target Group Information
</div>
"@
    #region Target Group Statistics
    $GroupStats = [pscustomobject]@{
        TotalGroups      = [int]$TargetGroups.count
        TotalEmptyGroups = [int]$EmptyTargetGroups.Count
    }

    $Pre += @"
        <div style='margin: 0px auto; BACKGROUND-COLOR:	#414141;Color:White;font-weight:bold;FONT-SIZE: 14pt;'>
            Target Group Statistics
        </div>

"@
    $Body = $GroupStats | ConvertTo-Html -Fragment | Out-String | Set-AlternatingCSSClass -CSSEvenClass 'even' -CssOddClass 'odd'
    $Post = "<br>"
    $htmlFragment += $Pre, $Body, $Post
    #endregion Target Group Statistics

    #region Compile HTML Report
    $HTMLParams = @{
        Head        = $Head
        Title       = "WSUS Report for $WSUSServer"
        PostContent = @"
    <p dir="auto">Dear SysAdmins, </p><br>
    <p dir="auto">As per Intranet | Employee Responsibilities | Windows Updates and related rules</p>
    <p dir="auto" style=" font-size: 11.0pt; margin: 0cm 0cm .0001pt 36.0pt;"><a href="https://rules.ssw.com.au/do-you-disable-automatic-windows-update-installations">https://rules.ssw.com.au/do-you-disable-automatic-windows-update-installations</a> [for PCs]<br>
    <p dir="auto" style=" font-size: 11.0pt; margin: 0cm 0cm .0001pt 36.0pt;"><a href="https://rules.ssw.com.au/do-you-turn-off-auto-update-on-your-servers">https://rules.ssw.com.au/do-you-turn-off-auto-update-on-your-servers</a> [for Servers]<br> 
    <p dir="auto" style=" font-size: 11.0pt; margin: 0cm 0cm .0001pt 36.0pt;"><a href="https://rules.ssw.com.au/do-you-use-group-policy-to-manage-your-windows-update-policy">https://rules.ssw.com.au/do-you-use-group-policy-to-manage-your-windows-update-policy</a> [for both]</p>
    <br><ol dir="auto" style="margin-bottom: 0cm; margin-top: 0cm;">
        <li>Please reboot the PCs below &ndash; see &lsquo;Computers with Pending Reboot&rsquo;</li>
        <li>Then after 15 mins check Whats Up Gold is all green</li>
        <li>Then log into each server and see if any more updates to be done</li>
        <li>Look in Windows Event Viewer | System | [for any Critical events today] = 0</li>
        <li>Re-run this report again to see &lsquo;Computers with Pending Reboot&rsquo; = 0</li>
    </ol><br>
"@ + "$($htmlFragment)" + "--Powered by SSWSysAdmins.WSUSReport <br> WSUS Server: $($args[0]) <br> Script Origin Server: $($args[3]) <br> GitHub Documentation: <a href=$($args[4])>$($args[4])</a>"  
    }
    $Report = ConvertTo-Html @HTMLParams | Out-String
    #endregion Compile HTML Report

    If ($ShowFile) {
        $Report | Out-File WSUSReport.html
        Invoke-Item WSUSReport.html
    }

    #region Send Email
    If ($SendEmail) {
        $EmailParams.Body = $Report 
        Send-MailMessage @EmailParams
    }
    #endregion Send Email

    Send-MailMessage -from $args[1] -to $args[2] -Subject "SSW.SysAdmin - Windows Updates Report" -Body $Report -SmtpServer "ssw-com-au.mail.protection.outlook.com" -bodyashtml
}
Write-Log -File $LogFile -Message "Closing connection to $WSUSServer, email sent to $TargetEmail..."
