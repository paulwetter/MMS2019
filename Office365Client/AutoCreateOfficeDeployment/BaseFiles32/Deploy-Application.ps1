<#
.SYNOPSIS
	This script performs the installation or uninstallation of an application(s).
.DESCRIPTION
	The script is provided as a template to perform an install or uninstall of an application(s).
	The script either performs an "Install" deployment type or an "Uninstall" deployment type.
	The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
	The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
.PARAMETER DeploymentType
	The type of deployment to perform. Default is: Install.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
	Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER TerminalServerMode
	Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Destkop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
	Disables logging to file for the script. Default is: $false.
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"
.EXAMPLE
    Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"
.NOTES
	Toolkit Exit Code Ranges:
	60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
	69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
	70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK 
	http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false)]
	[ValidateSet('Install','Uninstall')]
	[string]$DeploymentType = 'Install',
	[Parameter(Mandatory=$false)]
	[ValidateSet('Interactive','Silent','NonInteractive')]
	[string]$DeployMode = 'Interactive',
	[Parameter(Mandatory=$false)]
	[switch]$AllowRebootPassThru = $false,
	[Parameter(Mandatory=$false)]
	[switch]$TerminalServerMode = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableLogging = $false
)

Try {
	## Set the script execution policy for this process
	Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}
	
	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	[string]$appVendor = 'Microsoft'
	[string]$appName = 'Office 365 ProPlus 2016'
	[string]$appVersion = '16.0.6965.2066'
	[string]$appArch = 'x86'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.1.1'
	[string]$appScriptDate = '07/19/2017'
	[string]$appScriptAuthor = 'Paul Wetter'
	##*===============================================
	## Variables: Install Titles (Only set here to override defaults set by the toolkit)
	[string]$installName = ''
	[string]$installTitle = ''
	
	##* Do not modify section below
	#region DoNotModify
	
	## Variables: Exit Code
	[int32]$mainExitCode = 0
	
	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.6.8'
	[string]$deployAppScriptDate = '02/06/2016'
	[hashtable]$deployAppScriptParameters = $psBoundParameters
	
	## Variables: Environment
	If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
	[string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent
	
	## Dot source the required App Deploy Toolkit Functions
	Try {
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
		If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
	}
	Catch {
		If ($mainExitCode -eq 0){ [int32]$mainExitCode = 60008 }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
	}
	
	#endregion
	##* Do not modify section above
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================
		
	If ($deploymentType -ine 'Uninstall') {
		##*===============================================
		##* PRE-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Installation'
		
		## Show Welcome Message, close Internet Explorer if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt
		#Show-InstallationWelcome -CloseApps 'iexplore' -AllowDefer -DeferTimes 3 -CheckDiskSpace -PersistPrompt
        Show-InstallationWelcome -CloseApps 'iexplore=Internet Explorer,winword=Microsoft Office Word,excel=Microsoft Office Excel,onenote=Microsoft Office OneNote,onenotem=Microsoft Office OneNote (Send to Onenote),outlook=Microsoft Office Outlook,msaccess=Microsoft Office Access,powerpnt=Microsoft Office PowerPoint,visio=Microsoft Office Visio,mspub=Microsoft Office Publisher,groove=Microsoft Office Groove,lync=Skype for Business or Lync,communicator=Lync,winproj=Microsoft Office Project' -PersistPrompt -ForceCloseAppsCountdown 600 -BlockExecution
		
		## Show Progress Message (with the default message)
		Show-InstallationProgress -StatusMessage 'Discovering current installed Office Applications...'
		
		## <Perform Pre-Installation tasks here>
        
        $Visio = $false
        $Project = $false

        Write-Log -Message "Checking WMI for other office apps." -Source 'Get-WmiObject'
        Write-Log -Message "Running the following WMI Query: Get-WmiObject -Query `"Select * from Win32_Product where IdentifyingNumber like `'%0FF1CE}`'`"" -Source 'Get-WmiObject'
        # Note if Visio or Project are installed
		$VisioID = @('0051','0053','0054','0055','0057')
        $ProjectID = @('003A','003B','00B4','00B5')
        $C2RVisio = @('VisioProRetail','VisioProXVolume','VisioStdRetail','VisioStdXVolume')
        $C2RProject = @('ProjectProRetail','ProjectProXVolume','ProjectStdRetail','ProjectStdXVolume')

        $OfficeProducts=Get-WmiObject -Query "Select * from Win32_Product where IdentifyingNumber like `'%0FF1CE}`'"

        if ($OfficeProducts){
                if ($OfficeProducts|where {$_.IdentifyingNumber.split('-')[1] -in $VisioID}) {$Visio = $true}
                if ($OfficeProducts|where {$_.IdentifyingNumber.split('-')[1] -in $ProjectID}) {$Project = $true}
            }

        $C2RProducts=(Get-ItemProperty -Path HKLM:SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name ProductReleaseIds).ProductReleaseIds -split ","

        if ($C2RProducts){
                if (Compare-Object $C2RProducts $C2RVisio -IncludeEqual -ExcludeDifferent) {$Visio = $true}
                if (Compare-Object $C2RProducts $C2RProject -IncludeEqual -ExcludeDifferent) {$Project = $true}
            }

        Write-Log -Message "Version of Visio found on system: $Visio" -Source 'Get-WmiObject'
        Write-Log -Message "Version of Project found on system: $Project" -Source 'Get-WmiObject'

        Write-Log -Message "Checking registry for flag for uninstalled Visio or Project..." -Source 'Get-ItemProperty'
        If((Get-ItemProperty -Path "HKLM:\SOFTWARE\CompanySysINfo\" -Name "VisioRemovedinCleanup" -ErrorAction Ignore).VisioRemovedinCleanup -eq "Yes"){
            Write-Log -Message "Registry flag set for Visio reinstall: HKLM:\SOFTWARE\CompanySysINfo\VisioRemovedinCleanup:Yes" -Source 'Get-ItemProperty'
            $Visio = $true
        }
        If((Get-ItemProperty -Path "HKLM:\SOFTWARE\CompanySysINfo\" -Name "ProjectRemovedinCleanup" -ErrorAction Ignore).ProjectRemovedinCleanup -eq "Yes"){
            Write-Log -Message "Registry flag set for Project reinstall: HKLM:\SOFTWARE\CompanySysINfo\ProjectRemovedinCleanup:Yes" -Source 'Get-ItemProperty'
            $Project = $true
        }
        Write-Log -Message "Registry flag check for Visio or Project complete.  Results posted above." -Source 'Get-ItemProperty'
        Write-Log -Message "Visio previous install Status: $Visio"
        Write-Log -Message "Project previous install Status: $Project"

        if ($Visio -eq $true) {$newVisio = Show-InstallationPrompt -Title 'Reinstall Visio?' -Message 'An old version of Visio was found on this computer and will be removed during Office cleanup.  Would you like to install the new version of Visio?' -ButtonLeftText 'Yes' -ButtonRightText 'No' -Timeout 60 -MinimizeWindows $true -ExitOnTimeout $false}
        if ($Project -eq $true) {$newProject = Show-InstallationPrompt -Title 'Reinstall Project?' -Message 'An old version of Project was found on this computer and will be removed during Office cleanup.  Would you like to install the new version of Project?' -ButtonLeftText 'Yes' -ButtonRightText 'No' -Timeout 60 -MinimizeWindows $true -ExitOnTimeout $false}
        Write-Log -Message "New Visio Install Answer: [$($newVisio)]  --  New Project Install Answer: [$($newProject)]" -Source 'Deploy-Application'

		Show-InstallationProgress -StatusMessage 'Removing old versions of Microsoft Office...'
        ## Remove Lync 2010 Stand alone client.
        Execute-MSI -Action Uninstall -Path "{11849FBC-C416-4742-8279-17C3A2C85F72}" -Parameters "/QN /norestart"
        ## Remove Lync 2013 -- Covered by "CLIENTALL"
        ## Execute-Process -Path "$envWinDir\System32\cscript.exe" -Parameters  "//nologo $dirSupportFiles\OffScrub_O15msi.vbs LYNCENTRY /Bypass 1,3,4 /NoCancel /S /Log $configToolkitLogDir" -WindowStyle Hidden -IgnoreExitCodes '2'
        #Delete all sip profiles for lync 2010
        Get-Item C:\Users\*\AppData\Local\Microsoft\Communicator\sip_*|remove-item -force -recurse

        #Exit code 2 means that a reboot is required.
        Execute-Process -Path "$envWinDir\System32\cscript.exe" -Parameters  "//nologo $dirSupportFiles\OffScrub10.vbs CLIENTALL /Bypass 1,3,4 /NoCancel /S /Log $configToolkitLogDir" -WindowStyle Hidden -IgnoreExitCodes '2'
        Execute-Process -Path "$envWinDir\System32\cscript.exe" -Parameters  "//nologo $dirSupportFiles\OffScrub_O15msi.vbs CLIENTALL /Bypass 1,3,4 /NoCancel /S /Log $configToolkitLogDir" -WindowStyle Hidden -IgnoreExitCodes '2'
        Execute-Process -Path "$envWinDir\System32\cscript.exe" -Parameters  "//nologo $dirSupportFiles\OffScrub_O16msi.vbs CLIENTALL /Bypass 1,3,4 /NoCancel /S /Log $configToolkitLogDir" -WindowStyle Hidden -IgnoreExitCodes '2'
        ## Execute-Process -Path "$envWinDir\System32\cscript.exe" -Parameters  "//nologo $dirSupportFiles\OffScrub10.vbs ALL /S" -WindowStyle Hidden -IgnoreExitCodes '1,2'
		
		
		
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		## Handle Zero-Config MSI Installations
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Install'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat; If ($defaultMspFiles) { $defaultMspFiles | ForEach-Object { Execute-MSI -Action 'Patch' -Path $_ } }
		}
		
		## <Perform Installation tasks here>
        Show-InstallationProgress
        ##Installing Office 365 ProPlus

        Write-Log -Message "New Visio Install Answer: [$($newVisio)]  --  New Project Install Answer: [$($newProject)]" -Source 'Deploy-Application'

        $InstallProject = $false
        $InstallVisio = $false
        If ($newVisio -eq "Yes") {$InstallVisio = $true}
        If ((!$newVisio) -and ($Visio -eq $true)) {$InstallVisio = $true}
        If ($newProject -eq "Yes") {$InstallProject = $true}
        If ((!$newProject) -and ($Project -eq $true)) {$InstallProject = $true}

        Write-Log -Message "Calculated Visio Install Answer: [$($InstallVisio)]  --  Calculated Project Install Answer: [$($InstallProject)]" -Source 'Deploy-Application'

        If (($InstallVisio -eq $true) -and ($InstallProject -eq $false)) {
            Write-Log -Message "Install Office and Visio" -Source 'Deploy-Application'
            Execute-Process -Path "Setup.exe" -Parameters "/configure $dirFiles\Config+Visio.xml" -WindowStyle Hidden
            }
        If (($InstallProject -eq $true) -and ($InstallVisio -eq $false)) {
            Write-Log -Message "Install Office and Project" -Source 'Deploy-Application'
            Execute-Process -Path "Setup.exe" -Parameters "/configure $dirFiles\Config+Project.xml" -WindowStyle Hidden
            }
        If (($InstallProject -eq $false) -and ($InstallVisio -eq $false)) {
            Write-Log -Message "Install Office only" -Source 'Deploy-Application'
            Execute-Process -Path "Setup.exe" -Parameters "/configure $dirFiles\Config.xml" -WindowStyle Hidden
            }
        If (($InstallProject -eq $true) -and ($InstallVisio -eq $true)) {
            Write-Log -Message "Install Office,Visio,Project" -Source 'Deploy-Application'
            Execute-Process -Path "Setup.exe" -Parameters "/configure $dirFiles\Config+Visio+Project.xml" -WindowStyle Hidden
            }

		#Execute-Process -Path "Setup.exe" -Parameters "/configure $dirFiles\Config.xml" -WindowStyle Hidden
		
		
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>
		
		## Display a message at the end of the install
		## If (-not $useDefaultMsi) { Show-InstallationPrompt -Message 'You can customize text to appear at the end of an install or remove it completely for unattended installations.' -ButtonRightText 'OK' -Icon Information -NoWait }
	}
	ElseIf ($deploymentType -ieq 'Uninstall')
	{
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'
		
		## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
		##Show-InstallationWelcome -CloseApps 'iexplore' -CloseAppsCountdown 60
        Show-InstallationWelcome -CloseApps 'iexplore=Internet Explorer,winword=Microsoft Office Word,excel=Microsoft Office Excel,onenote=Microsoft Office OneNote,onenotem=Microsoft Office OneNote (Send to Onenote),outlook=Microsoft Office Outlook,msaccess=Microsoft Office Access,powerpnt=Microsoft Office PowerPoint,visio=Microsoft Office Visio,mspub=Microsoft Office Publisher,groove=Microsoft Office Groove,lync=Skype for Business or Lync,communicator=Lync,winproj=Microsoft Office Project' -PersistPrompt -ForceCloseAppsCountdown 600 -BlockExecution
		
		## Show Progress Message (with the default message)
		Show-InstallationProgress
		
		## <Perform Pre-Uninstallation tasks here>
		
		
		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'
		
		## Handle Zero-Config MSI Uninstallations
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat
		}
		
		# <Perform Uninstallation tasks here>
        Execute-Process -Path "Setup.exe" -Parameters "/configure Uninstall.xml" -WindowStyle Hidden
		
		##*===============================================
		##* POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'
		
		## <Perform Post-Uninstallation tasks here>
		
		
	}
	
	##*===============================================
	##* END SCRIPT BODY
	##*===============================================
	
	## Call the Exit-Script function to perform final cleanup operations
	Exit-Script -ExitCode $mainExitCode
}
Catch {
	[int32]$mainExitCode = 60001
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}