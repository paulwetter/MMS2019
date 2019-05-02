#Log location for actions in this script.
$BitLockerLog = "C:\Windows\Logs\Software\TPMActivation.log"

#This script will only configure TPM on Dell computers.

#This script will enable and activate the TPM if required for systems in your environment.

#Dell Command Configure 4 files should be located in a Subfolder called DellCommandConfigure4_1_0
#Dell Command Configure 3 files should be located in a Subfolder called LegacyCCTK


function Write-Log {
    [CmdletBinding()] 
    Param (
		[Parameter(Mandatory=$false)]
		$Message,
 
		[Parameter(Mandatory=$false)]
		$ErrorMessage,
 
		[Parameter(Mandatory=$false)]
		$Component,
 
		[Parameter(Mandatory=$false,HelpMessage="1 = Normal, 2 = Warning (yellow), 3 = Error (red)")]
        [ValidateSet(1, 2, 3)]
		[int]$Type,
		
		[Parameter(Mandatory=$false,HelpMessage="Size in KB")]
		[int]$LogSizeKB=512,

		[Parameter(Mandatory=$true)]
		$LogFile
	)
    <#
    Type: 1 = Normal, 2 = Warning (yellow), 3 = Error (red)
    #>
    $LogLength = $LogSizeKB*1024
    try{
        $log = Get-Item $LogFile -ErrorAction Stop
        If (($log.length) -gt $LogLength){
	        $Time = Get-Date -Format "HH:mm:ss.ffffff"
	        $Date = Get-Date -Format "MM-dd-yyyy"
            $LogMessage = "<![LOG[Closing log and generating new log file" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"1`" thread=`"`" file=`"`">"
            $LogMessage | Out-File -Append -Encoding UTF8 -FilePath $LogFile
            Move-Item -Path "$LogFile" -Destination "$($LogFile.TrimEnd('g'))_" -Force
        }
    }
    catch{Write-Verbose "Nothing to move or move failed."}

	$Time = Get-Date -Format "HH:mm:ss.ffffff"
	$Date = Get-Date -Format "MM-dd-yyyy"
 
	if ($ErrorMessage -ne $null) {$Type = 3}
	if ($Component -eq $null) {$Component = " "}
	if ($Type -eq $null) {$Type = 1}
 
	$LogMessage = "<![LOG[$Message $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"$Type`" thread=`"`" file=`"`">"
	$LogMessage | Out-File -Append -Encoding UTF8 -FilePath $LogFile
}

Write-Log -Message "Starting BIOS TPM Configuration..." -Type 1 -LogFile $BitLockerLog
$PCManufacturer = (Get-CimInstance -ClassName win32_ComputerSystem -ErrorAction SilentlyContinue).Manufacturer
Write-Log -Message "Computer Manufacturer: [$PCManufacturer]" -Type 1 -LogFile $BitLockerLog
$PCModel = (Get-CimInstance -ClassName win32_ComputerSystem -ErrorAction SilentlyContinue).Model
Write-Log -Message "Computer Model: [$PCModel]" -Type 1 -LogFile $BitLockerLog
Write-Log -Message "Attempting Dell Command Configure..." -Type 1 -LogFile $BitLockerLog
If (Test-Path "$PSScriptRoot\DellCommandConfigure4_1_0\cctk.exe"){
    Write-Log -Message "Dell Command Configure Executable found in [$($PSScriptRoot)\DellCommandConfigure4_1_0\cctk.exe" -Type 1 -LogFile $BitLockerLog
}else {
    Write-Log -Message "Dell Command Configure Executable Not found in [$($PSScriptRoot)\DellCommandConfigure4_1_0\cctk.exe" -Type 1 -LogFile $BitLockerLog
}
$TestWMIACPI = (&"$PSScriptRoot\DellCommandConfigure4_1_0\cctk.exe" --setuppwd=P@ssw0rd) -join ''
Write-Log -Message "$TestWMIACPI" -Type 1 -LogFile $BitLockerLog -Component 'DellCommandConfigure4_1_0'
# If we are unable to set a password on the BIOS with the "WMI-ACPI Buffer Size" error, then, 
#  the BIOS is likely older and we need to run the legacy 3.x CCTK against it.
If (($TestWMIACPI -like '*This system does not have a WMI-ACPI compliant BIOS*') -or ($TestWMIACPI -like '*WMI-ACPI Buffer Size*')){
    Write-Log -Message "Dell Command Configure 4.1.0 failed with: [$TestWMIACPI]" -Type 2 -LogFile $BitLockerLog
    Write-Log -Message "Attempting legacy CCTK..." -Type 1 -LogFile $BitLockerLog -Component 'LegacyCCTK'
    $TestLegacy = (&"$PSScriptRoot\LegacyCCTK\cctk.exe" --setuppwd=P@ssw0rd) -join ''
    If ($TestLegacy -like "*Password is set successfully*"){
        Write-Log -Message "Legacy CCTK set password successfully with: [$TestLegacy]" -Type 1 -LogFile $BitLockerLog -Component 'LegacyCCTK'
        Write-Log -Message "Enabling TPM." -Type 1 -LogFile $BitLockerLog -Component 'LegacyCCTK'
        $EnableTPM = (&"$PSScriptRoot\LegacyCCTK\cctk.exe" --tpm=on --valsetuppwd=P@ssw0rd) -join ''
        Write-Log -Message "$EnableTPM" -Type 1 -LogFile $BitLockerLog -Component 'LegacyCCTK'
        Write-Log -Message "Activating TPM." -Type 1 -LogFile $BitLockerLog -Component 'LegacyCCTK'
        $ActivateTPM = (&"$PSScriptRoot\LegacyCCTK\cctk.exe" --tpmactivation=activate --valsetuppwd=P@ssw0rd) -join ''
        Write-Log -Message "$ActivateTPM" -Type 1 -LogFile $BitLockerLog -Component 'LegacyCCTK'
        Write-Log -Message "Removing setup Password." -Type 1 -LogFile $BitLockerLog -Component 'LegacyCCTK'
        $UnsetPassword = (&"$PSScriptRoot\LegacyCCTK\cctk.exe" --setuppwd= --valsetuppwd=P@ssw0rd) -join ''
        Write-Log -Message "$UnsetPassword" -Type 1 -LogFile $BitLockerLog -Component 'LegacyCCTK'
    }else{
        Write-Log -Message "Legacy CCTK failed to set password. Something went wrong: [$TestLegacy]" -Type 3 -LogFile $BitLockerLog
    }
#If the first attempt to set the BIOS password was successful then, we can go ahead an use DCC 4.x to configure it.
}elseif ($TestWMIACPI -like '*Password is set successfully*'){
    Write-Log -Message "Password set successfully with Dell Command Configure 4.1.0: [$TestWMIACPI]" -Type 1 -LogFile $BitLockerLog -Component 'DellCommandConfigure4_1_0'
    Write-Log -Message "Enabling TPM." -Type 1 -LogFile $BitLockerLog -Component 'DellCommandConfigure4_1_0'
    $EnableTPM = (&"$PSScriptRoot\DellCommandConfigure4_1_0\cctk.exe" --tpm=on --valsetuppwd=P@ssw0rd) -join ''
    Write-Log -Message "$EnableTPM" -Type 1 -LogFile $BitLockerLog -Component 'DellCommandConfigure4_1_0'
    Write-Log -Message "Activating TPM." -Type 1 -LogFile $BitLockerLog -Component 'DellCommandConfigure4_1_0'
    $ActivateTPM = (&"$PSScriptRoot\DellCommandConfigure4_1_0\cctk.exe" --tpmactivation=activate --valsetuppwd=P@ssw0rd) -join ''
    Write-Log -Message "$ActivateTPM" -Type 1 -LogFile $BitLockerLog -Component 'DellCommandConfigure4_1_0'
    Write-Log -Message "Removing setup Password." -Type 1 -LogFile $BitLockerLog -Component 'DellCommandConfigure4_1_0'
    $UnsetPassword = (&"$PSScriptRoot\DellCommandConfigure4_1_0\cctk.exe" --setuppwd= --valsetuppwd=P@ssw0rd) -join ''
    Write-Log -Message "$UnsetPassword" -Type 1 -LogFile $BitLockerLog -Component 'DellCommandConfigure4_1_0'
}else{
    Write-Log -Message "Something went wrong: [$TestWMIACPI][$TestLegacy]" -Type 3 -LogFile $BitLockerLog
    #If $TestWMIACPI = Couldn't get WMI-ACPI Buffer Size!!  Then, it may not be a Dell computer.
    Write-Log -Message "May not be a Dell computer.  Common error when not a Dell: [Couldn't get WMI-ACPI Buffer Size]" -Type 3 -LogFile $BitLockerLog
}
Write-Log -Message "Ending BIOS TPM Configuration" -Type 1 -LogFile $BitLockerLog
