Function Get-PWBLInfo{
    [CmdletBinding()] 
    param (
        [Parameter(ValueFromPipelineByPropertyName=$false,Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="The drive we want to get BitLocker info from")] 
        [string]$Drive = 'C:'
    )
    try{$BitLocker = Get-WmiObject -Namespace "Root\cimv2\Security\MicrosoftVolumeEncryption" -Class "Win32_EncryptableVolume" -Filter "DriveLetter = `'$Drive`'" -ErrorAction stop}
    catch{$BitLocker = $false}
    $BitLocker
}

Function Get-PWBLStatus{
    [CmdletBinding()] 
    param (
        [Parameter(ValueFromPipelineByPropertyName=$false,Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="The WMI Object for a Win32_EncryptableVolume class")] 
        $BitLocker = (Get-PWBLInfo)
    )
    If ($BitLocker){
        #https://msdn.microsoft.com/en-us/library/windows/desktop/aa376448(v=vs.85).aspx
        switch ($BitLocker.GetProtectionStatus().protectionStatus){
            ("0"){$return = "Unprotected"}
            ("1"){$return = "Protected"}
            ("2"){$return = "Unknowned"}
            default {$return = "NoReturn"}
        }
    }else{
        $return = "NoReturn"
    }
    return $return
}


Function Get-PWBLProtectors{
    [CmdletBinding()] 
    param (
        [Parameter(ValueFromPipelineByPropertyName=$false,Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="The drive we want to get BitLocker info from")] 
        [string]$Drive = 'C:'
    )
    try{$BitLocker = Get-WmiObject -Namespace "Root\cimv2\Security\MicrosoftVolumeEncryption" -Class "Win32_EncryptableVolume" -Filter "DriveLetter = `'$Drive`'" -ErrorAction stop}
    catch{return "Error"}
    $ProtectorIds = $BitLocker.GetKeyProtectors().volumekeyprotectorID
    $return = @()
    foreach ($ProtectorID in $ProtectorIds){
	    $KeyProtectorType = $BitLocker.GetKeyProtectorType($ProtectorID).KeyProtectorType
	    $keyType = ""
		    #https://msdn.microsoft.com/en-us/library/windows/desktop/aa376442(v=vs.85).aspx
		    switch($KeyProtectorType){"0"{$Keytype = "Unknown or other protector type";break}
		    "1"{$Keytype = "Trusted Platform Module (TPM)";break}
		    "2"{$Keytype = "External key";break}
		    "3"{$Keytype = "Numerical password";break}
		    "4"{$Keytype = "TPM And PIN";break}
		    "5"{$Keytype = "TPM And Startup Key";break}
		    "6"{$Keytype = "TPM And PIN And Startup Key";break}
		    "7"{$Keytype = "Public Key";break}
		    "8"{$Keytype = "Passphrase";break}
		    "9"{$Keytype = "TPM Certificate";break}
		    "10"{$Keytype = "CryptoAPI Next Generation (CNG) Protector";break}
	    }
	    $return += new-object -typename PSObject -Property @{'Key' = "$ProtectorID";'KeyTypeID'=$KeyProtectorType;'KeyTypeName'="$KeyType";}
    }
    $return
}

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



$BitLockerLog = "C:\Windows\Logs\Software\BitlockerActions.log"


Write-Log -Message 'Starting Bitlocker key protector review...' -LogFile $BitLockerLog
Write-Log -Message 'Updating group policy for the machine.' -LogFile $BitLockerLog
$GPRes=(echo N | gpupdate.exe /Target:Computer /Force) -join ''
Write-Log -Message "Policy evaluation result: $GPRes" -LogFile $BitLockerLog
if ((Get-PWBLStatus) -eq "Protected"){
    Write-Log -Message 'System volume found as protected.  Reviewing key protectors.' -LogFile $BitLockerLog
    $status = 0
    Foreach ($KP in Get-PWBLProtectors){
        Write-Log -Message "Key Protector found on Volume: $($KP.KeyTypeName)]" -LogFile $BitLockerLog
        if ($KP.KeyTypeID -eq 1){
            $status = 1
        }
    }
    if ($status -eq 1){
        Write-Log -Message 'TPM only key protector found on volume.' -LogFile $BitLockerLog
        "Compliant"
    }else{
        Write-Log -Message 'TPM only key protector NOT found on volume.' -LogFile $BitLockerLog
        $ReturnValue = ((Get-PWBLInfo).ProtectKeyWithTPM()).ReturnValue
        If ($ReturnValue -eq 0){
            Write-Log -Message 'TPM only key protector successfully added to volume.' -LogFile $BitLockerLog
            "Compliant"
        }else{
            Write-Log -ErrorMessage "TPM only key protector not added to volume. Error [$ReturnValue]" -Type 3 -LogFile $BitLockerLog
        }
    }
}else{
    Write-Log -Message 'Drive not protected or access to WMI not granted' -Type 2 -LogFile $BitLockerLog
}

Write-Log -Message 'Key protector review complete.' -LogFile $BitLockerLog
exit $ReturnValue
