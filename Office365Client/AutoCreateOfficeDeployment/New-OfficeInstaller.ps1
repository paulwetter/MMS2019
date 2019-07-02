[CmdletBinding()]
param()

[string]$CMServer = 'localhost'
[string]$CMSite = 'MMS'

#The name of your application.  The version will be appended to the name in SCCM Console.  Software center will show without the name.
[string]$AppName = 'Office 365 Client'
#Application Publisher
[string]$Publisher = 'Microsoft'
#location of the icon you will use in software center
[string]$IconPath = 'G:\Staging\Office365\office365icon.ico'
#Install command string for deployment type
[string]$InstallCmd = 'Deploy-Application.exe -DeploymentType "Install" -AllowRebootPassThru'
#Uninstall command string for deployment type
[string]$UninstallCmd = 'Deploy-Application.exe -DeploymentType "Uninstall" -AllowRebootPassThru'
#Directory where you would like to store the built source files for your Office Install.
[string]$AppSourceFiles = '\\mms-cm1\Sources\Software\Microsoft\Office365Client'
#The Distribution Point Group you would like to distribute this application's content to.
[string]$DistPointGroup = 'The Dudes Stuff'

#This is the directory where you will download all your files for staging.
#This PS1 should be located in this directory along with your other staging files.
$StagingDir = "G:\Staging\Office365"




$StartDir = Get-Location

$DLDate = get-date -Format yyyyMMdd
$MyDate = get-date -Format MM/dd/yyyy
$DLDir32 = "$StagingDir\O365ProPlus32\$DLDate"
$DLDir64 = "$StagingDir\O365ProPlus64\$DLDate"

if (!(Test-Path $DLDir32)) { New-Item -Path $DLDir32 -ItemType directory }
if (!(Test-Path $DLDir64)) { New-Item -Path $DLDir64 -ItemType directory }


Function Get-O365Urls {
    [CmdletBinding()]
    param()

    $TempDir = $env:TEMP
    $o36532bitXML = "o365client_32bit.xml"
    $ChannelInfoXML = "VersionDescriptor.xml"
    $o365Cab = "ofl.cab"
    $Office365Cab = "http://officecdn.microsoft.com/pr/wsus/ofl.cab"

    #Script
    $workingO365CabPath = "$TempDir\$o365Cab"
    Invoke-WebRequest -Uri $Office365Cab -OutFile "$workingO365CabPath" -UseBasicParsing
    expand.exe "$workingO365CabPath" "$TempDir" -f:$o36532bitXML >$null
    $XMLFile = "$TempDir\$o36532bitXML"
    [xml]$o365BranchPaths = Get-Content $XMLFile
    $DeferredSourceURL = ($o365BranchPaths.UpdateFiles.baseURL | where { $_.branch -eq "Deferred" }).url
    $CurrentSourceURL = ($o365BranchPaths.UpdateFiles.baseURL | where { $_.branch -eq "Current" }).url
    $FRDCSourceURL = ($o365BranchPaths.UpdateFiles.baseURL | where { $_.branch -eq "FirstReleaseDeferred" }).url
    Write-Verbose "Deferred Channel: $DeferredSourceURL`nCurrent Channel: $CurrentSourceURL`nFirst Release Deferred: $FRDCSourceURL"
    Remove-Item $XMLFile -Force -ErrorAction SilentlyContinue
    Remove-Item $workingO365CabPath -Force -ErrorAction SilentlyContinue
    $SourceUrls = New-Object -TypeName PSObject -Prop @{DeferredSourceURL = "$DeferredSourceURL"; CurrentSourceURL = "$CurrentSourceURL"; FRDCSourceURL = "$FRDCSourceURL" }
    Return $SourceUrls
}


Function Get-O365Build {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$SourceURL
    )
    $TempDir = $env:TEMP
    $v32Cab = "v32.cab"
    $o365v32Cab = "Office/Data/v32.cab"
    $V32CabURL = "$SourceURL/$o365v32Cab"
    $workingV32Cab = "$TempDir\$v32Cab"
    $ChannelInfoXML = "VersionDescriptor.xml"
    
    Invoke-WebRequest -Uri $V32CabURL -OutFile $workingV32Cab -UseBasicParsing
    expand.exe "$workingV32Cab" "$TempDir" -f:$ChannelInfoXML >$null
    $XMLFile = "$TempDir\$ChannelInfoXML"
    [xml]$ChannelInfo = Get-Content $XMLFile
    $BuildNumber = $ChannelInfo.Version.Available.Build
    Remove-Item $XMLFile -Force -ErrorAction SilentlyContinue
    Remove-Item $workingV32Cab -Force -ErrorAction SilentlyContinue
    Return $BuildNumber
}


Function Test-OfficeBranch {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$SrcDir,
        [Parameter(Mandatory = $true)]
        [ValidateSet("DC", "FRDC", "CC")]
        [string]$Branch,
        [Parameter(Mandatory = $false)]
        [ValidateSet("32", "64")]
        [string]$Bitness = 32
    )

    switch ($Branch) {
        FRDC { $channel = "FRDCSourceURL"; break }
        CC { $channel = "CurrentSourceURL"; break }
        DC { $channel = "DeferredSourceURL"; break }
    }

    $OfficeBuild = Get-O365Build -SourceURL (Get-O365Urls).$channel

    $BranchSrcDir = "$SrcDir\$Branch\$OfficeBuild\$Bitness"

    Write-Verbose "Application Build Directory: $BranchSrcDir"

    if (!(Test-Path $BranchSrcDir)) {
        Return $false
    }
    Else {
        Return $true
    }
}


Function Download-OfficeProPlusChannel {

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$TargetDirectory,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$SourceConfigXML,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$SetupEXEPath,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [ValidateSet(32, 64)]
        [string]$Bitness,
        [Parameter(Mandatory = $true)]
        [ValidateSet("DC", "FRDC", "CC")]
        [string]$Branch,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$SrcDir
    )

    switch ($Branch) {
        FRDC { $channel = "FirstReleaseDeferred"; break }
        CC { $channel = "Current"; break }
        DC { $channel = "Deferred"; break }
    }
    
    $TempConfig = "$env:TEMP\Config.xml"
    $DownloadPath = "$TargetDirectory\$Branch"

    $FilesExist = Test-OfficeBranch -SrcDir $SrcDir -Branch $Branch -Bitness $Bitness

    if ($FilesExist -eq $false) {
        Copy-Item $SourceConfigXML $TempConfig -Force
        $content = [System.IO.File]::ReadAllText("$TempConfig").Replace("NameChannel", "$channel")
        [System.IO.File]::WriteAllText("$TempConfig", $content)
        $content = [System.IO.File]::ReadAllText("$TempConfig").Replace("Bitness", "$Bitness")
        [System.IO.File]::WriteAllText("$TempConfig", $content)
        $content = [System.IO.File]::ReadAllText("$TempConfig").Replace("DownloadPath", "$DownloadPath")
        [System.IO.File]::WriteAllText("$TempConfig", $content)

        Start-Process -FilePath $SetupEXEPath -ArgumentList "/download $TempConfig" -NoNewWindow -Wait
    }
    else {
        Write-Verbose "Files already exist for the current build of Office 365 ProPlus.  Skipping download of, $Bitness-bit, $channel"
    }
}


Function New-OfficeInstallSourceFiles {

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$DownloadDir,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$CfgMgrSrcDir,
        [Parameter(Mandatory = $true)]
        [ValidateSet("DC", "FRDC", "CC")]
        [string]$Branch,
        [Parameter(Mandatory = $true)]
        [string]$BaseFiles,
        [Parameter(Mandatory = $false)]
        [string]$MyDate = (get-date -Format MM/dd/yyyy),
        [Parameter(Mandatory = $false)]
        [ValidateSet("32", "64")]
        [string]$Bitness = 32
    )


    switch ($Branch) {
        FRDC { $OfficeBuild = Get-O365Build -SourceURL (Get-O365Urls).FRDCSourceURL; $channel = "FirstReleaseDeferred"; break }
        CC { $OfficeBuild = Get-O365Build -SourceURL (Get-O365Urls).CurrentSourceURL; $channel = "Current"; break }
        DC { $OfficeBuild = Get-O365Build -SourceURL (Get-O365Urls).DeferredSourceURL; $channel = "Deferred"; break }
    }

    $StageDir = "$DownloadDir\$Branch"
    $BranchSrcDir = "$CfgMgrSrcDir\$Branch\$OfficeBuild\$Bitness"

    Write-Verbose "Staging Source Directory: $StageDir"
    Write-Verbose "Application Build Directory: $BranchSrcDir"
    Write-Verbose "Office Build: $OfficeBuild"


    
    if ((Test-OfficeBranch -SrcDir $CfgMgrSrcDir -Branch $Branch -Bitness $Bitness) -eq $false) {
        
        Write-Verbose "Creating Directory: $BranchSrcDir"
        New-Item -Path $BranchSrcDir -ItemType directory

        #xcopy "$StagingDir\BaseFiles\*" "$FRDCSrcDir\"
        Write-Verbose "Copying: $BaseFiles\* To: $BranchSrcDir"
        copy-item "$BaseFiles\*" $BranchSrcDir -force -recurse -verbose 

        #xcopy "$FRDCStage\Office" "$FRDCSrcDir\Files\Office"
        Write-Verbose "Copying: $StageDir\Office\Data*  To: $BranchSrcDir\Files\Office"
        copy-item "$StageDir\Office\Data" "$BranchSrcDir\Files\Office\Data" -force -recurse -verbose 


        Write-Verbose "Rewriting versions and dates inside of source file: $BranchSrcDir\Deploy-Application.ps1"
        $content = [System.IO.File]::ReadAllText("$BranchSrcDir\Deploy-Application.ps1").Replace("16.0.6965.2066", "$OfficeBuild")
        [System.IO.File]::WriteAllText("$BranchSrcDir\Deploy-Application.ps1", $content)
        $content = [System.IO.File]::ReadAllText("$BranchSrcDir\Deploy-Application.ps1").Replace("07/19/2017", "$MyDate")
        [System.IO.File]::WriteAllText("$BranchSrcDir\Deploy-Application.ps1", $content)

        foreach ($confxml in (Get-Item "$BranchSrcDir\Files\*.xml").fullname) {
            Write-Verbose "Rewriting versions and dates inside of source file: $confxml"
            $content = [System.IO.File]::ReadAllText("$confxml").Replace("Deferred", "$channel")
            [System.IO.File]::WriteAllText("$confxml", $content)
        }
    }
    Else {
        Write-Verbose "Office Branch/Channel already compiled: $Branch,$channel,$OfficeBuild"
    }
}


Function Set-Stage {

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$Stage,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$StagingDir
    )
    If ($Stage -ieq "Reset") {
        If (Test-Path -Path "$StagingDir\Stage.log") { Remove-Item -Path "$StagingDir\Stage.log" }
    }
    else {
        Write-Verbose "Completing stage [$Stage].."
        Add-Content -Value "Stage: $Stage" -Path "$StagingDir\Stage.log"
    }
}


Function Get-Stage {

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$StagingDir
    )
    If (Test-Path -Path "$StagingDir\Stage.log") {
        ((Get-Content -Tail 1 -Path "$StagingDir\Stage.log") -split ": ")[1]
    }
    else {
        return 0
    }
}


Function Import-SCCMAppLibraries {
    [CmdletBinding()]
    param()
    Write-Verbose "Adding Microsoft.ConfigurationManagement.ApplicationManagement.dll..."
    if ($env:SMS_ADMIN_UI_PATH) { Write-Verbose "Console Install found at: ${env:SMS_ADMIN_UI_PATH} -- Continuing library load..." }
    else {
        write-error "Console install not found. Environmental variable SMS_ADMIN_UI_PATH not found. Unable to load required libraries."
        return "Failure"
    }
    Try {
        Add-Type -Path (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH) Microsoft.ConfigurationManagement.ApplicationManagement.dll)
        Add-Type -Path (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH) Microsoft.ConfigurationManagement.ApplicationManagement.MsiInstaller.dll)
        Add-Type -Path (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH) Microsoft.ConfigurationManagement.ManagementProvider.dll)
        Add-Type -Path (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH) Microsoft.ConfigurationManagement.ApplicationManagement.Extender.dll)
        Add-Type -Path (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH) DcmObjectModel.dll)
        Add-Type -Path (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH) AdminUI.WqlQueryEngine.dll)
        Add-Type -Path (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH) AdminUI.AppManFoundation.dll)
    }
    Catch {
        write-error "Failed to add all of the required libraries.  $($_.Exception.ErrorID) --- $($_.Exception.Message)"
    }
}


function New-SCCMSession {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false, ParameterSetName = "local")]
        [Parameter(Mandatory = $true, ParameterSetName = "withcred")]
        [ValidateNotNullorEmpty()]
        [string]$Server = "localhost",
        [Parameter(Mandatory = $false, ParameterSetName = "withcred")]
        [ValidateNotNullorEmpty()]
        [string]$User,
        [Parameter(Mandatory = $true, ParameterSetName = "withcred")]
        [ValidateNotNullorEmpty()]
        [string]$Password,
        [Parameter(Mandatory = $false)]
        $factory = (New-Object Microsoft.ConfigurationManagement.AdminConsole.AppManFoundation.ApplicationFactory)
    )
    $Session = New-Object Microsoft.ConfigurationManagement.ManagementProvider.WqlQueryEngine.WqlConnectionManager
    if ($Session -ne $null) {
        Write-Verbose "connection object Created"
    }
    else {
        Write-Verbose "Connection object Creation failed"
    }
 
    # Connect with credentials if not running on the server itself
    if (($env:computername -ieq $Server) -or ($Server -ieq "localhost")) {
        Write-Verbose  "Local WQL Connection Made:"
        [void]$Session.Connect($Server)
    }
    elseif ($User) {
        Write-Verbose "Remote WQL Connection Made: " + $Server
        [void]$Session.Connect($Server, $User, $Password)
    }
    else {
        Write-Verbose "Remote WQL Connection Made using current credentials: " + $Server
        [void]$Session.Connect($Server)
    }

    return $Session
}


Function New-SCCMAppBase {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$Title,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$Publisher,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$AppVersion,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$LocalTitle = $Title,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullorEmpty()]
        [string]$Language = "en-US",
        [Parameter(Mandatory = $false)]
        [ValidateScript( { Test-Path -LiteralPath $_ -PathType 'leaf' })]
        [string]$Icon

    )


    # create the application.
    $app = New-Object Microsoft.ConfigurationManagement.ApplicationManagement.Application

    # set the application properties.
    $app.Title = "$Title"
    $app.Publisher = "$Publisher"
    $app.SoftwareVersion = "$AppVersion"

    # prepare the localised display info.
    $info = New-Object Microsoft.ConfigurationManagement.ApplicationManagement.AppDisplayInfo

    # set the localised application name.
    $info.Title = "$LocalTitle"
    $info.Language = $Language
    if ($Icon) {
        $SCCMIcon = New-Object Microsoft.ConfigurationManagement.ApplicationManagement.Icon
        $SCCMIcon.data = [System.IO.File]::ReadAllBytes("$Icon")
        $info.Icon = $SCCMIcon
    }

    # save the display properties.
    $app.DisplayInfo.Add($info)
    
    Return $app

}


Function New-SCCMInstaller {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$InstallCmd,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$UninstallCmd,
        [Parameter(Mandatory = $true)]
        #[ValidateScript({ Test-Path -LiteralPath $_ -PathType 'Container' })]
        [string]$SourcePath,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullorEmpty()]
        [int]$RunTime = 15,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullorEmpty()]
        [int]$MaxRunTime = 60
    )

    # create a script-based installer.
    $installer = New-Object Microsoft.ConfigurationManagement.ApplicationManagement.ScriptInstaller
    $installer.InstallCommandLine = $InstallCmd
    $installer.UninstallCommandLine = $UninstallCmd
    $installer.MaxExecuteTime = $MaxRunTime
    $installer.ExecuteTime = $RunTime
	
    # reference the content source.
    $content = [Microsoft.ConfigurationManagement.ApplicationManagement.ContentImporter]::CreateContentFromFolder($SourcePath)
    $installer.Contents.Add($content)

    $reference = New-Object Microsoft.ConfigurationManagement.ApplicationManagement.ContentRef
    $reference.Id = $content.Id
    $installer.InstallContent = $reference
	
    $installer.RequiresLogOn = $null
    $installer.ExecutionContext = 'System'
    $installer.MachineInstall = $true
    $installer.RequiresUserInteraction = $true

    Return $installer
}


Function New-SCCMAppDetector {
    [CmdletBinding()]
    param()
    $detector = New-Object Microsoft.ConfigurationManagement.ApplicationManagement.EnhancedDetectionMethod
    Return $detector
}


Function New-SCCMRegistryDetectionMethod {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("LocalMachine", "ClassesRoot", "CurrentConfig", "CurrentUser", "Users")]
        [string]$RootKey,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$RegKey,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$RegValue,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$RegData,
        [Parameter(Mandatory = $true)]
        [ValidateSet("String", "Int64", "Version")]
        [string]$RegType = "String",
        [Parameter(Mandatory = $true)]
        [ValidateSet("And", "Or", "IsEquals", "NotEquals", "GreaterThan", "Between", "GreaterEquals", "LessEquals", "BeginsWith", "NotBeginsWith", "EndsWith", "NotEndsWith", "Contains", "NotContains", "AllOf", "OneOf", "NoneOf", "SetEquals", "SupportedOperators")]
        [string]$Operator,
        [Parameter(Mandatory = $false)]
        [bool]$Is64Bit = $false,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        $App, #[Microsoft.ConfigurationManagement.ApplicationManagement.Application]$App,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        $Installer, #[Microsoft.ConfigurationManagement.ApplicationManagement.ScriptInstaller]$installer
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        $detector #[Microsoft.ConfigurationManagement.ApplicationManagement.EnhancedDetectionMethod]$detector
    )

    #For some reason, registry detection methods return in an array with a couple members.  Returning the [1] member gets the actual detection item.
    #So, I just nest the detector function inside this function and then return the proper array member.
    Function New-RegDetection {
        [CmdletBinding()]
        Param (
            [Parameter(Mandatory = $true)][ValidateSet("LocalMachine", "ClassesRoot", "CurrentConfig", "CurrentUser", "Users")][string]$RootKey,
            [Parameter(Mandatory = $true)]
            [ValidateNotNullorEmpty()]
            [string]$RegKey,
            [Parameter(Mandatory = $true)]
            [ValidateNotNullorEmpty()]
            [string]$RegValue,
            [Parameter(Mandatory = $true)]
            [ValidateNotNullorEmpty()]
            [string]$RegData,
            [Parameter(Mandatory = $true)]
            [ValidateSet("String", "Int64", "Version")]
            [string]$RegType = "String",
            [Parameter(Mandatory = $true)]
            [ValidateSet("And", "Or", "IsEquals", "NotEquals", "GreaterThan", "Between", "GreaterEquals", "LessEquals", "BeginsWith", "NotBeginsWith", "EndsWith", "NotEndsWith", "Contains", "NotContains", "AllOf", "OneOf", "NoneOf", "SetEquals", "SupportedOperators")]
            [string]$Operator,
            [Parameter(Mandatory = $false)]
            [bool]$Is64Bit = $false,
            [Parameter(Mandatory = $true)]
            [ValidateNotNullorEmpty()]
            $App, #[Microsoft.ConfigurationManagement.ApplicationManagement.Application]$App,
            [Parameter(Mandatory = $true)]
            [ValidateNotNullorEmpty()]
            $Installer, #[Microsoft.ConfigurationManagement.ApplicationManagement.ScriptInstaller]$installer
            [Parameter(Mandatory = $true)]
            [ValidateNotNullorEmpty()]
            $Detector #[Microsoft.ConfigurationManagement.ApplicationManagement.EnhancedDetectionMethod]$detector
        )

        # use an enhanced detection method (registry).
        $Installer.DetectionMethod = [Microsoft.ConfigurationManagement.ApplicationManagement.DetectionMethod]::Enhanced
        #$detector = New-Object Microsoft.ConfigurationManagement.ApplicationManagement.EnhancedDetectionMethod
        #$Type = [Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ConfigurationItemPartType]::RegistryKey
        $oRegSetting = new-object Microsoft.ConfigurationManagement.DesiredConfigurationManagement.RegistrySetting(@($null))
        $oRegSetting.Name
        $oRegSetting.RootKey = $RootKey
        $oRegSetting.Key = $RegKey
        $oRegSetting.Is64Bit = $Is64Bit
        $oRegSetting.ValueName = $RegValue
        $oRegSetting.CreateMissingPath = $true
        $oRegSetting.SettingDataType = [Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType]::$RegType #Int64/String/Version/?
        $Detector.Settings.Add($oRegSetting)

        # configure the setting reference.
        $reference = New-Object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.SettingReference($App.Scope, $App.Name, $App.Version, $oRegSetting.LogicalName, $oRegSetting.SettingDataType, $oRegSetting.SourceType, [bool]0)
        $reference.MethodType = [Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ConfigurationItemSettingMethodType]::Value
        #$reference.PropertyPath = "RegistryValueExists"
    

        # create the version comparison value.
        $value = New-Object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.ConstantValue($RegData, $oRegSetting.SettingDataType)
    
        # configure the comparison operands.
        $operands = New-Object Microsoft.ConfigurationManagement.DesiredConfigurationManagement.CustomCollection``1[[Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.ExpressionBase]]
        $operands.Add($reference)
        $operands.Add($value)

        # build the comparison expression.
        $expression = New-Object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.Expression([Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator]::$Operator, $operands)
    
        Return $expression
    }

    $expression = New-RegDetection  -RootKey $RootKey -RegKey $RegKey -RegValue $RegValue -RegData $RegData -RegType $RegType -Operator $Operator -Is64Bit $Is64Bit -App $App -Installer $Installer -detector $Detector
    return $expression[1]

}


Function Set-SCCMDetector {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ParameterSetName = "expr1")]
        [Parameter(Mandatory = $true, ParameterSetName = "expr2")]
        $expression,
        [Parameter(Mandatory = $true, ParameterSetName = "expr1")]
        [Parameter(Mandatory = $true, ParameterSetName = "expr2")]
        $detector,
        [Parameter(Mandatory = $false, ParameterSetName = "expr2")]
        $expression2,
        [Parameter(Mandatory = $true, ParameterSetName = "expr2")]
        [ValidateSet("And", "Or", "IsEquals", "NotEquals", "GreaterThan", "Between", "GreaterEquals", "LessEquals", "BeginsWith", "NotBeginsWith", "EndsWith", "NotEndsWith", "Contains", "NotContains", "AllOf", "OneOf", "NoneOf", "SetEquals", "SupportedOperators")]
        [string]$Operator
    )

    If ($expression2) {
        $operands = New-Object Microsoft.ConfigurationManagement.DesiredConfigurationManagement.CustomCollection``1[[Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.ExpressionBase]]
        $operands.Add($expression)
        $operands.Add($expression2)

        # build the comparison expression.
        $FinalExpression = New-Object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.Expression([Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator]::$Operator, $operands)
    }
    else {
        $FinalExpression = $expression
    }
    # save the detection rule.
    $rule = New-Object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.Rule("IsInstalledRule", [Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.NoncomplianceSeverity]::None, $null, $FinalExpression)

    # associate the detection rule with the detection method.
    $detector.Rule = $rule 
    
    #Return $detector
}


Function Add-SCCMApplication {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false)]
        $Factory = (New-Object Microsoft.ConfigurationManagement.AdminConsole.AppManFoundation.ApplicationFactory),
        [Parameter(Mandatory = $true)]
        $Application, #[Microsoft.ConfigurationManagement.ApplicationManagement.Application]$App,
        [Parameter(Mandatory = $true)]
        $Installer, #[Microsoft.ConfigurationManagement.ApplicationManagement.ScriptInstaller]$installer
        [Parameter(Mandatory = $true)]
        $Detector,
        [Parameter(Mandatory = $true)]
        $Wrapper
    )

    # associate the detector with the installer.
    $Installer.EnhancedDetectionMethod = $Detector

    # create the deployment type.
    $Deployment = New-Object Microsoft.ConfigurationManagement.ApplicationManagement.DeploymentType($Installer, [Microsoft.ConfigurationManagement.ApplicationManagement.ScriptInstaller]::TechnologyId, [Microsoft.ConfigurationManagement.ApplicationManagement.NativeHostingTechnology]::TechnologyId)

    # name the deployment type after the application.
    $Deployment.Title = $Application.Title

    # associate the deployment type with the application.
    $Application.DeploymentTypes.Add($Deployment)

    # save the application.
    $Wrapper.InnerAppManObject = $Application
    $factory.PrepareResultObject($Wrapper)
    $Wrapper.InnerResultObject.Put()

}


Function New-CMAppRegDetectionMethod {
    [CmdletBinding()]
    Param(
        [Parameter (Mandatory = $false)]
        [ValidateSet("LocalMachine", "ClassesRoot", "CurrentConfig", "CurrentUser", "Users")]
        [string]$RegDetectionRoot = "LocalMachine",
        [Parameter (Mandatory = $true)]
        [string]$RegDetectionKey,
        [Parameter (Mandatory = $true)]
        [string]$RegDetectionValue,
        [Parameter (Mandatory = $true)]
        [string]$RegDetectionData,
        [Parameter (Mandatory = $false)]
        [ValidateSet("String", "Int64", "Version")]
        [string]$RegDetectionType = "String",
        [Parameter (Mandatory = $false)]
        [ValidateSet("And", "Or", "IsEquals", "NotEquals", "GreaterThan", "Between", "GreaterEquals", "LessEquals", "BeginsWith", "NotBeginsWith", "EndsWith", "NotEndsWith", "Contains", "NotContains", "AllOf", "OneOf", "NoneOf", "SetEquals", "SupportedOperators")]
        [string]$RegDetectionOper = "IsEquals"
    )
    @{
        "RegDetectionRoot"  = "$RegDetectionRoot";
        "RegDetectionKey"   = "$RegDetectionKey";
        "RegDetectionValue" = "$RegDetectionValue";
        "RegDetectionData"  = "$RegDetectionData";
        "RegDetectionType"  = "$RegDetectionType";
        "RegDetectionOper"  = "$RegDetectionOper"
    }
}


function Write-Log {
    [CmdletBinding()] 
    Param (
        [Parameter(Mandatory = $false)]
        $Message,
 
        [Parameter(Mandatory = $false)]
        $ErrorMessage,
 
        [Parameter(Mandatory = $false)]
        $Component,
 
        [Parameter(Mandatory = $false, HelpMessage = "1 = Normal, 2 = Warning (yellow), 3 = Error (red)")]
        [ValidateSet(1, 2, 3)]
        [int]$Type,
		
        [Parameter(Mandatory = $false, HelpMessage = "Size in KB")]
        [int]$LogSizeKB = 512,

        [Parameter(Mandatory = $true)]
        $LogFile
    )
    <#
    Type: 1 = Normal, 2 = Warning (yellow), 3 = Error (red)
    #>
    $LogLength = $LogSizeKB * 1024
    try {
        $log = Get-Item $LogFile -ErrorAction Stop
        If (($log.length) -gt $LogLength) {
            $Time = Get-Date -Format "HH:mm:ss.ffffff"
            $Date = Get-Date -Format "MM-dd-yyyy"
            $LogMessage = "<![LOG[Closing log and generating new log file" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"1`" thread=`"`" file=`"`">"
            $LogMessage | Out-File -Append -Encoding UTF8 -FilePath $LogFile
            Move-Item -Path "$LogFile" -Destination "$($LogFile.TrimEnd('g'))_" -Force
        }
    }
    catch { Write-Verbose "Nothing to move or move failed." }

    $Time = Get-Date -Format "HH:mm:ss.ffffff"
    $Date = Get-Date -Format "MM-dd-yyyy"
 
    if ($ErrorMessage -ne $null) { $Type = 3 }
    if ($Component -eq $null) { $Component = " " }
    if ($Type -eq $null) { $Type = 1 }
 
    $LogMessage = "<![LOG[$Message $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"$Type`" thread=`"`" file=`"`">"
    $LogMessage | Out-File -Append -Encoding UTF8 -FilePath $LogFile
}


###########################################################################
################       Download Content from the CDN       ################
###########################################################################

$SourceUrls = Get-O365Urls
$DCBuild = Get-O365Build -SourceURL $SourceUrls.DeferredSourceURL
#$FRDCBuild = Get-O365Build -SourceURL $SourceUrls.FRDCSourceURL
#$CCBuild = Get-O365Build -SourceURL $SourceUrls.CurrentSourceURL


[int]$Stage = Get-Stage -StagingDir $StagingDir
Write-verbose ("Beginning script at stage [" + ([int]$Stage + 1) + "]...")

if ($Stage -lt 1) {
    Write-verbose "Downloading 32bit Deferred Channel"
    Download-OfficeProPlusChannel -TargetDirectory $DLDir32 -SourceConfigXML "$StagingDir\Config.xml" -SetupEXEPath "$StagingDir\setup.exe" -Bitness 32 -Branch DC -SrcDir $AppSourceFiles
    Set-Stage -Stage 1 -StagingDir $StagingDir
}

if ($Stage -lt 2) {
    Write-verbose "Downloading 64bit Deferred Channel"
    Download-OfficeProPlusChannel -TargetDirectory $DLDir64 -SourceConfigXML "$StagingDir\Config.xml" -SetupEXEPath "$StagingDir\setup.exe" -Bitness 64 -Branch DC -SrcDir $AppSourceFiles
    Set-Stage -Stage 2 -StagingDir $StagingDir
}

if ($Stage -lt 3) {
    New-OfficeInstallSourceFiles -DownloadDir $DLDir32 -CfgMgrSrcDir $AppSourceFiles -Branch DC -BaseFiles "$StagingDir\BaseFiles32"
    Set-Stage -Stage 3 -StagingDir $StagingDir
}

if ($Stage -lt 4) {
    New-OfficeInstallSourceFiles -DownloadDir $DLDir64 -CfgMgrSrcDir $AppSourceFiles -Branch DC -BaseFiles "$StagingDir\BaseFiles64" -Bitness 64
    Set-Stage -Stage 4 -StagingDir $StagingDir
}


Set-Stage -Stage Reset -StagingDir $StagingDir
Write-Verbose "File Downloads and source file building complete.."
Write-Verbose "Semi-annual build downolad complete.  Version: $DCBuild"

Write-Verbose "***********************************All Downloads Complete***********************************"
Write-Verbose "********************************************************************************************"
###########################################################################
##################      Build Application In SCCM      ####################
###########################################################################

Write-Verbose "*******************************Starting CM Application Build********************************"

$ApplicationVersion = $DCBuild
$AppSourceFiles = "$($AppSourceFiles)\DC\$($ApplicationVersion)\32"

Write-Verbose "Building: $AppName"
Write-Verbose "Version: $ApplicationVersion"
Write-Verbose "Source Files: $AppSourceFiles"

Write-Verbose "Checking for existing application for $AppName $ApplicationVersion"

Try {
    Import-Module (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH) "ConfigurationManager.psd1") -Verbose:$false
    CD "$($CMSite):"
}
Catch {
    Set-location $StartDir
    #Write-Log -Message "Failed to load module and change to the Site Code provider.  $($_.Exception.ErrorID) --- $($_.Exception.Message)" -
    Write-Error "Failed to load module and change to the Site Code provider.  $($_.Exception.ErrorID) --- $($_.Exception.Message)"
    return "Failure"
}
if (Get-CMApplication -Name "$AppName $ApplicationVersion") {
    Set-location $StartDir
    Return "Application `"$AppName $ApplicationVersion`" already exists."
}

Write-Verbose "App does not already exist.  Building $AppName Version: $ApplicationVersion"


Import-SCCMAppLibraries #-Verbose

Write-Verbose "Initialising management scope for creating CM Application."
try {
    $factory = New-Object Microsoft.ConfigurationManagement.AdminConsole.AppManFoundation.ApplicationFactory
    $Session = New-SCCMSession -Server "$CMServer"
    $wrapper = [Microsoft.ConfigurationManagement.AdminConsole.AppManFoundation.AppManWrapper]::Create($Session, $factory)
}
Catch {
    Set-location $StartDir
    Write-Error "Failed to build Factory, Session, or Wrapper.  $($_.Exception.ErrorID) --- $($_.Exception.Message)"
    return "Failure"
}

Write-Verbose "Build Hash Tables for the Registry detections. Detecting on the version and the Product ID"
$RegRule1 = New-CMAppRegDetectionMethod -RegDetectionRoot LocalMachine -RegDetectionKey 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -RegDetectionValue 'ClientVersionToReport' -RegDetectionData $ApplicationVersion -RegDetectionType Version -RegDetectionOper GreaterEquals
$RegRule2 = New-CMAppRegDetectionMethod -RegDetectionRoot LocalMachine -RegDetectionKey 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -RegDetectionValue 'ProductReleaseIds' -RegDetectionData 'O365ProPlusRetail' -RegDetectionType String -RegDetectionOper Contains

Write-verbose "Building App, Installer, or base Detector...."
try {
    If (Test-Path $IconPath) {
        $app = New-SCCMAppBase -Title "$AppName $ApplicationVersion" -Publisher "$Publisher" -AppVersion "$ApplicationVersion" -LocalTitle "$AppName" -Icon "$IconPath"
    }
    else {
        $app = New-SCCMAppBase -Title "$AppName $ApplicationVersion" -Publisher "$Publisher" -AppVersion "$ApplicationVersion" -LocalTitle "$AppName"
    }
    $installer1 = New-SCCMInstaller -InstallCmd "$InstallCmd" -UninstallCmd "$UninstallCmd" -SourcePath "$AppSourceFiles"
    $MyDetector1 = New-SCCMAppDetector
}
Catch {
    Set-location $StartDir
    Write-Error "Failed to build App, Installer, or base Detector.  $($_.Exception.ErrorID) --- $($_.Exception.Message)"
    return "Failure"
}

Write-Verbose "Building Registry detection for $AppName with version of $ApplicationVersion"
try {
    $RegExpression1 = New-SCCMRegistryDetectionMethod -RootKey $RegRule1.RegDetectionRoot -RegKey "$($Regrule1.RegDetectionKey)" -RegValue "$($Regrule1.RegDetectionValue)" -RegData "$($RegRule1.RegDetectionData)" -RegType "$($Regrule1.RegDetectionType)" -Operator $($RegRule1.RegDetectionOper) -Is64Bit $true -App $app -Installer $installer1 -detector $MyDetector1
    $RegExpression2 = New-SCCMRegistryDetectionMethod -RootKey $RegRule2.RegDetectionRoot -RegKey "$($Regrule2.RegDetectionKey)" -RegValue "$($Regrule2.RegDetectionValue)" -RegData "$($RegRule2.RegDetectionData)" -RegType "$($Regrule2.RegDetectionType)" -Operator $($RegRule2.RegDetectionOper) -Is64Bit $true -App $app -Installer $installer1 -detector $MyDetector1
    Set-SCCMDetector -detector $MyDetector1 -expression $RegExpression1 -expression2 $RegExpression2 -Operator And
}
Catch {
    Set-location $StartDir
    Write-Error "Failed to build app detection expressions.  $($_.Exception.ErrorID) --- $($_.Exception.Message)"
    return "Failure"
}

Write-verbose "Putting it all together and creating application with name [$($app.Title)]"
try {
    Add-SCCMApplication -Factory $factory -Application $app -Installer $installer1 -Detector $MyDetector1 -Wrapper $wrapper
    #    Add-SCCMApplication -Factory $factory -Application $app -Installer $installer2 -Detector $MyDetector2 -Wrapper $wrapper
}
Catch {
    Set-location $StartDir
    Write-Error "Failed to build Application.  $($_.Exception.ErrorID) --- $($_.Exception.Message)"
    return "Failure"
}
Write-Verbose "Application [$($app.Title)] has been successfully created in SCCM."
Write-Verbose "Distributing content for application [$($app.Title)] to [$DistPointGroup]."
Start-CMContentDistribution -ApplicationName "$($App.Title)" -DistributionPointGroupName $DistPointGroup
Set-location $StartDir
Write-Host "Application [$($app.Title)] has been successfully created and distributed in SCCM."
