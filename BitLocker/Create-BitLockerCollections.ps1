write-host "Creating Collections for BitLocker deployments (MMS 2019!)..."
$CollectionFolder = 'BitLocker-MMS2019'

#Load Configuration Manager PowerShell Module
Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)

#Get SiteCode
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-location $SiteCode":"

$FullFolderPath = $SiteCode.Name+":\DeviceCollection\"+$CollectionFolder

New-Item -Name "$CollectionFolder" -Path $($SiteCode.Name+":\DeviceCollection")

$Collection = New-CMDeviceCollection -Name 'OS - Workstation' -LimitingCollectionId 'SMS00001' -RefreshType Both
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $Collection.CollectionID -RuleName 'Workstations' -QueryExpression 'select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.OperatingSystemNameandVersion like "Microsoft Windows NT%Workstation%"'
Move-CMObject -FolderPath $FullFolderPath -InputObject (Get-CMDeviceCollection -CollectionId $Collection.CollectionID)

$Collection = New-CMDeviceCollection -Name 'Bitlocker - All Mobile Windows Computers' -LimitingCollectionId $Collection.CollectionID -RefreshType Both
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $Collection.CollectionID -RuleName 'All Laptop Chassis' -QueryExpression 'select distinct SMS_R_System.ResourceId, SMS_R_System.ResourceType, SMS_R_System.Name, SMS_R_System.SMSUniqueIdentifier, SMS_R_System.ResourceDomainORWorkgroup, SMS_R_System.Client from  SMS_R_System inner join SMS_G_System_SYSTEM_ENCLOSURE on SMS_G_System_SYSTEM_ENCLOSURE.ResourceID = SMS_R_System.ResourceId where SMS_G_System_SYSTEM_ENCLOSURE.ChassisTypes = "8" or SMS_G_System_SYSTEM_ENCLOSURE.ChassisTypes = "9" or SMS_G_System_SYSTEM_ENCLOSURE.ChassisTypes = "10" or SMS_G_System_SYSTEM_ENCLOSURE.ChassisTypes = "12" or SMS_G_System_SYSTEM_ENCLOSURE.ChassisTypes = "30" or SMS_G_System_SYSTEM_ENCLOSURE.ChassisTypes = "31" or SMS_G_System_SYSTEM_ENCLOSURE.ChassisTypes = "32" order by SMS_R_System.Name'
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $Collection.CollectionID -RuleName 'All Laptop Chassis - Historical' -QueryExpression 'select distinct SMS_R_System.ResourceId, SMS_R_System.ResourceType, SMS_R_System.Name, SMS_R_System.SMSUniqueIdentifier, SMS_R_System.ResourceDomainORWorkgroup, SMS_R_System.Client from  SMS_R_System inner join SMS_GH_System_SYSTEM_ENCLOSURE on SMS_GH_System_SYSTEM_ENCLOSURE.ResourceID = SMS_R_System.ResourceId inner join SMS_GEH_System_SYSTEM_ENCLOSURE on SMS_GEH_System_SYSTEM_ENCLOSURE.ResourceID = SMS_R_System.ResourceId where SMS_GH_System_SYSTEM_ENCLOSURE.ChassisTypes = "8" or SMS_GH_System_SYSTEM_ENCLOSURE.ChassisTypes = "9" or SMS_GH_System_SYSTEM_ENCLOSURE.ChassisTypes = "10" or SMS_GEH_System_SYSTEM_ENCLOSURE.ChassisTypes = "10" or SMS_GEH_System_SYSTEM_ENCLOSURE.ChassisTypes = "9" or SMS_GEH_System_SYSTEM_ENCLOSURE.ChassisTypes = "8" or SMS_GH_System_SYSTEM_ENCLOSURE.ChassisTypes = "12" or SMS_GEH_System_SYSTEM_ENCLOSURE.ChassisTypes = "12" or SMS_GH_System_SYSTEM_ENCLOSURE.ChassisTypes = "30" or SMS_GEH_System_SYSTEM_ENCLOSURE.ChassisTypes = "30" or SMS_GH_System_SYSTEM_ENCLOSURE.ChassisTypes = "31" or SMS_GEH_System_SYSTEM_ENCLOSURE.ChassisTypes = "31" or SMS_GH_System_SYSTEM_ENCLOSURE.ChassisTypes = "32" or SMS_GEH_System_SYSTEM_ENCLOSURE.ChassisTypes = "32"'
Move-CMObject -FolderPath $FullFolderPath -InputObject (Get-CMDeviceCollection -CollectionId $Collection.CollectionID)

$Collection = New-CMDeviceCollection -Name 'Bitlocker - TPM Enabled -> Install MBAM' -LimitingCollectionId $Collection.CollectionID -RefreshType Both
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $Collection.CollectionID -RuleName 'TPM Enabled and Active' -QueryExpression 'select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_TPM on SMS_G_System_TPM.ResourceID = SMS_R_System.ResourceId where (SMS_G_System_TPM.SpecVersion like "1.2%" and SMS_G_System_TPM.IsActivated_InitialValue = 1) or (SMS_G_System_TPM.SpecVersion like "2.0%" and SMS_G_System_TPM.IsActivated_InitialValue = 1)'
Move-CMObject -FolderPath $FullFolderPath -InputObject (Get-CMDeviceCollection -CollectionId $Collection.CollectionID)

$Collection = New-CMDeviceCollection -Name 'Bitlocker - TPM Enabled - MBAM Installed -> Encrypt' -LimitingCollectionId $Collection.CollectionID -RefreshType Both
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $Collection.CollectionID -RuleName 'MBAM Installed' -QueryExpression 'select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_INSTALLED_SOFTWARE on SMS_G_System_INSTALLED_SOFTWARE.ResourceID = SMS_R_System.ResourceId where SMS_G_System_INSTALLED_SOFTWARE.ProductName = "MDOP MBAM" and SMS_G_System_INSTALLED_SOFTWARE.ProductVersion >= "2.5.1135.0"'
Move-CMObject -FolderPath $FullFolderPath -InputObject (Get-CMDeviceCollection -CollectionId $Collection.CollectionID)

$Collection = New-CMDeviceCollection -Name 'Bitlocker - TPM Enabled - MBAM Installed - Encrypted -> TPM Only' -LimitingCollectionId $Collection.CollectionID -RefreshType Both
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $Collection.CollectionID -RuleName 'Encrypted C Drive' -QueryExpression 'select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_ENCRYPTABLE_VOLUME on SMS_G_System_ENCRYPTABLE_VOLUME.ResourceId = SMS_R_System.ResourceId where SMS_G_System_ENCRYPTABLE_VOLUME.ProtectionStatus = 1 and SMS_G_System_ENCRYPTABLE_VOLUME.DriveLetter = "C:"'
Move-CMObject -FolderPath $FullFolderPath -InputObject (Get-CMDeviceCollection -CollectionId $Collection.CollectionID)

Write-Host "A total of 5 collections should have been created in $FullFolderPath"