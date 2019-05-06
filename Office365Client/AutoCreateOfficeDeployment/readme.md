# Automate Building Office Installer

The files and directoreis contained in this directory are used for building an automatic installer for Office 365.  Follow the steps below to complete this.

### Requirements:
* PowerShell 4.0+
* CM Console installed on Executing server (PowerShell module).
* Executing account will need to be able to create applications and distribute content.
* Executing account will need to be able to execute elevated (for Setup.exe).
* Executing account  will need to be able to write to the AppSourceFiles directory specified.

### Setup steps
1. Copy the files and directories to a folder on a drive on your SCCM server or automation server. These include:
    * BaseFiles32
    * BaseFiles64
    * O365ProPlus32
    * O365ProPlus64
    * Config.xml
    * New-OfficeInstaller.ps1
    * office365icon.ico
    * readme.md
2. Open the New-OfficeInstaller.ps1 file in your favorite editor (VSCode).
3. Update the variables in the start of this script to represent your environment.

### Exectution
1. Run New-OfficeInstaller.ps1 from an elevated command prompt.
    * Example:
        * `.\New-OfficeInstaller.ps1`
    * Supports the Verbose switch for troubleshooting
        * `.\New-OfficeInstaller.ps1 -verbose`