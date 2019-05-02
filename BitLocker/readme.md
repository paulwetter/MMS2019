# Deployment and Management of BitLocker in an Enterprise

The code and files for the [BitLocker](https://sched.co/N6eN) session.


### Create-BitLockerCollections.ps1
Adds 5 collections into CM for your BitLocker Deployment.  They are all put in a folder called 'BitLocker-MMS2019' created in the root of Device Collections.

* **OS - Workstation**
* **Bitlocker - All Mobile Windows Computers** -- Deploy Script to Enable and Activate TPM to this Collection
* **Bitlocker - TPM Enabled -> Install MBAM** -- Delpoys the Installer for the MBAM Client
* **Bitlocker - TPM Enabled - MBAM Installed -> Encrypt** -- No actual deployments.  Acts as more of a placeholder to see the difference between those ready to encrypt and those encrypted.  A system is ready to encrypt when the TPM is enabled and active and the MBAM agent is installed.
* **Bitlocker - TPM Enabled - MBAM Installed - Encrypted -> TPM Only** -- Checks for and adds the TPM Only Key Protector

### Enable-DellTpm.ps1
This script will use the *Dell Command | Configure* tool to enable and active the TPM on Dell computers.

### Remove-BLPin.ps1
This script will remove the PIN from an already encrypted computer that still has a PIN with the TPM.  The script does not remove the TPM+PIN protector.  It only Adds a TPM only key protectors.  The TPM/BitLocker will use the path of least resistence and unlock the drive with the TPM only protector.