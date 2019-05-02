# Deployment and Management of BitLocker in an Enterprise

The code and files for the [BitLocker](https://sched.co/N6eN) session.


* Create-BitLockerCollections.ps1 -- Adds 5 collections into CM for your BitLocker Deployment.  They are all put in a folder called 'BitLocker-MMS2019' created in the root of Device Collections.

    * OS - Workstation
    * Bitlocker - All Mobile Windows Computers
    * Bitlocker - TPM Enabled -> Install MBAM
    * Bitlocker - TPM Enabled - MBAM Installed -> Encrypt
    * Bitlocker - TPM Enabled - MBAM Installed - Encrypted -> TPM Only