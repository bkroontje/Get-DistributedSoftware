#-------------------------------------------------------------------------
# Author      : bkroontje
# FileName    : Get-DistributedSoftware
# Version     : 1.0
# Revision    :
# Created     : June 7, 2021
# Description : Powershell Script that compiles list of installed software for a list of clients into a cvs file:
# Remarks     : Text file of list of clients needed to run script
#             : File Named InstalledSoftware.csv will be saved in the following location: C:\Temp\InstalledSoftware
#             : 
#             :
# Prerequisite: 
#             : 
#             : 
#-------------------------------------------------------------------------






Function Get-DistributedSoftware 
{
     <#
        .SYPNOPSIS
           This module retrieves installed software information from computers

        .DESCRIPTION
          Accepts a .txt file containing a list of client names and retrieves information about
          the software installed on each client. Results are stored in .csv file located at
          c:\temp\InstalledSoftware.csv on local machine
        .PARAMATER $FilePath
            Variable for path of .txt file containing client names

        .EXAMPLE
            Get-DistributedSoftware -FilePath c:\ps\lists\test_clients.txt

        .EXAMPLE
            PS C:\Windows\system32> Move-ClientOU

          cmdlet Get-DistributedSoftware at command pipeline position 1
          Supply values for the following parameters:
          (Type !? for Help.)
          FilePath: c:\ps\lists\hmi_clients.txt   
     
    #>

    Param
    (
        [Parameter(Mandatory=$true, position=1, HelpMessage="Client List file path")]
        [ValidateNotNullOrEmpty()]$FilePath
    )

    $Computers = get-content $FilePath

    #deletes old version of InstalledSoftware.csv if it exists
    $FileName = "C:\temp\InstalledSoftware.csv"
    if (Test-Path $FileName) 
    {
        Remove-Item $FileName
    }
  
    #registry key paths for installed software
    $lmKeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall","SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    $lmReg = [Microsoft.Win32.RegistryHive]::LocalMachine
    

    #foreach computer get software information
    Foreach($Name in $Computers) 
    {

        #test connectivity of each client
        if (!(Test-Connection -ComputerName $Name -count 1 -quiet)) 
        {
            Write-host "Unable to contact $Name. Please verify its network connectivity and try again." -ForegroundColor Red
        } else     
        {
            #enable Remote registry Service on each client to parse registry
            Get-Service -ComputerName $Name -Name RemoteRegistry | Set-Service -StartupType Manual -PassThru| Start-Service

            #master key object to store software registry keys
            $masterKeys = @()
            $remoteLMRegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($lmReg,$Name)
            foreach ($key in $lmKeys) 
            {
                $regKey = $remoteLMRegKey.OpenSubkey($key)
                foreach ($subName in $regKey.GetSubkeyNames()) 
                {
                    foreach($sub in $regKey.OpenSubkey($subName)) 
                    {

                        #properties of masterkey object
                        $masterKeys += (New-Object PSObject -Property @{
                        "ComputerName" = $Name
                        "Name" = $sub.GetValue("displayname")
                        "SystemComponent" = $sub.GetValue("systemcomponent")
                        "ParentKeyName" = $sub.GetValue("parentkeyname")
                        "Version" = $sub.GetValue("DisplayVersion")
                        "InstallDate" = $sub.GetValue("InstallDate")
                           
                        })
                    }
                }
            }
       
            #filters results and saves in .csv file at C:\Temp\InstalledSoftware
            $woFilter = {$null -ne $_.name -AND $_.SystemComponent -ne "1" -AND $null -eq $_.ParentKeyName}
            $props = 'ComputerName', 'Name','Version','Installdate'
            $masterKeys = ($masterKeys | Where-Object $woFilter | Select-Object $props | Sort-Object Name)
            $masterKeys | Export-csv C:\temp\InstalledSoftware.csv -Append

            #disables Remote-Registry service
            Get-Service -ComputerName $Name -Name RemoteRegistry | Set-Service -StartupType Manual -PassThru | Stop-Service
        }
    }   
}
