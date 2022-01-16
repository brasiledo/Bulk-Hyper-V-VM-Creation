  <# 
.SYNOPSIS
Bulk Hyper-V creation tool.  
 
.DESCRIPTION
-User inputted data (EXCEL FILE) used to automate all steps
-Auomates creating VMs in bulk
-Creates powershell scripts that automates the VM build process and copies these scripts to HOST server 
-Once VMs are created, must manually power on VM and complete OS setup
-Once manual part complete, Run first script to enable guest services and copy script files to each VM
-Last part, log in to each VM and run the PS1 file (Automatically, adds to domain, renames, setup net adapter settings, restarts)

.USER_INPUT
-Fill out and save excel sheet (Hyper-V_Setup_Details.xlsx) with headers -- 
Host,SourceOS,SourceData,VMNameHyperV,SwitchName,Memory,Generation,ProcessorCount,VLAN,VHDPath,TargetOS,TargetData,
ServerName,CurrentNetworkAdapterName,NewNetworkAdapterName,IPAddress,Subnet,GatewayAddress,DNS,WINS,Domain,User

Store this script along along with this created excel file in the same folder.  

Examples for parameters

Host - hostname of host
VMNameHyperV - VM Name
ServerName - VM Hostname
CurrentNetworkAdapterName - 'Ethernet'
NewNetworkAdapterName - ex. 'Web'
IPAddress - Machine IP
Subnet - Machine Subnet
GatewayAddress- Machine Gateway
DNS - Machine DNS
WINS - Machine WINS
Domain - Domain for VM
User - Domain\username (used for adding machine to domain)
SourceOS - source of OS VHD to be used
SourceDATA - source of DATA VHD to be used
VHDPath - path to store VM VHD -- ex. (D:\Hyper-v\Hard Disks)
TargetOS - OS VHDX file name   ex. 'OS.VHDX'
TargetData - Data VHDX file name ex. 'DATA.VHDX'

       (settings within the VM)
Memory- ex. 8GB
SwitchName - 
Generation
ProcessorCount
VLAN

Script uses an invoke-command for remote connection to the host, it will then send commands to create the VM and also create 
a folder on the host 'c:\scripts\serversetupscripts' that will have all the powershell commands needed to run on the host after
VM is turned on and OS setup process (Manual part)

Once the setup of VM is complete, you will need to first run the 'Run_First_Script_HyperV_GuestServices_CopySetupFiles.ps1'  -
This will copy the setup files to the VM (Other PS1 files created) as well as turn on guest services.

Once this completes, you will need to login to each server and run the PS1 saved to c:\powershell folder, this will change hostname, set IP and DNS, change NIC name,
apply to domain controller, then reboot.

.NOTES
Name: Mass_Create_Hyper-V VMs.ps1
Version: 1.0
Author: Brasiledo
Date of last revision: 1/5/2022

#>

#Deletes Current CSV File 

    Remove-Item ".\Hyper-V_Setup_Details.csv"
     

 #Convert Matster Excel File to Master CSV File
 
    $a=gci ".\Hyper-V_Setup_Details.xlsx"
    $xlsx=new-object -comobject excel.application
    $xlsx.DisplayAlerts = $False

    foreach($aa in $a){
    $csv=$xlsx.workbooks.open($aa.fullname)
    $csv.sheets(1).saveas($aa.fullname.substring(0,$aa.fullname.length -4) + 'csv',6)
    }

    $xlsx.quit()
    $xlsx=$null

    [GC]::Collect()

#End Excel to CSV Conversion#

#Set Inital Variables

    $MasterFile = ".\Hyper-V_Setup_Details.csv"
    $DateStamp = get-date -uformat "%Y-%m-%d--%H-%M-%S" # Get the date
    #$cred=(get-credential)

#Create VM's on host

   Import-Csv -Path $MasterFile  -Delimiter ',' | Where-Object { $_.PSObject.Properties.Value -ne '' } | foreach {
Param (
    $VMNameHyperV = $($_.VMNameHyperV),
    $Host = $($_.Host),
    $SwitchName = $($_.SwitchName),
    $Memory =$($_.Memory),
    $Generation = $($_.Generation),
    $ProcessorCount = $($_.ProcessorCount),
    $VLAN = $($_.VLAN),
    $VHDPATH = $($_.VHDPATH),
    $TargetOS = $($_.TargetOS),
    $TargetData = $($_.TargetData),
    $SourceOS = $($_.SourceOS),
    $SourceData = $($_.SourceData)
    
    )
  invoke-command -ComputerName "$Host" -credential $cred -ScriptBlock {param($VMNameHyperV,$memory,$ProcessorCount,$HostIP,$Generation,$VLAN,$TargetOS,$TargetData,$VHDPATH,$SourceOS,$SourceData)
     
  #Copy VHD for OS and DATA drives to set location, ONLY IF ADDED on Spreadsheet
  
   IF ($SourceOS -ne '') {
New-Item -Path $VHDPATH -Name $VMNameHyperV -ItemType directory
Copy-Item -Path $SourceData -Destination "$VHDPATH\$VMNameHyperV\targetOS"}
   else {write-host "SourceOS is empty, OS VHD NOT Copied!"}
   IF ($SourceDATA -ne ''){Copy-Item -Path $SourceData -Destination "$VHDPATH\$VMNameHyperV\$SourceDATA"}
   Else {write-host "Sourcedata is empty, DATA VHD NOT Copied!"}
   start-sleep -s 4
       } -ArgumentList $TargetOS,$TargetData,$VHDPATH,$SourceOS,$SourceData,$VMNameHyperV
       }
    
#Create New VMs

  New-VM -Name "$VMNameHyperV" -MemoryStartupBytes (Invoke-Expression $memory) -Generation "$Generation" | out-host
    Set-VMProcessor -VMName "$VMNameHyperV" -Count $ProcessorCount
    
#Add SwitchName, OS VHD Drive IF Present on the Spreadsheet

    IF ($TargetOS -and $SwitchName -ne ""){
    Set-VM -SwitchName "$SwitchName" -VHDPath "$VHDPATH\$VMNameHyperV\$SourceOS"}
    else {write-host 'VHD is empty, SwitchName not set'}

#Set Networking/VLAN and Add HardDrive IF Present on the Spreadsheet

    IF ($VLAN -ne ""){
    Set-VMNetworkAdapterVlan -VMName $VMNameHyperV -Access -VlanId "$VLAN"}
    else {Write-host 'VLAN is Empty, VLAN not set'}
    IF ($TargetData -ne "") {
    
#Set Data VHD IF Present on the Spreadsheet 
 
    Get-VM $VMNameHyperV | Add-VMHardDiskDrive -ControllerType SCSI -ControllerNumber 0 -Path "$TargetData,$VHDPATH,$SourceOS,$SourceData"}
     else {write-host 'TargetData is Empty, HardDrive not added'}

  }  -ArgumentList $VMNameHyperV,$memory,$ProcessorCount,$HostIP,$Generation,$VLAN,$TargetOS,$TargetData,$VHDPATH,$SourceOS,$SourceData
  }
  
#End Create VM's on Host Powershell Scripts#
 
#Add setup script path

    invoke-command -ComputerName "$HOST" -credential $cred -ScriptBlock {
    if (test-path "C:\scripts\ServerSetupScripts"){
    Remove-Item "C:\scripts\ServerSetupScripts" -Force -Recurse}
    start-sleep -Seconds 1
    New-Item "C:\scripts\ServerSetupScripts" -ItemType directory | out-host}
   
  Pause

<#
Below Code creates and copies VM setup scripts to Host.  These scripts are to be run on the host after VM is turned on and OS setup manually.
#>

#Create .PS1 setup script for each VM. (add to domain, sets IP and DNS, changes NIC name, Restarts VM)

    Import-Csv -Path $MasterFile -Delimiter ',' | Where-Object { $_.PSObject.Properties.Value -ne '' } | foreach {
    param (
    $CurrentNetworkAdapterName = $($_.CurrentNetworkAdapterName),
    $NewNetworkAdapterName = $($_.NewNetworkAdapterName),
    $ServerName = $($_.ServerName),
    $IPAddress = $($_.IPAddress),
    $Subnet = $($_.Subnet),
    $GatewayAddress = $($_.GatewayAddress),
    $DNS = $($_.DNS),
    $WINS = $($_.WINS),
    $Domain = $($_.Domain),
    $user = $($_.user),
    $Host = $($_.Host),
    $VMNameHyperV = $($_.VMNameHyperV),
    $outputfile = "C:\scripts\ServerSetupScripts\$ServerName.ps1"
    )
invoke-command -ComputerName "$Host" -credential $cred -ScriptBlock {param($outputfile,$wins,$VMNameHyperV,$CurrentNetworkAdapterName,$NewNetworkAdapterName,$GatewayAddres,$IPAddress,$Subnet,$ServerName,$Domain)

if($WINS -eq "" -or $WINS -eq $null)
    { 
    "Set-ExecutionPolicy Bypass" | Add-Content $OutputFile
    ""| Add-Content $OutputFile
    "#Rename Adapter" | Add-Content $OutputFile
    "Rename-NetAdapter -Name '$CurrentNetworkAdapterName' -NewName '$NewNetworkAdapterName'"| Add-Content $OutputFile
    ""| Add-Content $OutputFile
    "#Add Static IP Address, DNS & WINS" | Add-Content $OutputFile        
    "netsh interface ip set address '$NewNetworkAdapterName' static $IPAddress $Subnet $GatewayAddress" | Add-Content $OutputFile
    "Set-DnsClientServerAddress -InterfaceAlias '$NewNetworkAdapterName' -ServerAddresses $DNS" | Add-Content $OutputFile
    ""| Add-Content $OutputFile
    "#Rename VM and Join to Domain" | Add-Content $OutputFile
    "Add-Computer -DomainName $Domain -Credential (Get-Credential $User) -NewName '$ServerName' -Restart"  | Add-Content $OutputFile
    }
     Else {
    "Set-ExecutionPolicy Bypass" | Add-Content $OutputFile
    ""| Add-Content $OutputFile
    "#Rename Adapter" | Add-Content $OutputFile
    "Rename-NetAdapter -Name '$CurrentNetworkAdapterName' -NewName '$NewNetworkAdapterName'"| Add-Content $OutputFile
    ""| Add-Content $OutputFile
    "#Add Static IP Address, DNS & WINS" | Add-Content $OutputFile        
    "netsh interface ip set address '$NewNetworkAdapterName' static $IPAddress $Subnet $GatewayAddress" | Add-Content $OutputFile
    "Set-DnsClientServerAddress -InterfaceAlias '$NewNetworkAdapterName' -ServerAddresses $DNS" | Add-Content $OutputFile
    "netsh interface ip set wins '$NewNetworkAdapterName' static $WINS" | Add-Content $OutputFile
    ""| Add-Content $OutputFile
    "#Rename VM and Join to Domain" | Add-Content $OutputFile
    "Add-Computer -DomainName $Domain -Credential (Get-Credential $User) -NewName '$ServerName' -Restart"  | Add-Content $OutputFile 
        }  
      }-ArgumentList $outputfile,$wins,$VMNameHyperV,$CurrentNetworkAdapterName,$NewNetworkAdapterName,$GatewayAddres,$IPAddress,$Subnet,$ServerName,$Domain
    }
     
   invoke-command -ComputerName "$Host" -credential $cred -ScriptBlock {write-host '';gci "C:\scripts\ServerSetupScripts\*.ps1"}
   

#End Create VM Setup Scripts#


#Create .PS1 scripts that copies previous scripts to each VM (Enables Hyper-V Guest Services, and copy VM setup script to each VM; To be run First)

 Import-Csv -Path $MasterFile -Delimiter ',' | Where-Object { $_.PSObject.Properties.Value -ne '' } | foreach {

 param (
    $VMNameHyperV = $($_.VMNameHyperV),
    $ServerName = $($_.ServerName),
    $ScriptOutFile = "C:\scripts\ServerSetupScripts\Run_First_Script_HyperV_GuestServices_CopySetupFiles.ps1"
    )

 invoke-command -ComputerName "$Host" -credential $cred -ScriptBlock {param($ScriptOutFile,$DateStamp,$VMNameHyperV,$ServerName )
    
    $DateStamp = get-date -uformat "%Y-%m-%d--%H-%M-%S" # Get the date
    "Powershell Scripts to Enable/Disable Guest Services and copy Setupfiles to VM's - $DateStamp"| Add-Content $ScriptOutFile
    ""| Add-Content $ScriptOutFile
    "****Enable Guest Service Scripts****" | Add-Content $ScriptOutFile
    ""| Add-Content $ScriptOutFile 
    "# HyperV Host $HostIP" | Add-Content $ScriptOutFile 
    "Enable-VMIntegrationService -VMName ""$VMNameHyperV""  -Name ""Guest Service Interface""" | Add-Content $ScriptOutFile
    "Copy-VMFile ""$VMNameHyperV"" -SourcePath ""C:\Powershell\$ServerName.ps1"" -DestinationPath ""C:\Powershell\$ServerName.ps1"" -CreateFullPath -FileSource Host"| Add-Content $ScriptOutFile
    "Disable-VMIntegrationService -VMName ""$VMNameHyperV""  -Name ""Guest Service Interface""" | Add-Content $ScriptOutFile
    ""| Add-Content $ScriptOutFile  
   }-argumentlist $ScriptOutFile,$DateStamp,$VMNameHyperV,$ServerName 
   
    }
   invoke-command -ComputerName "$host" -credential $cred -ScriptBlock {write-host '';gci "C:\scripts\ServerSetupScripts\run_first*.ps1"
   write-host ' '
   read-host '                               End of script.  Press Enter to exit.'
   write-host''}
   
#End Of Script#
