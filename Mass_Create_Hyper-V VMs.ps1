  <# 
.SYNOPSIS
Bulk Hyper-V creation tool.  
 
.DESCRIPTION
-Fill out and save excel sheet (Hyper-V_Setup_Details.xlsx) with headers -- 
Host,SourceOS,SourceData,VMNameHyperV,SwitchName,Memory,Generation,ProcessorCount,VLAN,VHDPath,TargetOS,TargetData,
ServerName,CurrentNetworkAdapterName,NewNetworkAdapterName,IPAddress,Subnet,GatewayAddress,DNS,WINS,Domain,User

-For Host, use the computername
-Store the script along with the excel file in the same folder, once excel sheet filled out, run script from local machine

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
  invoke-command -ComputerName "Host" -credential $cred -ScriptBlock {param($VMNameHyperV,$memory,$ProcessorCount,$HostIP,$Generation,$VLAN,$TargetOS,$TargetData,$VHDPATH,$SourceOS,$SourceData)
     #copy VHD for OS and DATA drives to set location
     Copy-Item -Path $SourceOS -Destination $VHDPATH\$TargetOS
     Copy-Item -Path $SourceData -Destination $VHDPATH\$TargetData
    
    #Create New VMs
    New-VM -Name "$VMNameHyperV" -MemoryStartupBytes (Invoke-Expression $memory) -Generation "$Generation" -SwitchName "$SwitchName" -VHDPath "$VHDPATH\$TargetOS" | out-host 
    Set-VMProcessor -VMName "$VMNameHyperV" -Count $ProcessorCount 
    Set-VMNetworkAdapterVlan -VMName $VMNameHyperV -Access -VlanId "$VLAN"
    Get-VM $VMNameHyperV | Add-VMHardDiskDrive -ControllerType SCSI -ControllerNumber 0 -Path $VHDPATH\$TargetData
  }  -ArgumentList $VMNameHyperV,$memory,$ProcessorCount,$HostIP,$Generation,$VLAN,$TargetOS,$TargetData,$VHDPATH,$SourceOS,$SourceData
  }
  pause
  
 #End Create VM's on Host Powershell Scripts#
  
 #setup scripts
   invoke-command -ComputerName "HOST" -credential $cred -ScriptBlock {
   if (test-path "C:\scripts\ServerSetupScripts"){
    Remove-Item "C:\scripts\ServerSetupScripts" -Force -Recurse}
    start-sleep -Seconds 1
    New-Item "C:\scripts\ServerSetupScripts" -ItemType directory | out-host}
   
Pause

#copy to PS1 file serverscripts, to run direct on the VM

   Import-Csv -Path $MasterFile -Delimiter ',' | Where-Object { $_.PSObject.Properties.Value -ne '' } | foreach {
    $CurrentNetworkAdapterName = $($_.CurrentNetworkAdapterName)
    $NewNetworkAdapterName = $($_.NewNetworkAdapterName)
    $ServerName = $($_.ServerName)
    $IPAddress = $($_.IPAddress)
    $Subnet = $($_.Subnet)
    $GatewayAddress = $($_.GatewayAddress)
    $DNS = $($_.DNS)
    $WINS = $($_.WINS)
    $Domain = $($_.Domain)
    $user = $($_.user)
    $Host = $($_.Host)


invoke-command -ComputerName "HOST" -credential $cred -ScriptBlock {param($outputfile,$wins,$CurrentNetworkAdapterName,$NewNetworkAdapterName,$GatewayAddres,$IPAddress,$Subnet,$ServerName,$Domain)
 $outputfile = "C:\scripts\ServerSetupScripts\$ServerName.ps1"
if($WINS -eq "" -or $WINS -eq $null)
    {
    
    "Set-ExecutionPolicy Bypass" | Add-Content $OutputFile
    ""| Add-Content $OutputFile
    "#Rename Adapter" | Add-Content $OutputFile
    "Rename-NetAdapter -Name ""$CurrentNetworkAdapterName"" -NewName ""$NewNetworkAdapterName"""| Add-Content $OutputFile
    ""| Add-Content $OutputFile
    "#Add Static IP Address, DNS & WINS" | Add-Content $OutputFile        
    "netsh interface ip set address ""$NewNetworkAdapterName"" static $IPAddress $Subnet $GatewayAddress" | Add-Content $OutputFile
    "Set-DnsClientServerAddress -InterfaceAlias “"$NewNetworkAdapterName"” -ServerAddresses $DNS" | Add-Content $OutputFile
    ""| Add-Content $OutputFile
    "#Rename VM and Join to Domain" | Add-Content $OutputFile
    "Add-Computer -DomainName $Domain -Credential (Get-Credential $User) -NewName ""$ServerName"" -Restart"  | Add-Content $OutputFile
    }
     Else {
    "Set-ExecutionPolicy Bypass" | Add-Content $OutputFile
    ""| Add-Content $OutputFile
    "#Rename Adapter" | Add-Content $OutputFile
    "Rename-NetAdapter -Name ""$CurrentNetworkAdapterName"" -NewName ""$NewNetworkAdapterName"""| Add-Content $OutputFile
    ""| Add-Content $OutputFile
    "#Add Static IP Address, DNS & WINS" | Add-Content $OutputFile        
    "netsh interface ip set address ""$NewNetworkAdapterName"" static $IPAddress $Subnet $GatewayAddress" | Add-Content $OutputFile
    "Set-DnsClientServerAddress -InterfaceAlias “"$NewNetworkAdapterName"” -ServerAddresses $DNS" | Add-Content $OutputFile
    "netsh interface ip set wins ""$NewNetworkAdapterName"" static $WINS" | Add-Content $OutputFile
    ""| Add-Content $OutputFile
    "#Rename VM and Join to Domain" | Add-Content $OutputFile
    "Add-Computer -DomainName $Domain -Credential (Get-Credential $User) -NewName ""$ServerName"" -Restart"  | Add-Content $OutputFile
   
    }
       
     }-ArgumentList $outputfile,$wins,$CurrentNetworkAdapterName,$NewNetworkAdapterName,$GatewayAddres,$IPAddress,$Subnet,$ServerName,$Domain
     
     }
   invoke-command -ComputerName "teknetdc01" -credential $cred -ScriptBlock {write-host '';gci "C:\scripts\ServerSetupScripts"}
   pause
##End Create VM Setup Scripts and other HyperV Host Powershell Scripts


#Start Create Powershell scripts that copy Setupfiles to VM's on HyperV Host


 Import-Csv -Path $MasterFile -Delimiter ',' | Where-Object { $_.PSObject.Properties.Value -ne '' } | foreach {
    $VMNameHyperV = $($_.VMNameHyperV)
    $ServerName = $($_.ServerName)
    
 invoke-command -ComputerName "HOST" -credential $cred -ScriptBlock {param($ScriptOutFile,$DateStamp,$VMNameHyperV,$ServerName )
   $ScriptOutFile = "C:\scripts\ServerSetupScripts\Run_First_Script_HyperV_GuestServices_CopySetupFiles.ps1"
    $DateStamp = get-date -uformat "%Y-%m-%d--%H-%M-%S" # Get the date
    "Powershell Scripts to Enable/Disable Guest Services and copy Setupfiles to VM's - $DateStamp"| Add-Content $ScriptOutFile
    ""| Add-Content $ScriptOutFile
    "****Enable Guest Service Scripts****" | Add-Content $ScriptOutFile
    ""| Add-Content $ScriptOutFile

  
    "# HyperV Host "HOST" | Add-Content $ScriptOutFile 
    "Enable-VMIntegrationService -VMName ""$VMNameHyperV""  -Name ""Guest Service Interface""" | Add-Content $ScriptOutFile
    "Copy-VMFile ""$VMNameHyperV"" -SourcePath ""C:\Powershell\$ServerName.ps1"" -DestinationPath ""C:\Powershell\$ServerName.ps1"" -CreateFullPath -FileSource Host"| Add-Content $ScriptOutFile
    "Disable-VMIntegrationService -VMName ""$VMNameHyperV""  -Name ""Guest Service Interface""" | Add-Content $ScriptOutFile
    ""| Add-Content $ScriptOutFile
   
   
   }-argumentlist $ScriptOutFile,$DateStamp,$VMNameHyperV,$ServerName 
   
    }
    
   invoke-command -ComputerName "HOST" -credential $cred -ScriptBlock {write-host '';gci "C:\scripts\ServerSetupScripts\"
   write-host ' '
   read-host '                               End of Script.  Press Enter to Exit.'
   write-host''}
   
#End Create Powershell scripts that copy Setupfiles to VM's on HyperV Host  
