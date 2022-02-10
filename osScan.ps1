
#STOP STOP STOP
#STOP STOP STOP
#STOP STOP STOP
#STOP STOP STOP
#STOP STOP STOP
#TO RUN THIS SCRIPT, launch RunOSscan.ps1 from an administrative powershell platform.

#Establish admin login credentials
$uname = Read-Host "Enter Domain\Username "
$pword = Read-Host "Enter Password " -AsSecureString
$Credentials = New-Object System.Management.Automation.PSCredential $uname,$pword

#Create function to pull reg info
function Get-osScan { 
param ($Computer) 
$productName = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion').GetValue('ProductName') 
$releaseID = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion').GetValue('ReleaseID')  
$computername = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName').GetValue('ComputerName')
$Organization= [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion').GetValue('RegisteredOrganization')
$OrgOwner= [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion').GetValue('RegisteredOwner')
$BuildNumber= [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion').GetValue('CurrentBuild')

#Write output to console
Write-Host "Organization: $Organization"
Write-Host "Owner: $OrgOwner"
Write-Host "Endpoint: $computername"
Write-Host "Product version: $productName" 
Write-Host "Engine version: $releaseID"
Write-Host "Build: $BuildNumber" 

#Excel file creation


	}	

#Remote execution command using admin credentials
ForEach($RemoteComputer in Get-Content "C:\Users\JU30017\OneDrive - MIT Lincoln Laboratory\HomeDrive\Working\list.txt") {

Invoke-Command -ComputerName $RemoteComputer -Credential $Credentials ${Function:Get-osScan}

}