#
#   Name: appScan1.2.ps1
#   Date Created: 20201123
#   Description: Scan local system to return intalled software list
#
#
#
#   Justin Ross -   20201123          Created
#                                     List confined to single page
#                                     Filter drop down added to header
#                                     Variable needs to be reset|removed if the script does not fully complete to return scan
#                                     Freeze top row freeze
#
#   Justin Ross -   20201124          Header Modification
#                                     Clear variables at the end
#                                     Save Path added and file Name
#                                     Worksheet formating
#                                     OutPut Client names
#                                     Error handling for unreachable clients
#
#   Justin Ross -   20201130          Error handling added for "Access Denied" by Firewalls or WMI permissions
#                                     File save accounts for the instance of an existing file with the same name
#                                     Error handling for earsing null variables at the end of the script
#
#   Justin Ross -   20210920          File system change, oneDirve implementation
#
#   Justin Ross -   20210927          Second Worksheet added to contain osScan.ps1 information
#                                     With the addition of osScan using "Microsoft.Win32.RegistryKey", administative powershell environment is required
#
#
#  :Note:  If the script is interupted and does not fully complete, the variables will not be reset.  
#          An immediate execution following this result will have unpredictable results.  
#          The following will clear all variables for the next execution.
#              Remove-Variable *  -ErrorAction SilentlyContinue
#          
#
####################################################################################################################################


$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $True

#Credentials for osScan
$uname = Read-Host "Enter Domain\Username "
$pword = Read-Host "Enter Password " -AsSecureString
$Credentials = New-Object System.Management.Automation.PSCredential $uname,$pword

#$Excel.SheetsInNewWorkbook = @(get-content "S:\Working\scripts\ps\Servers.txt").count #This line can be used to create multiple sheets within the workbook

#Counter variable for rows
#Sheet2 row counters
      $intRow2 = 2
      $colorRow = 2

#osScan Function

function Get-osScan { 
param ($Computer) 
$productName = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion').GetValue('ProductName') 
$releaseID = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion').GetValue('ReleaseID')  
$computername = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName').GetValue('ComputerName')
$Organization= [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion').GetValue('RegisteredOrganization')
$OrgOwner= [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion').GetValue('RegisteredOwner')
$BuildNumber= [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion').GetValue('CurrentBuild')

Write-Host "Organization: $Organization"
Write-Host "Owner: $OrgOwner"
Write-Host "Endpoint: $computername"
Write-Host "Product version: $productName" 
Write-Host "Engine version: $releaseID"
Write-Host "Build: $BuildNumber" 

    $ray = $computername,$Organization,$OrgOwner,$productName,$releaseID,$BuildNumber
    return $ray
}

#Excel file creation	


#Read thru the contents of the Servers.txt file
foreach ($xPoints in get-content "C:\Users\JU30017\OneDrive - MIT Lincoln Laboratory\HomeDrive\Working\scripts\ps\Servers.txt")
{
 #   if(
    #$Excel = $Excel.Workbooks.Add()
    #$Sheet = $Excel.Worksheets.Item($i) #This line can be modified to create multiple sheets within the workbook
                                        #$Sheet.Name = $xPoints  #Workbook Sheet Names
    
    if (!$intRow){
        $intRow = 1
        $Excel = $Excel.Workbooks.Add()
        $Sheet2 = $Excel.Worksheets.Item(1) #This line can be modified to create multiple sheets within the workbook
        $Sheet2.Name = "Operating System"
        $Sheet = $Excel.Worksheets.item(2)  #Tring to create a second worksheet for osScan.ps1
        $Sheet.Name = "Applications"
                                            #$Sheet.Name = $xPoints  #Workbook Sheet Names
       
        }
           
#echo 1
#pause

    # Send a ping to verify if the Server is online or not.
    $ping = Get-WmiObject -ErrorAction SilentlyContinue `
    -query "SELECT * FROM Win32_PingStatus WHERE Address = '$xPoints'"
       if ($Ping.StatusCode -eq 0) {

         #Create column headers
         if ($intRow -eq 1){  
                
                #Application Worksheet                 
                $Sheet.Cells.Item($intRow,1) = "Endpoint: "
                $Sheet.Cells.Item($intRow,1).Font.Bold = $True
                $Sheet.Cells.Item($intRow,2).Font.Bold = $True
                $Sheet.Cells.Item($intRow,3).Font.Bold = $True

                $Sheet2.Cells.Item($intRow,1) = "Endpoint: "
                $Sheet2.Cells.Item($intRow,1).Font.Bold = $True
                $Sheet2.Cells.Item($intRow,2).Font.Bold = $True
                $Sheet2.Cells.Item($intRow,3).Font.Bold = $True
#echo 2
           
                if ($intRow -ge 2){   
                    $Sheet.Cells.Item($intRow,1) = $xPoints
                    }
                $Sheet.Cells.Item($intRow,2) = "Aapplication"
                $Sheet.Cells.Item($intRow,3) = "Version"
                $headerRange = $Sheet.Range("a1","c1")
                $headerRange.AutoFilter() | Out-Null

                $Sheet2.Cells.Item($intRow,2) = "Organization"
                $Sheet2.Cells.Item($intRow,3) = "Owner"
                $Sheet2.Cells.Item($intRow,4) = "OS Version"
                $Sheet2.Cells.Item($intRow,5) = "Release ID"
                $Sheet2.Cells.Item($intRow,6) = "Build ID"
                $headerRange2 = $Sheet2.Range("a1","f1")
                $headerRange2.AutoFilter() | Out-Null
                
                #Freeze top two rows
                $Sheet.Application.ActiveWindow.SplitRow = 1
                $Sheet.Application.ActiveWindow.FreezePanes = $True

                $Sheet2.Application.ActiveWindow.SplitRow = 1
                $Sheet2.Application.ActiveWindow.FreezePanes = $True

                #Format the column headers sheet 1
                for ($col = 1; $col –le 3; $col++){
                
                    $Sheet.Cells.Item($intRow,$col).Font.Bold = $True
                    $Sheet.Cells.Item($intRow,$col).Font.Underline = $True
                    $Sheet.Cells.Item($intRow,$col).Interior.ColorIndex = 48
                    $Sheet.Cells.Item($intRow,$col).Font.ColorIndex = 25
                }
                 #Format the column headers sheet 2
                for ($col = 1; $col –le 6; $col++){
                    $Sheet2.Cells.Item($intRow,$col).Font.Bold = $True
                    $Sheet2.Cells.Item($intRow,$col).Font.Underline = $True
                    $Sheet2.Cells.Item($intRow,$col).Interior.ColorIndex = 48
                    $Sheet2.Cells.Item($intRow,$col).Font.ColorIndex = 25
                }
             }
            
             $intRow++
#echo 3      #Creates a list of applications and stores it in software var
                $software = Get-WmiObject -ErrorAction SilentlyContinue `
                -ComputerName $xPoints -Class Win32_Product | Sort-Object Name
              
             #Failed Get-WmiObject error handling
                if (!$software){
                Write-Host "Access Denied. Check WMI permssions and firewall settings:"} 

             #Formatting using Excel 
#echo
             foreach ($objItem in $software){
                $Sheet.Cells.Item($intRow, 2) = $objItem.Name
                $Sheet.Cells.Item($intRow, 3) = $objItem.Version
                $Sheet.Cells.Item($intRow, 1) = $xPoints
             
             #Adding alternating colors for each row
                for ($col = 1; $col –le 3; $col++){

                    if($intRow % 2 -eq 0){                    
                        
                        $Sheet.Cells.Item($intRow,$col).Interior.ColorIndex = 34
                    }
                    else{
                        $Sheet.Cells.Item($intRow,$col).Interior.ColorIndex = 15
                    }
                }
                $intRow ++                 
             }            
        $intRow--
       
      #osScan remote endpoint  
      #Export data to excel spreadsheet   
         $osArray = Invoke-Command -ComputerName $xPoints -Credential $Credentials ${Function:Get-osScan}
         $Sheet2.Cells.Item($intRow2, 1) = $osArray[0]
         $Sheet2.Cells.Item($intRow2, 2) = $osArray[1]
         $Sheet2.Cells.Item($intRow2, 3) = $osArray[2]
         $Sheet2.Cells.Item($intRow2, 4) = $osArray[3]
         $Sheet2.Cells.Item($intRow2, 5) = $osArray[4]
         $Sheet2.Cells.Item($intRow2, 6) = $osArray[5]
           
      $intRow2++

                 #Adding alternating colors for each row
             
              for ($col = 1; $col –le 6; $col++){
              
                  if($colorRow % 2 -eq 0){                    
                        
                      $Sheet2.Cells.Item($colorRow,$col).Interior.ColorIndex = 34
                  }
                  else{
                      $Sheet2.Cells.Item($colorRow,$col).Interior.ColorIndex = 15
                  }
              }
              $colorRow ++ 
             
 ##Add if statement here to segment out additional output that needs to be parced

        if (!$software){
        Write-Host "$xPoints _ Applications Not Scanned"}
        else{
        echo "$xPoints _ Applications Scanned"}
        $Sheet.UsedRange.EntireColumn.AutoFit() | Out-Null
        $Sheet2.UsedRange.EntireColumn.AutoFit() | Out-Null
    }
    else{
        echo "$xPoints : Verifiy Connectivity. Scan Failed"
    }
 }
#Save to folder
$fSalt = Get-Random -Maximum 100
$savePath = "C:\Temp\$(get-date -f yyyy-MM-dd)_AppScan_$fSalt.xlsx"

#Loop checks for existing file and renames .xlsx file if file exists
while (Test-Path -Path $savePath){
    $fSalt++
    $savePath = "C:\Temp\$(get-date -f yyyy-MM-dd)_AppScan_$fSalt.xlsx"
    } 
#Saves .xlsx file
$Sheet.SaveAS($savePath)   
#Clears all used variables
Write-Output "Scan Completed, see results in $savePath !"     
#Clears all variables used  in ISE session
Remove-Variable * -ErrorAction SilentlyContinue
