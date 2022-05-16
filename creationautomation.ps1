#### PowerShell script to take an Excel.xlsx file, parse it for desired information, and use that information to 
#### automatically create new user accounts in Active Directory

## Desired information:
##
##  - First name : pulled from "Name" header cell in excel file
##  - Last name  : pulled from "Name" header cell in excel file
##  - Username   : $lastName appended to the first letter of $firstName (Currently this is hard-coded for VicFin naming convention)
##  - Job title  : pulled from "Job Title" header cell in excel file
##  - Department : pulled from "Department" header cell in excel file
##  - Manager    : **In order to automate this part of the account, an Active Directory account needs to be matched from the contents of "Manager" header cell in the excel file
##  - Password   : passord is input for each user
##  - Address    : 

## TODO:
##
#X  - What building will the user be located at? How does this affect naming convention/other AD attributes?
#X  - What are the min # of permissions that the user deserves based on job title?
#X  - Handle specific password requirement error message
#X  - Only add users to Active Directory if the script executes entirely and sucessfully.

## possible office locations:
##
##  - "Boyce HQ"
##  - "REMOTE"
##  - "Lafayette, LA"
##  - NULL

# Selects excel file via file browser
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
[void]$FileBrowser.ShowDialog()
$ExcelFile = $FileBrowser.FileName 

# Imports necessary modules / creates Excel object to be iterated through
Import-Module PSExcel
Import-Module ActiveDirectory
$objExcel = New-Excel -Path $ExcelFile
$WorkBook = $objExcel | Get-Workbook

# Iterate through each worksheet (only 1 is included in the downloadable Excel file from monday.com)
ForEach($WorkSheet in @($Workbook.Worksheets)) {

	$totalNoOfRecords = $WorkSheet.Dimension.Rows
    $totalNoOfColumns = $WorkSheet.Dimension.Columns
   
    # for every record, iterate through all columns and pull desired information, then add the row information as a new AD user
    for ($i=4; $i -lt $totalNoOfRecords; $i++) {
        for ($j=1; $j -lt $totalNoOfColumns; $j++) {
            if ($WorkSheet.Cells.Item(3,$j).text -eq "Name") {
                $name = $WorkSheet.Cells.Item($i,$j).text
                foreach($_ in $name) {
		        	$nameOut = $_.split()
		        }
		        $firstName = $nameOut[0]
		        $firstChar = $firstName.substring(0,1)
		        $lastName = $nameOut[1]
		        $userName = "$firstChar$lastName".ToLower()
                }
            elseif ($WorkSheet.Cells.Item(3,$j).text -eq "Job Title") {
                $jobTitle = $WorkSheet.Cells.Item($i,$j).text
                }
            elseif ($WorkSheet.Cells.Item(3,$j).text -eq "Manager") {
                $manager = $WorkSheet.Cells.Item($i,$j).text
                }
            elseif ($WorkSheet.Cells.Item(3,$j).text -eq "Department") {
                $department = $WorkSheet.Cells.Item($i,$j).text
                }
            }

            # check and see if the generated username already exists as a user in Active Directory
            if (Get-ADUser -F { SamAccountName -eq $userName}) {
                Write-Warning "A user account with username $userName already exists in Active Directory."
                }
            else {
                $meetsRequirements = $false        
                while (!$meetsRequirements) {
                    try {
                        Write-Output "jobtitle: $jobTitle"
                        $password = Read-Host "password for $name (${userName})"
                        New-ADUser `
                            -Name $name `
                            -GivenName $firstName `
                            -Surname $lastName `
                            -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -ChangePasswordAtLogon $False
                            -OtherAttributes @{'title'=$jobTitle; `
                                               'department'=$department; `
                                               'displayName'="$lastName, $firstName";}

                           $meetsRequirements = $true
                           Write-Host "The user account $userName has been created." -ForegroundColor Cyan
                        }
                    catch {
                        Remove-ADUser -Identity $userName -Confirm:$false
                        Write-Warning "Password requirements not met"
                        $meetsRequirements = $false
                
                        }
                    }
                }
        }
    }