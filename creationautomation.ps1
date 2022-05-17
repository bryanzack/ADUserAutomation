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
##  - Add company address given Monday.com information
##  - Create templates for different kinds of users in each department
##  - properly decide whhich OU path destination to create the user in given info from monday.com

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
#Install-Module -Name PSExcel
Import-Module PSExcel
Import-Module ActiveDirectory
$objExcel = New-Excel -Path $ExcelFile
$WorkBook = $objExcel | Get-Workbook

# Iterate through each worksheet (only 1 is included in the downloadable Excel file from monday.com)
ForEach($WorkSheet in @($Workbook.Worksheets)) {

	$totalNoOfRecords = $WorkSheet.Dimension.Rows
    $totalNoOfColumns = $WorkSheet.Dimension.Columns
    $ouPath = ""
   
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
            elseif ($WorkSheet.Cells.Item(3, $j).text -eq "Office Location") {
                $officeLocation = $WorkSheet.Cells.Item($i,$j).text
                if ($officeLocation -eq "Everett, WA") {
                    $userName = "$firstName".ToLower()
                    $ouPath = "OU=Everett,OU=Washington,OU=DwellMtg,OU=Users,OU=Accounts,DC=victorianfinance,DC=local"
                    }
                elseif ($officeLocation -eq "REMOTE") {
                    $userName = "$firstChar$lastName".ToLower()
                    $ouPath = "OU=RemoteUsers,OU=Users,OU=Accounts,DC=victorianfinance,DC=local"
                    }
                elseif ($officeLocation -eq "Boyce HQ") {
                    $userName = "$firstChar$lastName".ToLower()
                    $ouPath = "OU=USC,OU=Pennsylvania,OU=Users,OU=Accounts,DC=victorianfinance,DC=local"
                    }                
                }
            }

            # check and see if the generated username already exists as a user in Active Directory
            #Write-Output $ouPath
            if (Get-ADUser -SearchBase $ouPath -F { SAMAccountName -eq $userName} ) {
                Write-Warning "A user account with username $userName already exists in Active Directory path $oupath."
                }
            else {
                $meetsRequirements = $false        
                while (!$meetsRequirements) {
                    try {
                        $password = Read-Host "password for $name (${userName})"
                        New-ADUser `
                            -Path $ouPath `
                            -Name $name `
                            -GivenName $firstName `
                            -Surname $lastName `
                            -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -ChangePasswordAtLogon $False `
                            -OtherAttributes @{'title'=$jobTitle; `
                                               'department'=$department; `
                                               'displayName'="$lastName, $firstName"; `
                                               'userPrincipalName'=$userName; `
                                               'manager'=$manager} `

                           $meetsRequirements = $true
                           Write-Host "The user account $userName has been created." -ForegroundColor Cyan
                        }
                    # handles password complexity exception
                    catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException] {
                        #Write-Output $_ # prints out exact error message
                        Remove-ADUser -Identity $Name -Confirm:$false
                        Write-Warning "Password requirements not met"
                        $meetsRequirements = $false
                        }
                    # handles managernotfound exception
                    catch [Microsoft.ActiveDirectory.Management.ADIdentityResolutionException] {
                        Write-Warning "Manager '$manager' was not found, ensure that the designated manager for '$name' is correct and try again. Previously added users will persist in AD." -ErrorAction Continue
                    }
                    # handles any other exception and writes it to host
                    catch {
                        Write-Output $_
                        }
                    }
                }
        }
    }