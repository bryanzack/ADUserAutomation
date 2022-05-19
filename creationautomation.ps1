#### PowerShell script to take an Excel.xlsx file, parse it for desired information, and use that information to 
#### automatically create new user accounts in Active Directory

## Desired information:
##
##  - First name      : pulled from "Name" header cell in excel file
##  - Last name       : pulled from "Name" header cell in excel file
##  - Username        : $lastName appended to the first letter of $firstName (Currently this is hard-coded for VicFin naming convention)
##  - Job title       : pulled from "Job Title" header cell in excel file
##  - Department      : pulled from "Department" header cell in excel file
##  - Manager         : **In order to automate this part of the account, an Active Directory account needs to be matched from the contents of "Manager" header cell in the excel file
##  - Password        : passord is input for each user
##  - Office Location : this value is taken and used to search the active directory ou exactly where each user should be to determine if it exists, or not

## TODO:
##
##  - What are the min # of permissions that the user deserves based on job title?
##  - leave manager field blank if none is given, if it is given, assign manager to user
##  - be able to handle username concatenation when given a Name monday.com header value that is more than 2 words/has unexpected characters
##  - if user does exist, make sure all attributes match that of excel, if they do not, update them

## office locations implemented so far:
##
##  - "Boyce HQ"
##  - "REMOTE"
##  - "Everett, WA"

# Selects excel file via file browser
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
[void]$FileBrowser.ShowDialog()
$ExcelFile = $FileBrowser.FileName

# Imports necessary modules / creates Excel object to be iterated through. 
# Comment out the Install-Module statement if running offline
#Install-Module -Name PSExcel
Import-Module PSExcel
Import-Module ActiveDirectory
$objExcel = New-Excel -Path $ExcelFile
$WorkBook = $objExcel | Get-Workbook

$domain = "zacklabs"
$domainExt = "com"

# Iterate through each worksheet (only 1 is included in the downloadable Excel file from monday.com)
ForEach($WorkSheet in @($Workbook.Worksheets)) {

	$totalNoOfRecords = $WorkSheet.Dimension.Rows
    $totalNoOfColumns = $WorkSheet.Dimension.Columns
   
    # for every record, iterate through all columns and pull desired information, then add the row information as a new AD user
    for ($i=4; $i -lt $totalNoOfRecords; $i++) {
        for ($j=1; $j -lt $totalNoOfColumns; $j++) {
            # if header cell contains Name
            if ($WorkSheet.Cells.Item(3,$j).text -eq "Name") {
                $name = $WorkSheet.Cells.Item($i,$j).text
                if ($name -match '!|@|#|\$') {
                    Write-Warning "Name '$name' contains illegal characters"
                    $name = Read-Host "New name for '$name'"
                    }
                foreach($_ in $name) {
		        	$nameOut = $_.split()
		        }
		        $firstName = $nameOut[0]
		        $firstChar = $firstName.substring(0,1)
		        $lastName = $nameOut[1]
		        
                }
            # if header cell contains Job Title
            elseif ($WorkSheet.Cells.Item(3,$j).text -eq "Job Title") {
                $jobTitle = $WorkSheet.Cells.Item($i,$j).text
                if ($jobTitle -eq $null) {
                    $jobTitle = "N/A"
                    }
                }
            # if header cell contains Manager
            elseif ($WorkSheet.Cells.Item(3,$j).text -eq "Manager") {
                $manager = $WorkSheet.Cells.Item($i,$j).text
                if ($manager -eq $null) {
                    $hasManager = $false
                    }
                }
            # if header cell contains Department
            elseif ($WorkSheet.Cells.Item(3,$j).text -eq "Department") {
                $department = $WorkSheet.Cells.Item($i,$j).text
                if ($departmnet -eq $null) {
                    $department = "N/A"
                    }
                }
            # if header cell contains Office Location / also handles OU path locations and their respective naming conventions
            elseif ($WorkSheet.Cells.Item(3, $j).text -eq "Office Location") {
                $officeLocation = $WorkSheet.Cells.Item($i,$j).text
                if ($officeLocation -eq "Everett, WA") {
                    $userName = "$firstName".ToLower()
                    $upnSuffix = "@dwellmtg.com"
                    $streetAddress = "2707 Colby Ave, Ste 1212"
                    $company = "DwellMTG"
                    $city = "Everett"
                    $state = "WA"
                    $zipCode = "98201"
                    $emailAddress = "$userName$upnSuffix"
                    $ouPath = "OU=Everett,OU=Washington,OU=DwellMtg,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                    }
                elseif ($officeLocation -eq "REMOTE") {
                    $userName = "$firstChar$lastName".ToLower()
                    $upnSuffix = "@victorianfinance.com"
                    $streetAddress = "2570 Boyce Plaza Rd"
                    $company = "Victorian Finance, LLC"
                    $city = "Pittsburgh"
                    $state = "PA"
                    $zipCode = "15241"
                    $emailAddress = "$userName$upnSuffix"
                    $ouPath = "OU=RemoteUsers,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                    }
                elseif ($officeLocation -eq "Boyce HQ") {
                    $userName = "$firstChar$lastName".ToLower()
                    $upnSuffix = "@victorianfinance.com"
                    $streetAddress = "2570 Boyce Plaza Rd"
                    $company = "Victorian Finance, LLC"
                    $city = "Pittsburgh"
                    $state = "PA"
                    $zipCode = "15241"
                    $emailAddress = "$userName$upnSuffix"
                    $ouPath = "OU=USC,OU=Pennsylvania,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                    }
                elseif ($officeLocation -eq "Lafayette, LA") {
                    $userName = $firstName.ToLower()
                    $upnSuffix = "@completemortgagela.com"
                    $streetAddress = "100 Asma Blvd, Suite 100"
                    $company = "The Complete Mortgage Team"
                    $city = "Lafayette"
                    $state = "La"
                    $zipCode = "70506"
                    $emailAddress = "$userName$upnSuffix"
                    $ouPath = "OU=Louisiana,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                    }             
                }
            }

            # check and see if the generated username already exists as a user in the respective OU in Active Directory
            if (Get-ADUser -SearchBase $ouPath -F { samaccountname -eq $userName} ) {
                Write-Warning "A user account with username $userName already exists in Active Directory. Checking for discrepancies..."
                hasDiscrepancy = $false
                if ((Get-ADUser -SearchBase $ouPath -F {street -ne $streetAddress}) -or 
                        (Get-ADUser -SearchBase $ouPath -F {company -ne $company}) -or
                        (Get-ADUser -SearchBase $ouPath -F {city -ne $city}) -or
                        (Get-ADUser -SearchBase $ouPath -F {state -ne $state}) -or
                        (Get-ADUser -SearchBase $ouPath -F {postalcode -ne $zipCode}) -or
                        (Get-ADUser -SearchBase $ouPath -F {emailaddress -ne $emailAddress})){

                    hasDiscrepancy = $true
                    }
                $answer = Read-Host "Discrepancy found, overwrite existing AD info with Excel data? (y/n)"
                if ($answer -eq "y") {
                    Set-ADUser `
                            -Identity $userName `
                            -Enabled $true `
                            -SamAccountName $userName `
                            -GivenName $firstName `
                            -Surname $lastName `
                            -Company $company `
                            -Street $streetAddress `
                            -City $city `
                            -State $state `
                            -postalCode $zipCode `
                            -Title $jobTitle `
                            -Department $department `
                            -displayName $displayName `
                            -userPrincipalName "$userName$upnSuffix" `
                            -EmailAddress $emailAddress `
                    Write-Host "Updated user $userName" -ForegroundColor Cyan
                    }
                else {
                continue
                    }


                }
            else {
                #$ouPath = "OU=USC,OU=Pennsylvania,OU=Users,OU=Accounts,DC=zacklabs,DC=com"
                #Write-Output $ouPath
                $upnSuffix = "@zacklabs.com"
                $meetsRequirements = $false        
                while (!$meetsRequirements) {
                    try {
                        $password = Read-Host "password for $name ($userName)"
                        New-ADUser `
                            -Enabled $true `
                            -Path $ouPath `
                            -Name $name `
                            -SamAccountName $userName `
                            -GivenName $firstName `
                            -Surname $lastName `
                            -Company $company `
                            -Street $streetAddress `
                            -City $city `
                            -State $state `
                            -postalCode $zipCode `
                            -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -ChangePasswordAtLogon $False `
                            -OtherAttributes @{'title'=$jobTitle; `
                                               'department'=$department; `
                                               'displayName'="$lastName, $firstName"; `
                                               'userPrincipalName'="$userName$upnSuffix"; `
                                               'mail'=$emailAddress} `

                           $meetsRequirements = $true
                           Write-Host "The user account $userName has been created." -ForegroundColor Cyan
                        }
                    # handles password complexity exception
                    catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException] {
                        #Write-Output $_ # prints out exact error message
                        Remove-ADUser -Identity $userName -Confirm:$false
                        Write-Warning "Password requirements not met"
                        $meetsRequirements = $false
                        }
                    # handles managernotfound exception
                    catch [Microsoft.ActiveDirectory.Management.ADIdentityResolutionException] {
                        #Write-Warning "Manager '$manager' was not found, ensure that the designated manager for '$name' is correct and try again. Previously added users will persist in AD." -ErrorAction Exit
                    }
                    # handles any other exception and writes it to host
                    catch {
                        Write-Output $_
                        }
                    }
                }
        }
    }
