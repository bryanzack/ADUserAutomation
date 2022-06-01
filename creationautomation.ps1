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
##  - Branch          : the branch matters when the office location is 'Remote'. the branch value is used to determine what attribute information the remote user should have. 

## office locations implemented so far:
##
##  - "Boyce HQ"
##  - "REMOTE"
##  - "Everett, WA"


# Imports necessary modules / creates Excel object to be iterated through. 
# Comment out the Install-Module statement if running offline
# Install-Module -Name PSExcel

# Reads script arguments
param(
		[string]$p
     )

Import-Module PSExcel
Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

function Main-Function {
	if($p -eq $null) {
		Write-Warning "Invalid argument. Use 'add' or 'remove'."
	}
	else {
		if($p -eq 'add') {
			Add-Users	
		}
		elseif($p -eq 'remove') {
			$result = [System.Windows.MessageBox]::Show("WARNING: You are about to delete users from Active Directory", "Question", "YesNo", "Question")
			if ($result -eq 'Yes') {
				Remove-Users
				}
			else {
				Write-Output "You chose NO"
					exit
				}
			}
	else {
		Write-Warning "Invalid argument. Use 'add' or 'remove'."
		}
	}
}

function Remove-Users {
	$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
	[void]$FileBrowser.ShowDialog()
	$ExcelFile = $FileBrowser.FileName
	$objExcel = New-Excel -Path $ExcelFile
	$WorkBook = $objExcel | Get-Workbook
	$domain = "victorianfinance"
	$domainExt = "local"



	# Iterate through each worksheet (1 in this case)
	ForEach($WorkSheet in @($Workbook.Worksheets)) {
		$totalNoOfRecords = $WorkSheet.Dimension.Rows
		$totalNoOfColumns = $WorkSheet.Dimension.Columns

		# For every record, iterate through all columns and pull desired information
		for ($i=4; $i -lt $totalNoOfRecords; $i++) {
			for ($j=1; $j -lt $totalNoOfColumns; $j++) {
				if ($WorkSheet.Cells.Item(3,$j).text -eq "Name") {
					$name = $WorkSheet.Cells.Item($i,$j).text
					$nameOut = $name.split()
					$firstName = $nameOut[0]
                                        $firstChar = $firstName.substring(0,1)
					$lastName = $nameOut[1]

					[String[]]$items += $name
				}

                                elseif ($WorkSheet.Cells.Item(3,$j).text -eq "Branch") {
                                    $branch = $WorkSheet.Cells.Item($i,$j).text
                                }

                                elseif ($WorkSheet.Cells.Item(3,$j).text -eq "Office Location") {
                                    $officeLocation = $WorkSheet.Cells.item($i,$j).text
                                    if ($officeLocation -eq "") {
                                        $userName = "$lastName".ToLower()
                                        Write-Warning "No office location found for '$name', using convention '$userName' by default."
                                    }
                                    else{
                                        if ($officeLocation -eq "Everett, WA") {
                                            $userName = "$firstName".ToLower()
                                            $ouPath = "OU=Everett,OU=Washington,OU=DwellMtg,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                                            $ou = "Everett"
                                            $hasOfficeLocation = $true
                                        }
                                        elseif ($officeLocation -eq "Boyce HQ") {
                                            $userName = "$firstChar$lastName".ToLower()
                                            $ouPath = "OU=USC,OU=Pennsylvania,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                                            $ou = "USC"
                                            $hasOfficeLocation = $true
                                        }
                                        elseif ($officeLocation -eq "Lafayette, LA") {
                                            $userName = $firstName.ToLower()
                                            $ouPath = "OU=Louisiana,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                                            $ou = "Lousiana"
                                            $hasOfficeLocation = $true
                                        }
                                        elseif ($officeLocation -eq "REMOTE") {
                                            if ($branch -eq "7-Everett, WA") {
                                                $userName = "$firstName".ToLower()
                                                $ouPath = "OU=Everett,OU=Washington,OU=DwellMtg,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                                                $ou = "Everett"
                                                $hasOfficeLocation = $true    
                                            }
                                            elseif ($branch -eq "10000-Corporate") {
                                                $userName = "$firstChar$lastName".ToLower()
                                                $ouPath = "OU=USC,OU=Pennsylvania,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                                                $ou = "USC"
                                                $hasOfficeLocation = $true
                                            }
                                            else {
                                                #Write-Warning "Script is not programmed to add users to OU for '$branch'. User has been created at OU 'USC' by default." 
                                                $userName = "$firstChar$lastName".ToLower()       
                                                $ouPath = "OU=USC,OU=Pennsylvania,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                                                $ou = "USC"
                                                $hasOfficeLocation = $true
                                                #Write-Warning "Branch not recognized. Depending on it's naming convention the user '$name' might not be found. Default searching for '$userName'"
                                            }
                                        }
                                        try {
                                            Get-ADUser $userName
                                            [String[]]$userNames += $userName
                                        }
                                        catch {
                                            Write-Error $_
                                        }
                                    }
                                }
			}
		}
	}

	# create checkbox form
	$form = New-Object System.Windows.Forms.Form
	$form.StartPosition = 'CenterScreen'
	$form.size = '600,800'
	$form.Text = "Select users to remove"

	$okButton = New-Object System.Windows.Forms.Button
	$form.Controls.Add($okButton)
	$okButton.Dock = 'Bottom'
	$okButton.Height = 80
	$okButton.Font = New-Object System.Drawing.Font("Times New Roman", 18, [System.Drawing.FontStyle]::Bold)
	$okButton.Text = 'Ok'
	$okButton.DialogResult = 'Ok'

	$checkedlistbox = New-Object System.Windows.Forms.CheckedListBox
	$form.Controls.Add($checkedlistbox)
	$checkedlistbox.Dock = 'Fill'
	$checkedlistbox.CheckOnClick = $true

	$checkedlistbox.DataSource = [collections.arraylist]$userNames
	$checkedlistbox.DisplayMember = 'Caption'

	$form.ShowDialog()
        
        $size = $checkedlistbox.CheckedItems.Count
        if ($checkedlistbox.CheckedItems.Count -eq 0) {
           Write-Warning "No users were selected for deletion. Terminating program." 
        }
}

function Add-Users {

	$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
	[void]$FileBrowser.ShowDialog()
	$ExcelFile = $FileBrowser.FileName
	$objExcel = New-Excel -Path $ExcelFile
	$WorkBook = $objExcel | Get-Workbook
	$domain = "victorianfinance"
	$domainExt = "local"

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
			
			if ($manager -eq "") {
			    $hasManager = $false
			    }
			else {
			    $hasManager = $true
			    foreach($_ in $manager) {
				$mNameOut = $_.split()
				}
			    #Write-Output "mfirstname: $mNameOut[0]"
			    $mFirstName = $mNameOut[0]
			    $mFirstChar = $mFirstName.substring(0,1)
			    $mLastName = $mNameOut[1]
			    $manager = "$mFirstChar$mLastName".ToLower()
			    }
			}
		    # if header cell contains Department
		    elseif ($WorkSheet.Cells.Item(3,$j).text -eq "Department") {
			$department = $WorkSheet.Cells.Item($i,$j).text
			if ($department -eq $null) {
			    $department = "N/A"
			    }
			}
		    # if header cell is 'Branch'
		    elseif ($WorkSheet.Cells.item(3,$j).text -eq "Branch") {
			$branch = $WorkSheet.Cells.Item($i,$j).text
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
			    $ou = "Everett"
			    $hasOfficeLocation = $true
			    #$manager = "$mFirstName".ToLower()
			    }
			elseif ($officeLocation -eq "REMOTE") {
			    if ($branch -eq "7-Everett, WA") {
				$userName = "$firstName".ToLower()
				$upnSuffix = "@dwellmtg.com"
				$streetAddress = "2707 Colby Ave, Ste 1212"
				$company = "DwellMTG"
				$city = "Everett"
				$state = "WA"
				$zipCode = "98201"
				$emailAddress = "$userName$upnSuffix"
				$ouPath = "OU=RemoteUsers,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
				$ou = "RemoteUsers"
				$hasOfficeLocation = $true
				}
			    elseif ($branch -eq "10000-Corporate") {
				$userName = "$firstChar$lastName".ToLower()
				$upnSuffix = "@victorianfinance.com"
				$streetAddress = "2570 Boyce Plaza Rd"
				$company = "Victorian Finance, LLC"
				$city = "Pittsburgh"
				$state = "PA"
				$zipCode = "15241"
				$emailAddress = "$userName$upnSuffix"
				$ouPath = "OU=RemoteUsers,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
				$ou = "RemoteUsers"
				$hasOfficeLocation = $true
				#$manager = "$mFirstChar$mLastName".ToLower()
				}
			    else {
				Write-Warning "Script is not programmed to add users to OU for '$branch'. User has been created at OU 'USC' by default."
				$userName = "$firstChar$lastName".ToLower()
				$upnSuffix = "@victorianfinance.com"
				$streetAddress = "2570 Boyce Plaza Rd"
				$company = "Victorian Finance, LLC"
				$city = "Pittsburgh"
				$state = "PA"
				$zipCode = "15241"
				$emailAddress = "$userName$upnSuffix"
				$ouPath = "OU=USC,OU=Pennsylvania,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
				$ou = "USC"
				$hasOfficeLocation = $true
				#$manager = "$mFirstChar$mLastName".ToLower()
				}
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
			    $ou = "USC"
			    $hasOfficeLocation = $true
			    #$manager = "$mFirstChar$mLastName".ToLower()
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
			    $ou="Lousiana"
			    $hasOfficeLocation = $true
			    #$manager = "$mFirstName".ToLower()
			    }
			else {
			    #Write-Warning "No office location found for $name, AD account will be empty."
			    $hasOfficeLocation = $false
			    $username = "$firstChar$lastName".ToLower()
			    
			    }            
			}
		    }

		    # check and see if the generated username already exists as a user in the respective OU in Active Directory
		    if (Get-ADUser -F { samaccountname -eq $userName} ) {
			Write-Warning "User '$userName' already exists in ou '$ou'."
			}
		    else {
			#$ouPath = "OU=USC,OU=Pennsylvania,OU=Users,OU=Accounts,DC=zacklabs,DC=com"
			#Write-Output $ouPath
			#$upnSuffix = "@zacklabs.com"
			$meetsRequirements = $false 
			while (!$meetsRequirements) {
			    try {
				$password = Read-Host "password for $name ($userName)"
				#Write-Output $hasManager
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
						   'displayName'= "$firstName $lastName"; `
						   'userPrincipalName'="$userName$upnSuffix"; `
						   'mail'=$emailAddress;}
				$meetsRequirements = $true
				Write-Host "The user account '$userName' has been created." -ForegroundColor Cyan

				if ($hasManager = $true) {
				    #$targetManager = Get-ADUser -Identity $manager
				    #Write-Output "Adding manager '$manager' to '$userName'"
				    #Write-Output "userName: $userName manager: $manager"
				    Set-ADUser -Identity $userName -Manager $manager
				    
				    }
				else {
				    Write-Warning "No manager was given for $firstName $lastname so they were created without one."
				    }
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
				Write-Warning "Manager '$manager' was not found, user '$userName' was created without one." -ErrorAction Ignore
				
			    }
			    # handles any other exception and writes it to host
			    catch {
				Write-Output $_
				}
			    }
			}
			Write-Output " "
		}
	}
}


Main-Function





