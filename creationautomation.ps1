#### PowerShell script to take an Excel.xlsx file, parse it for desired information, and use that information to 
#### both add and disable Active Directory users

## TODO as of 6/27/2022
# * Change how script determines preexisting users from being hardcoded to an editable json file
# * remove direct reports

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
		        connectExchange
                        Disable-Users
			}
	else {
		Write-Warning "Invalid argument. Use 'add' or 'remove'."
		}
	}
}

function connectExchange {

        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

        if ($isAdmin -eq $false) {
            Write-Warning "Please run script with Administrator privileges"
            exit
        }
        else {
            $currentSessions = Get-PSSession
            #write-warning "sessions: $currentSessions"
            if ($currentSessions -eq $null) {
                Write-Host "NO CURRENT SESSIONS"
                try {
                    $UserCredential = Get-Credential
                    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionURI https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
                    Import-PSSession $Session
                }
                catch {
                    write-warning $_
                }
            }
            else {
                Write-Host "some sessions"
                Remove-PSSession *
                try {
                    $UserCredential = Get-Credential
                    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionURI https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
                    Import-PSSession $Session
                }
                catch {
                    write-warning $_
                }
            }
        }
}
function Disable-Users {
        #connectExchange
        
        # select the file and define the domain/domain extension
	$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
	[void]$FileBrowser.ShowDialog()
	$ExcelFile = $FileBrowser.FileName
	$objExcel = New-Excel -Path $ExcelFile
	$WorkBook = $objExcel | Get-Workbook
	$domain = "victorianfinance"
	$domainExt = "local"

        # create form template
        $newForm = New-Object System.Windows.Forms.Form
        $newForm.topmost=$true
        $newForm.Text="Select Users to Disable/Wipe"
        $newForm.Location.x=400
        $newForm.Location.y=400
        $newform.size=New-Object System.Drawing.Size(900,600)
        $newForm.MaximumSize=New-Object System.Drawing.Size(900,600)
        $newForm.MinimumSize=New-Object System.Drawing.Size(900,600)
        
        # create main listview
        $listview = New-Object System.Windows.Forms.ListView
        $listview.Location = New-Object System.Drawing.Size(25,25)
        #$listview.size = New-Object System.Drawing.Size(350, 503)
        $listview.size = New-Object System.Drawing.Size(350, 437)
        $listview.Checkboxes=$true
        $listView.name="main"
        $listView.autoarrange=$true
        $listview.gridlines=$true
        $listview.multiselect=$true
        $listview.View = "details"
        $listview.headerstyle = 1
        $listview.columns.add("User", -2) | out-null
        $listview.columns.add("Name", -2) | out-null
        $listview.columns.add("Term Date", -2) |out-null

        # create the 'search' button
        $searchBtn = New-Object System.Windows.Forms.Button
        $searchBtn.Location = New-Object System.Drawing.Size(398,70)
        $searchBtn.size = New-Object System.Drawing.Size(75,23)
        $searchBtn.Text = "Find"
        $searchBtn_Click = {
            #Write-Host "Click"
            if ($searchField.Text -ne "") {
                try {
                    if (Get-ADUser -Filter {displayName -like $searchField.Text}) {
                        Write-Host "User exists"
                        #$Name = Get-ADUser -Identity $searchField.Text -Properties Name | Select-Object -ExpandProperty Name
                        #$OUpath = Get-ADUser -Identity $searchField.Text -Properties DistinguishedName | Select-Object -ExpandProperty DistinguishedName
                        $Name = Get-ADUser -Filter {displayName -Like $searchField.Text} -Properties samAccountName | Select-Object -ExpandProperty samAccountName
                        $OUpath = Get-ADUser -Filter {displayName -Like $searchField.Text} -Properties DistinguishedName | Select-Object -ExpandProperty DistinguishedName
                        $OUpatharray = $OUpath -split ","
                        $resultview.items.clear()
                        $item = New-Object System.Windows.Forms.ListViewItem($searchField.Text)
                        $item.subitems.add($Name) | out-null
                        $item.subitems.add($OUpatharray[1] + "," + $OUpatharray[2] + "," + $OUpatharray[3])
                        $resultview.items.add($item) | out-null
                        $resultview.AutoResizeColumns(1)
                        $addBtn.Enabled = $true
                    }
                    else {
                        Write-Host "No users found"
                        $resultview.items.clear()
                        $item = New-Object System.Windows.Forms.ListViewItem("DNE")
                        $resultview.items.add($item)
                        $resultview.AutoResizeColumns(1)
                        $addBtn.Enabled = $false
                    }
                }
                catch {
                    #Write-Host $_
                    Write-Host "User does not exist"
                    $resultview.items.clear()
                    $item = New-Object System.Windows.Forms.ListViewItem("DNE")
                    $resultview.items.add($item)
                    $resultview.AutoResizeColumns(1)
                    $addBtn.Enabled = $false
                }
            }
            else {
                Write-Host "blank field"
                $addBtn.Enabled = $false
                $resultview.items.clear()
            }
        } 
        $searchBtn.Add_Click($SearchBtn_Click)

        Write-Output $searchBtn.Size

        # create the search input field
        $searchField = New-Object System.Windows.Forms.TextBox
        $searchField.Location = New-Object System.Drawing.Point(478,70)
        $searchField.Size = New-Object System.Drawing.Size(378,20)
        $searchField.Text = "Search AD by full name"
        $searchField.Add_KeyDown({
            if ($_.KeyCode -eq "Return") {
                #$searchBtn_Click
                $_.SuppressKeyPress = $true

                if ($searchField.Text -ne "") {
                    try {
                        if (Get-ADUser -Filter {displayName -like $searchField.Text}) {
                            Write-Host "User exists"
                            #$Name = Get-ADUser -Identity $searchField.Text -Properties Name | Select-Object -ExpandProperty Name
                            #$OUpath = Get-ADUser -Identity $searchField.Text -Properties DistinguishedName | Select-Object -ExpandProperty DistinguishedName
                            $Name = Get-ADUser -Filter {displayName -Like $searchField.Text} -Properties samAccountName | Select-Object -ExpandProperty samAccountName
                            $OUpath = Get-ADUser -Filter {displayName -Like $searchField.Text} -Properties DistinguishedName | Select-Object -ExpandProperty DistinguishedName
                            $OUpatharray = $OUpath -split ","
                            $resultview.items.clear()
                            $item = New-Object System.Windows.Forms.ListViewItem($searchField.Text)
                            $item.subitems.add($Name) | out-null
                            $item.subitems.add($OUpatharray[1] + "," + $OUpatharray[2] + "," + $OUpatharray[3])
                            $resultview.items.add($item) | out-null
                            $resultview.AutoResizeColumns(1)
                            $addBtn.Enabled = $true
                        }
                        else {
                            Write-Host "No users found"
                            $resultview.items.clear()
                            $item = New-Object System.Windows.Forms.ListViewItem("DNE")
                            $resultview.items.add($item)
                            $resultview.AutoResizeColumns(1)
                            $addBtn.Enabled = $false
                        }
                    }
                    catch {
                        #Write-Host $_
                        Write-Host "User does not exist"
                        $resultview.items.clear()
                        $item = New-Object System.Windows.Forms.ListViewItem("DNE")
                        $resultview.items.add($item)
                        $resultview.AutoResizeColumns(1)
                        $addBtn.Enabled = $false
                    }
                }
                else {
                    Write-Host "blank field"
                    $addBtn.Enabled = $false
                    $resultview.items.clear()
                }
            }

        })


        # create the search result listbox
        $resultview = New-Object System.Windows.Forms.ListView
        $resultview.Location = New-Object System.Drawing.Size(398,98)
        $resultview.size = New-Object System.Drawing.Size(458,22)
        $resultview.Checkboxes=$false
        $resultview.name="search"
        $resultview.autoarrange=$true
        $resultview.gridlines=$true
        $resultview.multiselect=$false
        $resultview.View = "details"
        $resultview.columns.add("User_____", -2) | out-null
        $resultview.columns.add("Name_____", -2) | out-null
        $resultview.columns.add("OUPath__________________", -2) | out-null
        $resultview.headerstyle = 0
        

        # create the 'Add' button
        $addBtn = New-Object System.Windows.Forms.Button
        $addBtn.Location = New-Object System.Drawing.Size(398,124)
        $addBtn.Size = New-Object System.Drawing.Size(458,22)
        $addBtn.Text = "Add"
        $addBtn.Enabled = $false
        $addBtn_Click = {
            Write-Host "Click"
            $Name = Get-ADUser -Filter {displayName -like $searchField.Text} -Properties Name | Select-Object -ExpandProperty Name
            $samAccountName = Get-ADUser -Filter {displayName -Like $searchField.Text} -Properties samAccountName | Select-Object -ExpandProperty samAccountName
            Write-Host $samAccountName;
            $item = New-Object System.Windows.Forms.ListViewItem($samAccountName)
            $item.subitems.add($Name)
            $item.subitems.add("N/A")
            $listview.items.add($item)
            
        }
        $addBtn.Add_Click($addBtn_Click)
        
        # create 'nuke' button'
        $nukeBtn = New-Object System.Windows.Forms.Button
        $nukeBtn.Location = New-Object System.Drawing.Size(398,487)
        $nukeBtn.Size = New-Object System.Drawing.Size(458, 40)
        $nukeBtn.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 12, [System.Drawing.Fontstyle]::Bold)
        $nukeBtn.Text = "NUKE"
        $nukeBtn_Click = {
            #Write-Host "Nuke"
            $progressBar.Visible = $true
            $count = $listview.CheckedItems.Count
            $percent

            foreach($item in $listview.CheckedItems) {
                $_ = $item.text

                #retrieve email through upn
                $emailaddress = Get-ADUser -Identity $_ -Properties UserPrincipalName | select-object -expandproperty userprincipalname
                #get distinuishedname of manager
                $managerName = Get-ADUser -Identity $_ -Properties manager | Select-Object -expandproperty manager

                
                Write-Host $_
                Write-Host $emailAddress
                if ($managerName -ne $null) {
                    $managername = $managerName.split(",")[0].replace('CN=','')
                    $managerEmail = Get-ADUser -Filter "Name -eq '$managerName'" | select-object -expandproperty userprincipalname
                    $message = "I'm out of the office, please contact $managerName at ", $managerEmail.ToLower()
                    Write-Host $managerName

                }
                else {

                    $message = "I'm out of the office, please contact our main office at (888)333-0191"
                    Write-Host "NULL MANAGER"
                }

                Write-Host $message
                Write-Host ""
    
                # set out of office
                try {
                    $progressBar.text = "Setting out of office message for $_..."
                    Set-MailboxAutoReplyConfiguration -Identity $emailAddress -AutoReplyState Enabled -InternalMessage $message -ExternalMessage $message
                    Write-Host "Set out of office for $_" -foregroundcolor cyan     
                    try {
                        $progressBar.text = "Disabling and wiping info for $_..."
                        # disable users
                        $userName = Get-ADUser -Identity $_ -Properties SamAccountName | Select-Object -ExpandProperty SamAccountName
                        Set-ADUser `
                        -Identity $_ `
                        -Enabled $false `
                        -Clear @('mail', 'title', 'department', 'company', 'manager', 'mobile', 'postalCode', 'st', 'streetAddress', 'telephoneNumber', 'url', 'physicalDeliveryOfficeName', 'l')
                        Write-Host "Disabled $_" -ForegroundColor Cyan

                        $percent += (100/$count)
                        $progressBar.Value = $percent
                        $progressBar.text = "DONE" 
                    }
                    catch {
                        Write-Host $_
                        $progressBar.visible = $false
                    }
                }
                catch {
                    Write-Host $_
                    $progressBar.visible = $false
                }
                if ($progressbar.value -eq 100) {
                    Start-Sleep -s 1
                    $progressBar.visible = $false
                    $progressBar.value = 0
                }
            }
            $progressBar.value = 0
            $progressBar.visible = $false
        }
        $nukeBtn.Add_Click($nukeBtn_Click)

        # create the 'select all' button
        $selectBtn = New-Object System.Windows.Forms.Button
        $selectBtn.Location = New-Object System.Drawing.Size(25,487)
        $selectBtn.Size = New-Object System.Drawing.Size(172, 40)
        $selectBtn.Text = "Select All"
        $selectBtn_Click = {
            foreach ($item in $listview.items) {
               $item.Checked = $true
            }
        }
        $selectBtn.Add_Click($selectBtn_click)

        #create the "unselect all' button
        $unSelectBtn = New-Object System.Windows.Forms.Button
        $unSelectBtn.Location = New-Object System.Drawing.Size(203,487)
        $unSelectBtn.Size = New-Object System.Drawing.Size(172, 40)
        $unSelectBtn.Text = "Deselect All"
        $unSelectBtn_Click = {
            foreach ($item in $listview.items) {
               $item.Checked = $false
            }
        }
        $unSelectBtn.Add_Click($unSelectBtn_click)

        # create progress bar
        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Name = 'progressBar'
        $progressBar.Value = 0
        $progressBar.Style = "Continuous"
        $progressBar.Size = New-Object System.Drawing.Size(456,30)
        $progressBar.Location = New-Object System.Drawing.Size(400,440)
        $progressBar.visible = $false
        $progressBar.RightToLeftLayout = $false

        # create progress bar text overlay
        $barOverlay = New-Object System.Windows.Forms.Label
        $barOverlay.Size = New-Object System.Drawing.Size(400,40)
        $barOverlay.Location = New-Object system.Drawing.Size(456,440)
        $barOverlay.text = "SAMPLE TEXT SAMPLE TEXT SAMPLE TEXT"
        $barOverlay.visible = $false
        #$barOverlay.controls.bringtofront = $true

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
                                elseif ($WorkSheet.Cells.Item(3,$j).text -eq "Termination Date") {
                                    $termDate = $WorkSheet.Cells.item($i,$j).text
                                }
                                elseif ($WorkSheet.Cells.Item(3,$j).text -eq "Office Location") {
                                    $officeLocation = $WorkSheet.Cells.item($i,$j).text
                                    if ($officeLocation -eq "") {
                                        $hasOfficeLocation = $false
                                        $userName = "$firstChar$lastName".ToLower()
                                        #$userName = "$lastname".ToLower()
                                        Write-Warning "No office location found for '$name', search will use 'firstCharlastName' by default."
                                    }
                                    else{
                                        $hasOfficeLocation = $true
                                        if (($officeLocation -eq "Everett, WA") -or ($officeLocation -eq "Gilbert AZ")) {
                                            $userName = "$firstName".ToLower()
                                            $ouPath = "OU=Everett,OU=Washington,OU=DwellMtg,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                                            $ou = "Everett"
                                        }
                                        elseif ($officeLocation -eq "Boyce HQ") {
                                            $userName = "$firstChar$lastName".ToLower()
                                            $ouPath = "OU=USC,OU=Pennsylvania,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                                            $ou = "USC"
                                        }
                                        elseif ($officeLocation -eq "Lafayette, LA") {
                                            $userName = $firstName.ToLower()
                                            $ouPath = "OU=Louisiana,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                                            $ou = "Lousiana"
                                        }
                                        elseif ($officeLocation -eq "Panama City, FL") {
                                            $userName = "$firstChar$lastName".ToLower()
                                            $ouPath = "OU=OasisMTG,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                                            $ou = "OasisMTG"
                                        }
                                        elseif ($officeLocation -eq "REMOTE") {
                                            if ($branch -eq "7-Everett, WA") {
                                                $userName = "$firstName".ToLower()
                                                $ouPath = "OU=Everett,OU=Washington,OU=DwellMtg,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                                                $ou = "Everett"
                                            }
                                            elseif ($branch -eq "10000-Corporate") {
                                                $userName = "$firstChar$lastName".ToLower()
                                                $ouPath = "OU=USC,OU=Pennsylvania,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                                                $ou = "USC"
                                            }
                                            else {
                                                #Write-Warning "Script is not programmed to add users to OU for '$branch'. User has been created at OU 'USC' by default." 
                                                $userName = "$firstChar$lastName".ToLower()       
                                                $ouPath = "OU=USC,OU=Pennsylvania,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                                                $ou = "USC"
                                                #Write-Warning "Branch not recognized. Depending on it's naming convention the user '$name' might not be found. Default searching for '$userName'"
                                            }
                                        }
                                        else {
                                            #Write-Warning "Unrecognized office location '$officeLocation' for user '$name', program cannot determine naming convention to search for if no location info is provided."
                                            #$Write-Output ""
                                            $userName = ""
                                        }
                                    }
                                    try {
                                        if ($userName -ne "") {
                                            write-output $username
                                            if((Get-ADUser $userName)) {
                                                #Write-Output "$username exists"
                                                #[String[]]$userNames += $userName
                                                #Write-Output "$_"
                                                $item = New-Object System.Windows.Forms.ListViewItem($userName)
                                                $item.subitems.add($name) | out-null
                                                $item.subitems.add($termDate) | out-null
                                                $listview.items.add($item) | out-null
                                            }
                                        }
                                        else {
                                            break
                                        }
                                    }
                                    catch {
                                        #Write-Output $_
                                        #Write-Output ""
                                        Write-Warning "User '$name' ($userName) was not found in Active Directory. "
                                    }
                                }
			}
		}
	}

        $newForm.controls.add($searchBtn)
        $newForm.controls.add($searchField)
        $newForm.controls.add($resultview)
        $newForm.controls.add($addBtn)
        $newForm.controls.add($nukeBtn)
        $newForm.controls.add($selectBtn)
        $newForm.controls.add($unSelectBtn)
        $newForm.controls.add($progressBar)
        $newForm.controls.add($barOverlay)
        $newForm.controls.add($listview) 


        $newForm.showdialog()



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
			$Name = $WorkSheet.Cells.Item($i,$j).text
                        # since HR will not put consistent item names in the boards, the name field must be brute forced into "firstname lastname" format with regex		
                        $Name = ($Name -replace '\(.*\) ', '')
			foreach($_ in $name) {
					$nameOut = $_.split()
				}
                        $firstName = $nameOut[0]
                        $firstChar = $firstName.substring(0,1)
                        $lastName = $nameOut[1]
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
                            $unknownBranch = $false
			    #$manager = "$mFirstName".ToLower()
			}
                        elseif ($officeLocation -eq "Hurricane, WV") {
                            $userName = "$firstname".ToLower()
                            $upnSuffix = "@myhomloan.com"
                            $streetAddress = "3818 Teays Valley Road"
                            $company = "Victorian Finance, LLC."
                            $city = "Hurricane"
                            $zipCode = "25526"
                            $emailAddress = "$userName$upnSuffix"
                            $ouPath = "OU=Teays Valley,OU=West Virginia,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                            $ou = "Teays Valley"
                            $hasOfficeLocation = $true
                            $unknownBranch = $false
                            write-host "HURRICANE"
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
                                $unknownBranch = $false
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
                                $unknownBranch = $false
				#$manager = "$mFirstChar$mLastName".ToLower()
				}
                            elseif ($branch -eq "50-Wyrick, WV") {    
                                $userName = "$firstname".ToLower()
                                $upnSuffix = "@myhomloan.com"
                                $streetAddress = "3818 Teays Valley Road"
                                $company = "Victorian Finance, LLC."
                                $city = "Hurricane"
                                $zipCode = "25526"
                                $emailAddress = "$userName$upnSuffix"
                                $ouPath = "OU=Teays Valley,OU=West Virginia,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                                $ou = "Teays Valley"
                                $hasOfficeLocation = $true
                                $unknownBranch = $false
                            }
			    else {
                                $unknownBranch = $true
				#Write-Warning "Script is not programmed to fill in user information for REMOTE office locations with branch '$branch'. User will be created under OU RemoteUsers with Boyce location information."
				$userName = "$firstChar$lastName".ToLower()
				$upnSuffix = "@victorianfinance.com"
				$streetAddress = "2570 Boyce Plaza Rd"
				$company = "Victorian Finance, LLC"
				$city = "Pittsburgh"
				$state = "PA"
				$zipCode = "15241"
				$emailAddress = "$userName$upnSuffix"
				$ouPath = "OU=RemoteUsers,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
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
                            $unknownBranch = $false
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
                            $unknownBranch = $false
			    #$manager = "$mFirstName".ToLower()
			    }
                        elseif ($officeLocation -eq "Panama City, FL") {
                            $userName = "$firstName".ToLower()
                            $upnSuffix = "@oasismortgage.net"
                            $streetAddress = "160 Oasis, Panama City,  FL"
                            $company = "Oasis Mortgage"
                            $city = "Panama City"
                            $state = "FL"
                            $zipCode = "32405"
                            $emailAddress = "$userName$upnSuffix"
			    $ouPath = "OU=OasisMTG,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
                            $ou = "OasisMTG"
                            $hasOfficeLocation = $true
                            $unknownBranch = $false
                            }
			else {
			    Write-Warning "Office location '$officeLocation' for '$name' is not recognized, account will be created under USC by default with empty location information."
			    $hasOfficeLocation = $false
			    $username = "$firstChar$lastName".ToLower()
			    $ouPath = "OU=USC,OU=Pennsylvania,OU=Users,OU=Accounts,DC=$domain,DC=$domainExt"
			    }            
			}
		    # if header cell contains Job Title
		    elseif ($WorkSheet.Cells.Item(3,$j).text -eq "Job Title") {
			$jobTitle = $WorkSheet.Cells.Item($i,$j).text
                        $loanOfficer = "Group_e4f091c5-2b89-47aa-ab1a-7f7275aa3a98"
                         
			if ($jobTitle -eq $null) {
			    $jobTitle = "N/A"
                            Write-Warning "No job title given"
			    }
                        elseif ($jobTitle -eq "Loan Officer") {
                            Add-ADGroupMember -Identity $loanOfficer -members $userName
                        }
		    }
                }


		    # check and see if the generated username already exists as a user in the respective OU in Active Directory
		    if (Get-ADUser -F { displayName -eq $Name} ) {
			Write-Warning "User '$Name' already exists in ou '$ou'."
		    }
		    else {

                        #check if username exists in AD already, if so, change naming convention to first.last
                        try {
                            if (Get-ADUser -Identity $userName) {
                                $existingName = Get-ADUser -Identity $userName -Properties Name | Select-Object -ExpandProperty Name
                                Write-Warning "Username '$userName' is taken by '$existingName', using '$firstName.$lastName' instead"
                                $userName = "$firstname.$lastName".ToLower()
                            }
                        }
                        catch {
                            #write-warning "No user found with name '$Name' or samAccountName '$userName'"
                        }

			#$ouPath = "OU=USC,OU=Pennsylvania,OU=Users,OU=Accounts,DC=zacklabs,DC=com"
			#Write-Output $ouPath
			#$upnSuffix = "@zacklabs.com"
			$meetsRequirements = $false 
			while (!$meetsRequirements) {
			    try {
                                if ($hasOfficeLocation) {
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
                                }
                                else {
                                    $password = Read-Host "password for $name ($userName)"
                                    New-ADUser `
                                    -Enabled $true `
                                    -Path  $ouPath `
                                    -Name $name `
                                    -SamAccountName $userName `
                                    -GivenName $firstName `
                                    -Surname $lastName `
                                    -Company $company `
                                    -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -ChangePasswordAtLogon $False `
                                    -OtherAttributes @{'title'=$jobTitle; `
                                                       'department'=$department; `
                                                       'displayName'="$firstName $lastName"; `
                                                       'userPrincipalName'="$userName$upnSuffix";}
                                }
				$meetsRequirements = $true
                                
                                if ($unknownBranch) {
				    Write-Host "The user account '$userName' has been created." -ForegroundColor Cyan
                                    Write-Warning "Script is not programmed to fill in user information for REMOTE office locations with branch '$branch'. User will be created under OU RemoteUsers with Boyce location information."
                                } else {
                                     
				    Write-Host "The user account '$userName' has been created." -ForegroundColor Cyan
                                }

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
                            catch [Microsoft.ActiveDirectory.Management.ADIdentityAlreadyExistsException] {

                            }
			    # handles any other exception and writes it to host
			    catch {
                                Write-Host "EEEE"
				Write-Output $_
				}
			    }
			}
			Write-Output " "
		}
	}
}

Main-Function
