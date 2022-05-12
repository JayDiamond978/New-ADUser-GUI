Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Import-Module ActiveDirectory

function createForm
{   
   
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "AD User Creation"
    $form.Size = New-Object System.Drawing.Size (500,300)
    $form.FormBorderStyle = "FixedDialog"
    $form.TopMost = $true
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.StartPosition = "CenterScreen"
    $form.Font = "Segoe UI"

    modifyForms

    $form.Add_shown({$form.Activate()})
    $form.ShowDialog()

}

function modifyForms
{

    createLabels
    createTextBox
    createButton
    createDropDowns

}

function createLabels
{
    $labelList = @("First Name:", "Initials:", "Last Name:", "Username:", "Password:", "User Type:")
    $y = 10

    foreach ($labelName in $labelList)
    {

        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Size(8,$y)
        $label.Size = New-Object System.Drawing.Size(80,20)
        $label.Text = "$labelName"
        $form.controls.Add($label)

        $y = $y + 30

    }

}

function createTextBox
{
        
        #user's first name
        $script:fBox = New-Object System.Windows.Forms.TextBox
        $fBox.Location = New-Object System.Drawing.Size(100,10)
        $fBox.Size = New-Object System.Drawing.Size(100,100)
        $form.Controls.Add($fBox)
        
        #initials
        $script:iBox = New-Object System.Windows.Forms.TextBox
        $iBox.Location = New-Object System.Drawing.Size(100,40)
        $iBox.Size = New-Object System.Drawing.Size(100,100)
        $form.Controls.Add($iBox)

        #user's last name
        $script:lBox = New-Object System.Windows.Forms.TextBox
        $lBox.Location = New-Object System.Drawing.Size(100,70)
        $lBox.Size = New-Object System.Drawing.Size(100,100)
        $form.Controls.Add($lBox)

        #user's username
        $script:uBox = New-Object System.Windows.Forms.TextBox
        $uBox.Location = New-Object System.Drawing.Size(100,100)
        $uBox.Size = New-Object System.Drawing.Size(100,100)
        $form.Controls.Add($uBox)

        #user's password
        $script:pBox = New-Object System.Windows.Forms.MaskedTextBox
        $pBox.Location = New-Object System.Drawing.Size(100,130)
        $pBox.Size = New-Object System.Drawing.Size(100,100)
        $pbox.passwordchar = "*"
        $form.Controls.Add($pBox)

}

function createDropDowns
{
    #Detects if AC3 based on current admin's username, limits options in dropdown
    if ($env:USERNAME -like "a1.*" -or $env:USERNAME -like "a2.*")
    {
        $userTypes = @("General", "M1", "M2", "A1", "A2")     
    }
    else
    {
        $userTypes = @("General", "Admin")
    }

    $script:typeList = New-Object System.Windows.Forms.ComboBox
    $typeList.Location = New-Object System.Drawing.Size(100,160)
    $typeList.Size = New-Object System.Drawing.Size(100,100)
    $form.Controls.Add($typeList)
    ForEach ($type in $userTypes)
    {
        [void] $typeList.items.Add($type)
    }
}

function createButton
{

    $button = New-Object System.Windows.Forms.Button
    $button.Size = New-Object System.Drawing.Size(50,25)
    $button.location = New-Object System.Drawing.Size(400,220)
    $button.TextAlign = "MiddleCenter"
    $button.text = "OK"
    $button.Add_Click({createADuser})
    $form.Controls.Add($button)

}

#This is where it gets real messy at some point this will get separated into multiple functions
function createADuser
{
    #Sets user's first, initial, last, username and domain in that order based on the response from the form
    $fName = $fBox.Text
    $iName = $iBox.Text
    $lName = $lBox.Text
    $uName = $uBox.Text
    $domain = "@$env:USERDNSDOMAIN"

    #Sets user's roaming profile, homedrive and homedirectory all based on the current admin's properties
    $currentUserProfile = (Get-ADUser -Identity $env:USERNAME -Properties ProfilePath).ProfilePath
    $userProfile = $currentUserProfile.replace($env:USERNAME, "$uname") #<-- Throws an error if profile is empty
    $userHomeDrive = (Get-ADUser -Identity $env:USERNAME -Properties HomeDrive).HomeDrive 
    $currentUserHomeDirectory =  (Get-ADUser -Identity $env:USERNAME -Properties HomeDirectory).HomeDirectory
    $userHomeDirectory = $currentUserHomeDirectory.replace($env:USERNAME, "$uName") #<-- Throws an error if no homedir

    #Takes the password from the password textbox and converts it to a secure string
    $plainP = $pBox.Text
    $secureP = ConvertTo-SecureString $plainP -AsPlainText -Force
    
    #Sets up user's display name
    if (!$iName)
    {$dName = "$lName, $fName"}

    else 
    {$dName = "$lName, $fName $iName."}    

    #This section sets the user's OU and description <-- General has not been set up yet
    switch ($typeList.SelectedItem)
    {
        "General" 
        {
            #Need OU, Description and name
            $adminCN = (Get-ADUser -Identity $env:USERNAME -Properties DistinguishedName).distinguishedname
            $tempList = $adminCN -split ","
            $organizationUnit = "OU=CAUsers,"
            $index = 1
            foreach ($i in $tempList)
            {
                if ($i -like "*user*") 
                {
                    $index += 1
                    $organizationUnit = $organizationUnit + ($adminCN -split ",", $index)[-1]
                }
                else {$index += 1}
            }
            $description = "Domain User"
            $name = $dName
        }

        "Admin" 
        {
            #This finds the OU the current admin is in to assign to the new admin
            $adminCN = (Get-ADUser -Identity $env:USERNAME -Properties DistinguishedName).distinguishedname 
            $organizationUnit = ($adminCN -split ",", 3)[2]

            $description = (Get-ADUser -Identity $env:USERNAME -Properties description).description #gets description
            $name = "!" + $dName 
        }

        "A1" 
        {
            #This finds the OU the current a1 admin is in to assign to the new a1 admin
            $a1CN = (Get-ADUser -Identity $env:USERNAME -Properties DistinguishedName).distinguishedname 
            $organizationUnit = ($a1CN -split ",", 3)[2]

            $description = (Get-ADUser -Identity $env:USERNAME -Properties description).description #gets description
            $name = "A1 " + $dName
        }

        "A2" 
        {
            #If using A1 account to create, it will look for the user's a2 account to copy properties
            if ($env:USERNAME.split(".")[0] -eq "a1")
            {
                $a2UserName = "a2" + ($env:USERNAME -split ".", 3)[2]
                $a2CN = (Get-ADUser -Identity $a2UserName -Properties DistinguishedName).distinguishedname 
                $organizationUnit = ($a2CN -split ",", 2)[1]
                $description = (Get-ADUser -Identity $env:USERNAME -Properties description).description #gets description
                $name = "A2 " + $dName
            }

            else
            {
                $a2CN = (Get-ADUser -Identity $env:USERNAME -Properties DistinguishedName).distinguishedname 
                $organizationUnit = ($a2CN -split ",", 3)[2]
                $description = (Get-ADUser -Identity $env:USERNAME -Properties description).description #gets description
            }
        }

        #M1 and M2 accounts get queried, first active m1 account is used as reference
        "M1" 
        {
            $m1SamName = ((get-aduser -Filter "UserPrincipalName -like 'm1.*' -and Enabled -eq 'true'").samaccountname[0])
            $m1CN = (get-aduser -Identity $m1SamName).distinguishedname
            $organizationUnit = ($m1CN -split ",", 3)[2]
            $description = (Get-ADUser -Identity $m1SamName -Properties description).description
            $name = "M1 " + $dName
        }

        "M2" 
        {
            $m2SamName = ((get-aduser -Filter "UserPrincipalName -like 'm2.*' -and Enabled -eq 'true'").samaccountname[0])
            $m2CN = (get-aduser -Identity $m2SamName).distinguishedname
            $organizationUnit = ($m1CN -split ",", 3)[2]
            $description = (Get-ADUser -Identity $m2SamName -Properties description).description
            $name = "M2 " + $dName
        }

        default {$typeList.SelectedItem = $typeList.items[0]} #If dropdown is blank it defaults to general user
    }
   
    #Command to create AD user
    New-ADUser -Name $name -SamAccountName $uname -UserPrincipalName "$uname$domain" `
    -GivenName $fName -Initials $iName -Surname $lName -DisplayName $dName `
    -AccountPassword $secureP -description $description -ProfilePath $userProfile `
    -HomeDrive $userHomeDrive -HomeDirectory $userHomeDirectory -Path "$organizationUnit" `
    -Enabled $true
    

    #After user objects get created they are added to their respective groups <-- Need to set up General
    switch ($typeList.SelectedItem)
    {
        "General" 
        {
        }

        "Admin" 
        {
            #This finds the groups the current admin is in to assign to the new admin
            $adminGroupList = (Get-ADUser $env:USERNAME -Properties MemberOf).MemberOf

            Foreach ($groupDN in $adminGroupList)
            {
                $isolateCN = $groupDN.split(",")[0]
                $groupName = $isolateCN.split("=")[1]
                Add-ADGroupMember -Identity $groupName -Members $uName
            }
        }

        "A1" 
        {
            $a1GroupList = (Get-ADUser $env:USERNAME -Properties MemberOf).MemberOf

            Foreach ($groupDN in $a1GroupList)
            {
                $isolateCN = $groupDN.split(",")[0]
                $groupName = $isolateCN.split("=")[1]
                Add-ADGroupMember -Identity $groupName -Members $uName
            }
        }

        "A2"
        {
            #If using A1 account to create, it will look for the user's a2 account to copy groups
            if ($env:USERNAME.split(".")[0] -eq "a1")
            {
                $a2UserName = "a2" + ($env:USERNAME -split ".", 3)[2]
                $a2GroupList = (Get-ADUser $a2UserName -Properties MemberOf).MemberOf
                Foreach ($groupDN in $a2GroupList)
                {
                    $isolateCN = $groupDN.split(",")[0]
                    $groupName = $isolateCN.split("=")[1]
                    Add-ADGroupMember -Identity $groupName -Members $uName
                }
            }

            else
            {
                $a2GroupList = (Get-ADUser $env:USERNAME -Properties MemberOf).MemberOf

                    Foreach ($groupDN in $a2GroupList)
                    {
                        $isolateCN = $groupDN.split(",")[0]
                        $groupName = $isolateCN.split("=")[1]
                        Add-ADGroupMember -Identity $groupName -Members $uName
                    }
            }
        }

        <#Finds same typed user and loops through assigned groups, if groups contains all same type user -1 (The new user) 
        Then add new user to group                                                                             #>
        "M1" 
        {
            $totalM1Users = (get-aduser -Filter "SamAccountName -like 'm1.*' -and Enabled -eq 'true'").SamAccountName
            $m1SamName = (get-aduser -Filter "SamAccountName -like 'm1.*' -and Enabled -eq 'true'").SamAccountName[0]
            $m1Groups = (Get-ADUser -Identity $m1SamName -Properties MemberOf).MemberOf

            $allM1UserDN = (get-aduser -Filter "SamAccountName -like 'm1.*' -and Enabled -eq 'true'").DistinguishedName
            

            foreach ($groupDN in $m1Groups)
            {
                $matchCount = 0
                $isolateCN = $groupDN.split(",")[0]
                $groupName = $isolateCN.split("=")[1]
                $allGroupUsers = (Get-ADGroup -Identity "$groupName" -Properties Members).Members #Saves DN of all users in group
                
                foreach ($m1UserDN in $allM1UserDN)
                {
                    if ($allGroupUsers -like $m1UserDN)
                    {
                        $matchCount += 1
                        if ($matchcount -eq ($allM1UserDN.count -1))
                        {
                            Add-ADGroupMember -Identity $groupName -Members $uName  
                
                        }
                    }
                    else {break}
                }
            }                
        }
        

        "M2" 
        {
            $totalM2Users = (get-aduser -Filter "SamAccountName -like 'm2.*' -and Enabled -eq 'true'").SamAccountName
            $m2SamName = (get-aduser -Filter "SamAccountName -like 'm2.*' -and Enabled -eq 'true'").SamAccountName[0]
            $m2Groups = (Get-ADUser -Identity $m2SamName -Properties MemberOf).MemberOf

            $allM2UserDN = (get-aduser -Filter "SamAccountName -like 'm2.*' -and Enabled -eq 'true'").DistinguishedName
            

            foreach ($groupDN in $m2Groups)
            {
                $matchCount = 0
                $isolateCN = $groupDN.split(",")[0]
                $groupName = $isolateCN.split("=")[1]
                $allGroupUsers = (Get-ADGroup -Identity "$groupName" -Properties Members).Members #Saves DN of all users in group
                
                foreach ($m2UserDN in $allM2UserDN)
                {
                    if ($allGroupUsers -like $m2UserDN)
                    {
                        $matchCount += 1
                        if ($matchcount -eq ($allM1UserDN.count -1))
                        {
                            Add-ADGroupMember -Identity $groupName -Members $uName  
                
                        }
                    }
                    else {break}
                }
            }                
        }
        
        default {$typeList.SelectedItem = $typeList.items[0]}
    }


    $form.Close()

}

function main
{
    createForm
}

main
