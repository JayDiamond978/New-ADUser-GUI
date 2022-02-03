'''
02/01/22
Changed admin to auto detect groups and user properties to assign to new admin accounts that will be created
Changed global properties to auto detect Home Drive, Home Directory and Profile based on current admin properties
02/02/22
Edited A1 and A2 switches <-- A1 can create a2 with no need for setting manual properties
Auto set homedrive, homefolder and profile based on admin`s settings
Fixed other small bugs I found here and there
Need to put validation, only allow a1/admin accounts create other a1/admins
'''

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

        #Adding labels to forms
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
    $userTypes = Import-Csv -Path .\typeList.csv

    $global:typeList = New-Object System.Windows.Forms.ComboBox
    $typeList.Location = New-Object System.Drawing.Size(100,160)
    $typeList.Size = New-Object System.Drawing.Size(100,100)
    $form.Controls.Add($typeList)
    ForEach ($type in $userTypes)
    {
        [void] $typeList.items.Add(($type).type)
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

#This is where it gets real messy
function createADuser
{
    #Sets user's first, last, initial and username and domain
    $fName = $fBox.Text
    $iName = $iBox.Text
    $lName = $lBox.Text
    $uName = $uBox.Text
    $domain = "@$env:USERDNSDOMAIN"

    <#
    Sets up new profile, homedrive and homefolder. $userHomeDirectory and $userProfile will throw an error
    if the admin creating the account does not have a homefolder and/or profile but it can be ignored
    #>     
    $currentUserProfile = (Get-ADUser -Identity $env:USERNAME -Properties ProfilePath).ProfilePath
    $userProfile = $currentUserProfile.replace($env:USERNAME, "$uname")
    $userHomeDrive = (Get-ADUser -Identity $env:USERNAME -Properties HomeDrive).HomeDrive
    $currentUserHomeDirectory =  (Get-ADUser -Identity $env:USERNAME -Properties HomeDirectory).HomeDirectory
    $userHomeDirectory = $currentUserHomeDirectory.replace($env:USERNAME, "$uName")

    #Takes the password and converts it to a secure string
    $plainP = $pBox.Text
    $secureP = ConvertTo-SecureString $plainP -AsPlainText -Force
    
     #Sets up user's display name
    if (!$iName)
    {
        $dName = "$lName, $fName"
    }
    else 
    {
        $dName = "$lName, $fName $iName."
    }    

    #designates OU
    switch ($typeList.SelectedItem)
    {
        "general" 
        {
            $userOU = Import-Csv -Path .\general\generalUserOU.csv
            $organizationUnit = ($userOU).ou

            $userDescription = Import-Csv -Path .\general\generalDescription.csv
            $description = ($userDescription).description
        }

        "admin" 
        {
            
            #This finds the OU the current admin is in to assign to the new admin
            $adminCN = (Get-ADUser -Identity $env:USERNAME -Properties DistinguishedName).distinguishedname 
            $organizationUnit = ($adminCN -split ",", 3)[2]

            $description = (Get-ADUser -Identity $env:USERNAME -Properties description).description #gets description
            $name = "!" + $dName
            
        }

        "a1" 
        {
            
            #This finds the OU the current a1 admin is in to assign to the new a1 admin
            $a1CN = (Get-ADUser -Identity $env:USERNAME -Properties DistinguishedName).distinguishedname 
            $organizationUnit = ($a1CN -split ",", 3)[2]

            $description = (Get-ADUser -Identity $env:USERNAME -Properties description).description #gets description
            $name = "A1 " + $dName

        }

        "a2" 
        {
            #If using A1 account to create, it will look for the user's a2 account to copy properties
            if ($env:USERNAME.split(".")[0] -eq "a1")
            {
                $a2UserName = "a2" + ($env:USERNAME -split ".", 3)[2] #<-- Big sus, should really be 2 and 1
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

        "m1" 
        {
            $userOU = Import-Csv -Path .\m1\m1UserOU.csv
            $organizationUnit = ($userOU).ou

            $userDescription = Import-Csv -Path .\m1\m1Description.csv
            $description = ($userDescription).description
        }

        "m2" 
        {
            $userOU = Import-Csv -Path .\m2\m2UserOU.csv
            $organizationUnit = ($userOU).ou

            $userDescription = Import-Csv -Path .\m2\m2Description.csv
            $description = ($userDescription).description
        }

        default {$typeList.SelectedItem = $typeList.items[0]}
    }


   
    #Command to create AD user
    New-ADUser -Name $name -SamAccountName $uname -UserPrincipalName "$uname$domain" `
    -GivenName $fName -Initials $iName -Surname $lName -DisplayName $dName `
    -AccountPassword $secureP -description $description -ProfilePath $userProfile `
    -HomeDrive $userHomeDrive -HomeDirectory $userHomeDirectory -Path "$organizationUnit" `
    -Enabled $true
    

    #After user objects get created they are added to their respective groups\
    switch ($typeList.SelectedItem)
    {
        "general" 
        {
        }

        "admin" 
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

        "a1" 
        {
            $a1GroupList = (Get-ADUser $env:USERNAME -Properties MemberOf).MemberOf

            Foreach ($groupDN in $a1GroupList)
            {
                $isolateCN = $groupDN.split(",")[0]
                $groupName = $isolateCN.split("=")[1]
                Add-ADGroupMember -Identity $groupName -Members $uName
            }
        }

        "a2" 
        {
            if ($env:USERNAME.split(".")[0] -eq "a1")
            {
                $a2UserName = "a2" + ($env:USERNAME -split ".", 3)[2] #<-- Big sus, should really be 2 and 1
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

        "m1" 
        {
            $groups = Import-Csv -Path .\m1\m1UserGroups.csv
            foreach ($group in $groups)
            {
                $groupName = ($group).groups
                Add-ADGroupMember -Identity $groupName -Members $uName
            }
        }

        "m2" 
        {
            $groups = Import-Csv -Path .\m2\m2UserGroups.csv
            foreach ($group in $groups)
            {
                $groupName = ($group).groups
                Add-ADGroupMember -Identity $groupName -Members $uName
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