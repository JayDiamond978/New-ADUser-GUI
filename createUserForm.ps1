Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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
        $global:fBox = New-Object System.Windows.Forms.TextBox
        $fBox.Location = New-Object System.Drawing.Size(100,10)
        $fBox.Size = New-Object System.Drawing.Size(100,100)
        $form.Controls.Add($fBox)
        
        #initials
        $global:iBox = New-Object System.Windows.Forms.TextBox
        $iBox.Location = New-Object System.Drawing.Size(100,40)
        $iBox.Size = New-Object System.Drawing.Size(100,100)
        $form.Controls.Add($iBox)

        #user's last name
        $global:lBox = New-Object System.Windows.Forms.TextBox
        $lBox.Location = New-Object System.Drawing.Size(100,70)
        $lBox.Size = New-Object System.Drawing.Size(100,100)
        $form.Controls.Add($lBox)

        #user's username
        $global:uBox = New-Object System.Windows.Forms.TextBox
        $uBox.Location = New-Object System.Drawing.Size(100,100)
        $uBox.Size = New-Object System.Drawing.Size(100,100)
        $form.Controls.Add($uBox)

        #user's password
        $global:pBox = New-Object System.Windows.Forms.MaskedTextBox
        $pBox.Location = New-Object System.Drawing.Size(100,130)
        $pBox.Size = New-Object System.Drawing.Size(100,100)
        $global:pbox.passwordchar = "*"
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

    #Extra Properties

    #global properties that apply to everyone
    #properties.csv | OU, user profile, mapped home drive, home directory
    $userProps = Import-Csv -Path .\properties.csv
    foreach ($d in $userProps)
    {
        $organizationUnit = ($d).OU  #this needs to be moved to specific
        $userProfile = ($d).profile
        $userHomeDrive = ($d).homeDrive
        $userHomeDirectory = ($d).homeDirectory
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
            $userOU = Import-Csv -Path .\admin\adminUserOU.csv
            $organizationUnit = ($userOU).ou

            $userDescription = Import-Csv -Path .\admin\adminDescription.csv
            $description = ($userDescription).description
        }

        "a1" 
        {
            $userOU = Import-Csv -Path .\a1\a1UserOU.csv
            $organizationUnit = ($userOU).ou

            $userDescription = Import-Csv -Path .\a1\a1Description.csv
            $description = ($userDescription).description
        }

        "a2" 
        {
            $userOU = Import-Csv -Path .\a2\a2UserOU.csv
            $organizationUnit = ($userOU).ou

            $userDescription = Import-Csv -Path .\a2\a2Description.csv
            $description = ($userDescription).description
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
    New-ADUser -Name $uname -SamAccountName $uname -UserPrincipalName "$uname$domain" `
    -GivenName $fName -Initials $iName -Surname $lName -DisplayName $dName `
    -AccountPassword $secureP -description $description -ProfilePath $userProfile$uName `
    -HomeDrive $userHomeDrive -HomeDirectory $userHomeDirectory$uName -Path "$organizationUnit" `
    -Enabled $true
    

    #After user objects get created they are added to their respective groups\
    switch ($typeList.SelectedItem)
    {
        "general" 
        {
        }

        "admin" 
        {
            $groups = Import-Csv -Path .\admin\adminUserGroups.csv
            foreach ($group in $groups)
            {
                $groupName = ($group).groups
                Add-ADGroupMember -Identity $groupName -Members $uName
            }
        }

        "a1" 
        {
            $groups = Import-Csv -Path .\a1\a1UserGroups.csv
            foreach ($group in $groups)
            {
                $groupName = ($group).groups
                Add-ADGroupMember -Identity $groupName -Members $uName
            }
        }

        "a2" 
        {
            $groups = Import-Csv -Path .\a2\a2UserGroups.csv
            foreach ($group in $groups)
            {
                $groupName = ($group).groups
                Add-ADGroupMember -Identity $groupName -Members $uName
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