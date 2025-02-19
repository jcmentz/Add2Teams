﻿<# 
Script created by Dr JC Mentz with great effort and enthusiasm
The interface was created using POSHGUI.com a free online gui designer for PowerShell
Usage Rights - use it , fiddel with it BUT do no tmake me responsible for the outcome!

ON the the other hand, if the script works for you, send me an email at mentzjc@unisa.ac.za

Important:
the input file MUST only contain student numbers, the script will add the email address to the student number.

The result of adding students is written in a file with the name 'result' on your desktop, look for a teamoutput folder.
When you add students to a team make sure that you first select the team ID before you select the input file.

I tried to create error trapping but this is minimal and in need of further development. The string handling is also clunky
and in need of refinement.

Change Log:
30 January 2020 - added @mylife.unisa.ac.za in script instead of expecting from file. Input file now only need to be a 
list of student numbers. NB! make sure that there are no spaces after the student number and that each student number
is on its own line.

4 February 2020 - refined error file processing. The idea is to create a file with potential license issues.

19 February 2020 - add functionaloty to add members to private teams. Dependancy - MicrosoftTeams 1.0.20 must be installed 
#>

<# version 6.1 interface

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '493,302'
$Form.text                       = "Teams Management Console"
$Form.TopMost                    = $false

$Btn_Exit                        = New-Object system.Windows.Forms.Button
$Btn_Exit.text                   = "EXIT"
$Btn_Exit.width                  = 155
$Btn_Exit.height                 = 30
$Btn_Exit.location               = New-Object System.Drawing.Point(324,49)
$Btn_Exit.Font                   = 'Microsoft Sans Serif,10'
$Btn_Exit.Enabled                = $false

$Btn_TeamsConnect                = New-Object system.Windows.Forms.Button
$Btn_TeamsConnect.text           = "Connect"
$Btn_TeamsConnect.width          = 155
$Btn_TeamsConnect.height         = 30
$Btn_TeamsConnect.location       = New-Object System.Drawing.Point(6,14)
$Btn_TeamsConnect.Font           = 'Microsoft Sans Serif,10'
$Btn_TeamsConnect.Enabled        = $true

$Btn_Load_TMS                    = New-Object system.Windows.Forms.Button
$Btn_Load_TMS.text               = "Load Teams"
$Btn_Load_TMS.width              = 155
$Btn_Load_TMS.height             = 30
$Btn_Load_TMS.location           = New-Object System.Drawing.Point(165,14)
$Btn_Load_TMS.Font               = 'Microsoft Sans Serif,10'
$Btn_Load_TMS.Enabled            = $false

$DataGridView1                   = New-Object system.Windows.Forms.DataGridView
$DataGridView1.width             = 472
$DataGridView1.height            = 202
$DataGridView1.location          = New-Object System.Drawing.Point(6,87)
$DataGridView1.MultiSelect       = $false
$DataGridView1.columns.Add("Teams_ID", "Team ID")
$DataGridView1.columns.Add("Teams_Name", "Team Name")
$DataGridView1.ColumnHeadersVisible = $true

$Btn_ShowSelection               = New-Object system.Windows.Forms.Button
$Btn_ShowSelection.text          = "Add Users"
$Btn_ShowSelection.width         = 155
$Btn_ShowSelection.height        = 30
$Btn_ShowSelection.location      = New-Object System.Drawing.Point(5,48)
$Btn_ShowSelection.Font          = 'Microsoft Sans Serif,10'
$Btn_ShowSelection.Enabled       = $false

$Btn_GetUsers                    = New-Object system.Windows.Forms.Button
$Btn_GetUsers.text               = "Get User List"
$Btn_GetUsers.width              = 155
$Btn_GetUsers.height             = 30
$Btn_GetUsers.location           = New-Object System.Drawing.Point(165,49)
$Btn_GetUsers.Font               = 'Microsoft Sans Serif,10'
$Btn_GetUsers.Enabled            = $false

$Btn_Disconnect                  = New-Object system.Windows.Forms.Button
$Btn_Disconnect.text             = "Disconnect"
$Btn_Disconnect.width            = 155
$Btn_Disconnect.height           = 30
$Btn_Disconnect.location         = New-Object System.Drawing.Point(324,14)
$Btn_Disconnect.Font             = 'Microsoft Sans Serif,10'
$Btn_Disconnect.Enabled          = $false

$Form.controls.AddRange(@($Btn_Exit,$Btn_TeamsConnect,$Btn_Load_TMS,$DataGridView1,$Btn_ShowSelection,$Btn_GetUsers,$Btn_Disconnect))
#>

<# version 6.3 interface#>
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '560,302'
$Form.text                       = "Teams Management Console"
$Form.TopMost                    = $false

$Btn_Exit                        = New-Object system.Windows.Forms.Button
$Btn_Exit.text                   = "EXIT"
$Btn_Exit.width                  = 155
$Btn_Exit.height                 = 30
$Btn_Exit.location               = New-Object System.Drawing.Point(324,49)
$Btn_Exit.Font                   = 'Microsoft Sans Serif,10'
$Btn_Exit.Enabled                = $false

$Btn_TeamsConnect                = New-Object system.Windows.Forms.Button
$Btn_TeamsConnect.text           = "Connect"
$Btn_TeamsConnect.width          = 155
$Btn_TeamsConnect.height         = 30
$Btn_TeamsConnect.location       = New-Object System.Drawing.Point(6,14)
$Btn_TeamsConnect.Font           = 'Microsoft Sans Serif,10'
$Btn_TeamsConnect.Enabled        = $true

$Btn_Load_TMS                    = New-Object system.Windows.Forms.Button
$Btn_Load_TMS.text               = "Load Teams"
$Btn_Load_TMS.width              = 155
$Btn_Load_TMS.height             = 30
$Btn_Load_TMS.location           = New-Object System.Drawing.Point(165,14)
$Btn_Load_TMS.Font               = 'Microsoft Sans Serif,10'
$Btn_Load_TMS.Enabled            = $false

$DataGridView1                   = New-Object system.Windows.Forms.DataGridView
$DataGridView1.width             = 334
$DataGridView1.height            = 202
$DataGridView1.location          = New-Object System.Drawing.Point(6,87)
$DataGridView1.MultiSelect       = $false
$DataGridView1.columns.Add("Teams_ID", "Team ID")
$DataGridView1.columns.Add("Teams_Name", "Team Name")
$DataGridView1.ColumnHeadersVisible = $true

$Btn_ShowSelection               = New-Object system.Windows.Forms.Button
$Btn_ShowSelection.text          = "Add To Team"
$Btn_ShowSelection.width         = 155
$Btn_ShowSelection.height        = 30
$Btn_ShowSelection.location      = New-Object System.Drawing.Point(354,126)
$Btn_ShowSelection.Font          = 'Microsoft Sans Serif,10'
$Btn_ShowSelection.enabled       = $false

$Btn_GetUsers                    = New-Object system.Windows.Forms.Button
$Btn_GetUsers.text               = "Get User List"
$Btn_GetUsers.width              = 155
$Btn_GetUsers.height             = 30
$Btn_GetUsers.location           = New-Object System.Drawing.Point(165,49)
$Btn_GetUsers.Font               = 'Microsoft Sans Serif,10'
$Btn_GetUsers.Enabled            = $false

$Btn_Disconnect                  = New-Object system.Windows.Forms.Button
$Btn_Disconnect.text             = "Disconnect"
$Btn_Disconnect.width            = 155
$Btn_Disconnect.height           = 30
$Btn_Disconnect.location         = New-Object System.Drawing.Point(324,14)
$Btn_Disconnect.Font             = 'Microsoft Sans Serif,10'
$Btn_Disconnect.Enabled          = $false

$Btn_AddPrivate                  = New-Object system.Windows.Forms.Button
$Btn_AddPrivate.text             = "Add To Private Channel"
$Btn_AddPrivate.width            = 155
$Btn_AddPrivate.height           = 30
$Btn_AddPrivate.location         = New-Object System.Drawing.Point(354,201)
$Btn_AddPrivate.Font             = 'Microsoft Sans Serif,10'
$Btn_AddPrivate.Enabled          = $false

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Select a TEAM on the left"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(357,93)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$CBPrivateChannels               = New-Object system.Windows.Forms.ComboBox
$CBPrivateChannels.width         = 100
$CBPrivateChannels.height        = 20
$CBPrivateChannels.location      = New-Object System.Drawing.Point(354,165)
$CBPrivateChannels.Font          = 'Microsoft Sans Serif,10'


$Form.controls.AddRange(@($Btn_Exit,$Btn_TeamsConnect,$Btn_Load_TMS,$DataGridView1,$Btn_ShowSelection,$Btn_GetUsers,$Btn_Disconnect,$Btn_AddPrivate,$Label1,$CBPrivateChannels))

#$Btn_Exit.Add_Click({ $Form.close() })
#$Btn_TeamsConnect.Add_Click({  })
#$Btn_Load_TMS.Add_Click({  })
#$Btn_ShowSelection.Add_Click({  })
#$Btn_Disconnect.Add_Click({  })
#$Btn_GetUsers.Add_Click({  })

<# end INterface#>
$Credentials = $null
$file_name = $null
$selection = $null

#some housekeeping
#create output folder on desktop. Consider creating try{} catch{} error trapping
$result_path = [Environment]::GetFolderPath('Desktop') + '\teamoutput'
if(!(Test-Path -Path $result_path)){
 New-Item -Path $result_path -ItemType Directory
}
else{
 Write-Host $result_path + " already created"
}

#closing connection and program
$Btn_Exit.Add_Click({
    Disconnect-MicrosoftTeams
    $DataGridView1.rows.Clear()
    $Form.close()   
})

#Connecting to TEAMS
$Btn_TeamsConnect.Add_Click({ 
    try{
        $Credentials = Get-Credential -Message "Enter your TEAMS Username and Password" -ErrorAction Stop
        Connect-MicrosoftTeams -Credential $Credentials -ErrorAction Stop
        $Form.text = "Logged in as " + $Credentials.UserName.ToString()
        $Btn_Disconnect.Enabled = $true
        $Btn_Exit.Enabled = $true
        $Btn_ShowSelection.Enabled = $false
        $Btn_Load_TMS.Enabled = $true
        $Btn_AddPrivate.Enabled = $false
        
    }
    catch
    {
        [System.Windows.MessageBox]::Show($error[0])
        $Form.close()
    }
    #finally{}
})

#display list of teams associated with logged in user
$Btn_Load_TMS.Add_Click({ 


   try{
      $team = Get-Team -User $Credentials.UserName -ErrorAction Stop
      $DataGridView1.rows.Clear()
      ForEach ($member in $team){ 
        $group_ID = $member.GroupId    
        $group_Name = $member.DisplayName
        $DataGridView1.rows.Add($Group_ID,$Group_Name)
      }      
    }
    catch{
      [System.Windows.MessageBox]::Show($error[0])
      $Form.close()
    }
    #finally{}
  
   
  
  

    $Btn_ShowSelection.Enabled = $true
    $Btn_GetUsers.Enabled = $true
    $Btn_AddPrivate.Enabled = $true

})

#load users to selected team
$Btn_ShowSelection.Add_Click({ 
    
   $selection = $DataGridView1.SelectedCells[0].FormattedValue.ToString()  
   $result_path = [Environment]::GetFolderPath('Desktop') + '\teamoutput'
   $result_filename = $selection + '_result.txt'
   $result_backup = $selection + (get-date -Format "_MM-dd-yyyy-HHmm_").ToString() + "Backup.txt"
   $path_backup = $result_path + '\' + $result_backup
   $path_filename = $result_path + '\' + $result_filename
   
   #test if file already exists, if so, rename with date time stamp
   if(Test-Path -Path $path_filename){
     #create backup by renaming
     Rename-Item -Path $path_filename -NewName $path_backup
     Write-Host "backup created"
     Write-Host "Creating new file"
     New-Item -path $path_filename
     Set-Content -Path $path_filename 'error messages'
   }
   else{
     Write-Host "Creating new file"
     New-Item -path $path_filename
     Set-Content -Path $path_filename 'error messages'
   }
  
   #open classlist file at preset location
   $Dlg_Open_file = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
   $null = $Dlg_Open_file.ShowDialog()
   $file_name = $Dlg_Open_file.FileName
   
   $text = Get-Content -Path $file_name
   $text.GetType() | Format-Table -AutoSize
   $size = $text.Length
   Write-Host "array size = " $size
      ForEach ($member in $text){ 
        try{
           $StudentToAdd = $text[$size-1] +'@mylife.unisa.ac.za'
           Add-TeamUser -GroupId $selection -User $StudentToAdd -ErrorAction Continue
           write-host $size -> $text[$size-1] " added to team " $selection
           $size = $size - 1;
        }
        catch{
          #write error to file, two types of error, 1) no license 2) user already exists
          $temp_message = $error[0].ToString().split([environment]::NewLine)
          Add-Content -path $path_filename -value $temp_message[2]
          write-host $size -> $temp_message[2]
          $size = $size - 1;
          $continue
        }
        #finally{}
      }      
 })

 $CBPrivateChannels.Add_Click({
 
  $selection = $DataGridView1.SelectedCells[0].FormattedValue.ToString()
  
   #get private channel names and display in Combobox
   try{
      $PChannels = Get-TeamChannel -GroupId $selection -MembershipType Private -ErrorAction Stop
      $CBPrivateChannels.Items.Clear()
      $CBPrivateChannels.Items.AddRange($PChannels.DisplayName)
      $Btn_AddPrivate.Enabled = $true      
    }
    catch{
      #[System.Windows.MessageBox]::Show($error[0])
      Write-Host "NO private channels in this team!"
      $Btn_AddPrivate.Enabled = $false
      $break      
    }
    #finally{}  
  })

 #add to private channels
 $Btn_AddPrivate.Add_Click({
   $selection = $DataGridView1.SelectedCells[0].FormattedValue.ToString()  
   $result_path = [Environment]::GetFolderPath('Desktop') + '\teamoutput'
   $result_filename = $selection + 'PChAdd_result.txt'
   $result_backup = $selection + (get-date -Format "_MM-dd-yyyy-HHmm_").ToString() + "PChAdd_Backup.txt"
   $path_backup = $result_path + '\' + $result_backup
   $path_filename = $result_path + '\' + $result_filename
  

   #test if file already exists, if so, rename with date time stamp
   if(Test-Path -Path $path_filename){
     #create backup by renaming
     Rename-Item -Path $path_filename -NewName $path_backup
     Write-Host "backup created"
     Write-Host "Creating new file"
     New-Item -path $path_filename
     Set-Content -Path $path_filename 'error messages'
   }
   else{
     Write-Host "Creating new file"
     New-Item -path $path_filename
     Set-Content -Path $path_filename 'error messages'
   }
  
   #open private channel list file at preset location
   $Dlg_Open_file = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
   $null = $Dlg_Open_file.ShowDialog()
   $file_name = $Dlg_Open_file.FileName
   
   $text = Get-Content -Path $file_name
   $text.GetType() | Format-Table -AutoSize
   $size = $text.Length
   Write-Host "array size = " $size
      ForEach ($member in $text){ 
        try{
           $StudentToAdd = $text[$size-1] +'@mylife.unisa.ac.za'
           #Add-TeamUser -GroupId $selection -User $StudentToAdd -ErrorAction Continue
           Add-TeamChannelUser -GroupId $selection -DisplayName $CBPrivateChannels.Text -User $StudentToAdd -ErrorAction Continue
           write-host $size -> $text[$size-1] " added to private channel " $CBPrivateChannels.Text
           $size = $size - 1;
        }
        catch{
          #write error to file, two types of error, 1) no license 2) user already exists
          $temp_message = $error[0].ToString().split([environment]::NewLine)
          Add-Content -path $path_filename -value $StudentToAdd->$temp_message[2] 
          write-host $size " :" $StudentToAdd -> $temp_message[2]
          $size = $size - 1;
          $continue
        }
        #finally{}
      }      
  })

#disconnect from TEAMS
$Btn_Disconnect.Add_Click({ 
    Disconnect-MicrosoftTeams
    $DataGridView1.rows.Clear()
    $form.text = "Logged out! "
    write-host $credentials.UserName " logged out!"
    $Btn_Disconnect.Enabled = $false
    $Btn_Exit.Enabled = $true
    $Btn_ShowSelection.Enabled = $false
    $Btn_Load_TMS.Enabled = $false
    [System.Windows.MessageBox]::Show("You are now disconnected from TEAMS")
})

$Btn_GetUsers.Add_Click({ 
      
   $selection = $DataGridView1.SelectedCells[0].FormattedValue.ToString() 
   $result_path = [Environment]::GetFolderPath('Desktop') + '\teamoutput'
   $result_filename = $selection + '_TeamList.txt'
   $path_filename = $result_path + '\' + $result_filename
   $date = get-date -Format "MM-dd-yyyy-HHmm"
   $result_backup = $selection + $date + 'Backup.txt'

   #test if file already exists, if so, rename with date time stamp
   Write-Host (Test-Path -Path $path_filename).ToString()
   if(Test-Path -Path $path_filename){
     #create backup by renaming
     Rename-Item -Path $path_filename -NewName $result_backup
     Write-Host "backup created"
     $result = Get-TeamUser -GroupId $selection -Role Member
     Add-Content -Path $path_filename -Value $result.User
   }
   else{
     Write-Host "Creating new file"
     $result = Get-TeamUser -GroupId $selection -Role Member
     Add-Content -Path $path_filename -Value $result.User
   }

    
})



$Form.ShowDialog()