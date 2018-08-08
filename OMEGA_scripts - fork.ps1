#any imports?

#graficky vypis utilizace serveru, potrebuje string jmena serveru

function is_service_running{
    $is_service_running_server = (Read-Host -Prompt 'Please provide SERVER name').Trim()
    Clear-Host
    $is_service_running_service = '*' + ((Read-Host -Prompt 'Please provide SERVICE name').Trim()) + '*'
    Clear-Host
    Get-Service $is_service_running_service -ComputerName $is_service_running_server

    press_F12_to_continue
    menu    
}

function CPU_Utilization
{
    param([string]$var_server_name)
    $current_window_name = $host.ui.RawUI.WindowTitle
    $host.ui.RawUI.WindowTitle = $global_server_variable
    Clear-Host
    Write-Host "You can end monitoring anytime by pressing F12`nFirst results may be delayed based on connection speed and current utilization"
    Start-Sleep -s 5
    while($var_server_name){
        if([console]::KeyAvailable){
            if ([console]::ReadKey().key.ToString().Trim() -eq 'F12'){
                break
            }
        }

        else{
            $server_load = (Get-Counter "\238(*)\6" -ComputerName $var_server_name).readings
            $load_values = $server_load.Split([Environment]::NewLine)
            $var_current_time = Get-Date -UFormat "%d.%m. %H:%M:%S"

            for($i=1;$i -lt $load_values.Length; $i=$i+3){
                $too_long = [int]$load_values[$i].replace(",",".")

                if($i -eq $load_values.Length-3){Write-Host "         Total " -NoNewline}

                if ($too_long -le 10) {Write-Host '|o_________| ' -NoNewline }
                elseif ($too_long -le 20) {Write-Host '|oo________| ' -NoNewline}
                elseif ($too_long -le 30) {Write-Host '|ooo_______| ' -ForegroundColor Cyan -NoNewline}
                elseif ($too_long -le 40) {Write-Host '|oooo______| ' -ForegroundColor Cyan -NoNewline}
                elseif ($too_long -le 50) {Write-Host '|ooooo_____| ' -ForegroundColor Green -NoNewline}
                elseif ($too_long -le 60) {Write-Host '|oooooo____| ' -ForegroundColor Green -NoNewline}
                elseif ($too_long -le 70) {Write-Host '|ooooooo___| ' -ForegroundColor Yellow -NoNewline}
                elseif ($too_long -le 80) {Write-Host '|oooooooo__| ' -ForegroundColor Yellow -NoNewline}
                elseif ($too_long -le 90) {Write-Host '|ooooooooo_| ' -ForegroundColor Red -NoNewline}
                elseif ($too_long -lt 100) {Write-Host '|oooooooooo| ' -ForegroundColor Red -NoNewline}
                else {Write-Host '|___Error__| ' -ForegroundColor Red -NoNewline}
                }

            $var_current_time            
        }
    }
    $host.ui.RawUI.WindowTitle = $current_window_name
    menu
}
<#function change_server_name{
    Clear-Host
    $global_server_variable = (Read-Host -Prompt 'Please provide new server name').Trim()
    menu
}#>

function new_powershell_instance{
    Start-Process powershell -ArgumentList 'OMEGA_scripts.ps1'
    menu
}

function get_server_boot_time{
    $get_server_boot_time_name = (Read-Host -Prompt 'Please provide server name').Trim()
    Clear-Host
    Write-Host ("Last server boot time:`n")
    Write-Host (gcim Win32_OperatingSystem -ComputerName $get_server_boot_time_name).LastBootUpTime | Select-Object Days, Hours, Minutes, Seconds
    press_F12_to_continue
    menu
}
function mailbox_permissions{
    ##########################
    # Bulk permission script #
    # @ Jiri Ondracek        #
    # @ Jakub Kopecky        #
    ##########################
    ###  Powershell setup  ###
    ##########################
    Add-Type -AssemblyName System.Windows.Forms
    $ErrorActionPreference = "stop"
    $error_list = @()
    $mailbox_error_list = @()
    Clear-Host

    #############
    ### Input ###
    #############
    $mailboxlist = @()
    Write-Host("Add mailboxes, one per line, leave empty to end")
    while ($mailbox = Read-Host -prompt "Add Mailbox") {$mailboxlist += $mailbox}
    Clear-Host

    $fullaccesslist = @()
    Write-Host("Add Full Access permission requester, one per line, leave empty to end")
    while ($mailbox = Read-Host -prompt "Add requester") {$fullaccesslist += $mailbox}
    Clear-Host

    $sendaslist = @()
    Write-Host("Add Send As permission requester, one per line, leave empty to end")
    while ($mailbox = Read-Host -prompt "Add requester") {$sendaslist += $mailbox}
    Clear-Host

    $sendonbehalflist = @()
    Write-Host("Add Send on Behalf permission requester, one per line, leave empty to end")
    while ($mailbox = Read-Host -prompt "Add requester") {$sendonbehalflist += $mailbox}
    Clear-Host

    # void check (embrace the void)
    if ($mailboxlist -and ($fullaccesslist -or $sendonbehalflist -or $sendaslist)) {

        ######################
        ### Exchange setup ###
        ######################    
        Write-Host("`n`n`n`n`n`n`n`n***Establishing exchange connection***")
        $SessionOpt = New-PSSessionOption -SkipCACheck:$true -SkipCNCheck:$true -SkipRevocationCheck:$true
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://DEUSEFRAN1505/powershell/ -AllowRedirection -SessionOption $SessionOpt 
        Import-PSSession $Session -AllowClobber
        Clear-Host

        for ($k = 0; $k -lt $mailboxlist.length; $k++) {
            $mailbox = get-Mailbox -identity $mailboxlist[$k] -EA SilentlyContinue
            if ($mailbox -ne $null) {

		
                for ($i = 0; $i -lt $sendonbehalflist.length; $i++) {
                    $user = get-Mailbox -identity $sendonbehalflist[$i] -EA SilentlyContinue
                    if ($user -ne $null) {
                        Set-Mailbox $mailbox.Name.ToString() -GrantSendOnBehalfTo @{Add=$sendonbehalflist[$i]}
                        Write-Host ("*** SEND ON BEHALF OF GRANTED *** for user: " + $sendonbehalflist[$i])  -ForegroundColor Green
                    }
            
                    else {
                        $error_list += $mailboxlist[$k]
                        $error_list += $sendonbehalflist[$i]
                        $error_list += "Send on Behalf"
                    }
                }

                for ($y = 0; $y -lt $fullaccesslist.length; $y++) {
                    $user = get-Mailbox -identity $fullaccesslist[$y] -EA SilentlyContinue
                    if ($user -ne $null) {
                        Add-MailboxPermission -Identity $mailbox.ToString() -User $fullaccesslist[$y] -AccessRights FullAccess -AutoMapping $true
                        Write-Host ("*** FULL ACCESS GRANTED *** for user: " + $fullaccesslist[$y])  -ForegroundColor Green
                    }
            
                    else {
                        $error_list += $mailboxlist[$k]
                        $error_list += $fullaccesslist[$y]
                        $error_list += "Full Mailbox Access"
                    }
                }

                for ($z = 0; $z -lt $sendaslist.length; $z++) {
                    $user = get-Mailbox -identity $sendaslist[$z] -EA SilentlyContinue
                    if ($user -ne $null) {
                        Add-ADPermission -Identity $mailbox.Name.ToString() -User $sendaslist[$z].ToString() -Extendedrights "Send As"
                        Write-Host ("*** SEND AS GRANTED *** for user: " + $sendaslist[$z])  -ForegroundColor Green
                    }
            
                    else {
                        $error_list += $mailboxlist[$k]
                        $error_list += $sendaslist[$z]
                        $error_list += "Send As"
                    }
                }
            }
            else {$mailbox_error_list += $mailboxlist[$k]}
        }

        ########################
        ###  Error reporting ###
        ########################  
        if ($mailbox_error_list.Length -gt 0) {
            Write-Host ("Number of mailboxes not found: " + $mailbox_error_list.Length) -ForegroundColor Red
            if ($such_variable = Read-Host -Prompt 'Would you like to export missing mailboxes? (Leave empty to skip)') {
                $export_list = @()
                foreach ($kekitem in $mailbox_error_list) {
                    $temp_export_object = New-Object -TypeName PSObject
                    $temp_export_object | Add-Member -MemberType NoteProperty -Name "Mailbox not found" -Value $kekitem
                    $export_list += $temp_export_object
                }
                $SaveChooser = New-Object -Typename System.Windows.Forms.SaveFileDialog
                $SaveChooser.Filter = "Csv file (*.csv)|*.csv"
                $SaveChooser.ShowDialog()
                if ($SaveChooser.FileName) {
                    $export_list | Export-Csv -Path $SaveChooser.FileName -Delimiter ";" -NoTypeInformation
                }
                else {
                    Write-Host "***Export was not successful, printing results to terminal***" -ForegroundColor Red
                    Write-Host "***MAILBOXES NOT FOUND***" -ForegroundColor Red
                    $mailbox_error_list
                    Write-Host("")
                }
            }
            else {
                Write-Host "** MAILBOXES NOT FOUND **" -ForegroundColor Red
                $mailbox_error_list
                Write-Host("")
            }
        }
        else {
            Write-Host("***All provided mailboxes were found***") -ForegroundColor Green
        }

        if ($error_list.Length -gt 0) {
            Write-Host("Number of permission errors: " + ($error_list.Length) / 3 ) -ForegroundColor Red
            if ($such_variable = Read-Host -Prompt 'Would you like to export permission errors? (Leave empty to skip)') {
                $export_list = @()
                for ($p = 0; $p -lt $error_list.length; $p = $p + 3) {
                    $temp_export_object = New-Object -TypeName PSObject
                    $temp_export_object | Add-Member -MemberType NoteProperty -Name "Requester" -Value $error_list[$p + 1]
                    $temp_export_object | Add-Member -MemberType NoteProperty -Name "Mailbox" -Value $error_list[$p]
                    $temp_export_object | Add-Member -MemberType NoteProperty -Name "Permission" -Value $error_list[$p + 2]
                    $export_list += $temp_export_object
                }
                $SaveChooser = New-Object -Typename System.Windows.Forms.SaveFileDialog
                $SaveChooser.Filter = "Csv file (*.csv)|*.csv"
                $SaveChooser.ShowDialog()
                if ($SaveChooser.FileName) {
                    $export_list | Export-Csv -Path $SaveChooser.FileName -Delimiter ";" -NoTypeInformation
                }
                else {
                    Write-Host "***Export was not successful, printing results to terminal***" -ForegroundColor Red
                    Write-Host "***FAILED GRANTING PERMISSIONS***" -ForegroundColor Red
                    for ($p = 0; $p -lt $error_list.length; $p = $p + 3) {
                        $where = $error_list[$p]
                        $who = $error_list[$p + 1]
                        $what = $error_list[$p + 2]
                        write-host("Mailbox:	$where `nRequester:	$who `nPermission:	$what`n")
                    }
                }
        
            }
            else {
                Write-Host "***FAILED GRANTING PERMISSIONS***" -ForegroundColor Red
                for ($p = 0; $p -lt $error_list.length; $p = $p + 3) {
                    $where = $error_list[$p]
                    $who = $error_list[$p + 1]
                    $what = $error_list[$p + 2]
                    write-host("Mailbox:	$where `nRequester:	$who `nPermission:	$what`n")
                }
            }
        }
        else {
            Write-Host("***There have been no permission errors***") -ForegroundColor Green
        }
    }
    else {
        Write-Host ("***Either no mailboxes or no permission requesters have been provided***`n***No changes have been done***")
    }

    Remove-PSSession $Session
    press_F12_to_continue
    menu
}

function jabber_client{
    $var_username = Read-Host -Prompt "Username"
    [string]$jabber_display_location_information = "noInput"
    [double]$var_bandwidth = 0
    $var_SID = "none"
    $var_user_info = Get-ADUser -Identity $var_username -Properties 'msRTCSIP-PrimaryUserAddress', Office | Select-Object Name, 'msRTCSIP-PrimaryUserAddress', Office
    $var_name_surname = (($var_user_info.'msRTCSIP-PrimaryUserAddress').Split(('@', ':')))[1]
    Clear-Host

    while ($jabber_display_location_information -ne "y" -and $jabber_display_location_information -ne "n") {
        $jabber_display_location_information = Read-Host -Prompt "Do you want to display location information`nYes: y`nNo:  n`n"
        Clear-Host
        if ($jabber_display_location_information -eq "y") {
            $var_SID = Read-Host -Prompt "Enter SID (4 numbers)"
            $var_shortname = $var_user_info.Office.Split('/')[0] + $var_user_info.Office.Split('/')[1] + $var_SID
            Clear-Host
            $var_bandwidth = Read-Host -Prompt "For Video Bandwidth and Immersive Video Bandwidth value`ncheck CMDB - All circuits and search for Location ID`nEnter bandwidth (Mbps)"
            Clear-Host
            [int]$var_number_of_subnets = Read-Host -Prompt "Search  CMDB - All subnets  for Bussiness and Wireless subnets`nEnter number of Bussiness and Wireless subnets"
            $var_subnet = New-Object -TypeName 'object[]' -ArgumentList $var_number_of_subnets
            $var_subnet_bits = New-Object -TypeName 'object[]' -ArgumentList $var_number_of_subnets
            for ($i = 0; $i -lt $var_number_of_subnets; $i ++) {
                Write-Host (
                    "Subnet " + ($i + 1) + "`n" 
                )
                $var_subnet[$i] = Read-Host -Prompt "Subnet"
                $var_subnet_bits[$i] = Read-Host -Prompt "Subnet Mask (bits size, Example: 24)"
            }
            Clear-Host
            Write-Host (
                "Step 1:`n" +
                'System -> Physical Location -> Add New' + "`n" +
                "Physical Location information:`n" +
                "Name:   " + 'PLI-' + $var_shortname + "`n" +
                "Description:   " + $var_user_info.Office + ' (SID' + $var_SID + ')' + "`n" +
                "`n" +
                "Step 2:`n" +
                'System -> Location Info -> Location -> Add New' + "`n" +
                "Name:   " + 'LOC-' + $var_shortname + '-Jabber' + "`n"
            )
        
            if ($var_bandwidth -ge 1) {
                $var_bandwidth = $var_bandwidth * 2 * 384
            
            }
            else {
                $var_bandwidth = 384
            }
            Write-Host (
                "Video Bandwidth:    " + $var_bandwidth + "`n" +
                "Immersive Video Bandwidth    " + $var_bandwidth + "`n" +
                "`n"
            )
            Write-Host (
                "Step 3:`n" +
                'System -> Device Pool -> New' + "`n" +
                "`n" +
                "Device Pool settings:`n" +
                "Device pool name:    " + 'DP-' + $var_shortname + '-Jabber' + "`n" +
                "Cisco Unified Communications Manager Group:    Default `n" +
                "`n" +
                "Roaming Sensitive Settings:`n" +
                'Date/Time Group:    -select appropriate-' + "`n" +
                'Region:    REG-Video_384kbps' + "`n" +
                "Media Resource Group List:    MRGL CUCM`n" +
                "Location:    " + 'LOC-' + $var_shortname + '-Jabber' + "`n" +
                "Network Locale:    -select appropriate-`n" +
                "Physical Location:    " + 'PLI-' + $var_shortname + "`n" +
                "Device Mobility Group:    EMEA_dmg`n" +
                "`n" +
                "Device Mobility Related Information:`n" + 
                "Device Mobility Calling Search Space:    " + 'CSS-VideoandISND3241Only' + "`n" +
                "`n" +
                "Step 4:`n" +
                'System -> Device Mobility -> Device Mobility Info -> Add New' + "`n" +
                "`n"
            )
            for ($i = 0; $i -lt $var_number_of_subnets; $i ++) {
                Write-Host (
                    "Subnet " + ($i + 1) + "`n" 
                )
                Write-Host (
                    "Device Mobility Info Information:`n" +
                    "Name:    " + 'DMI-' + $var_shortname + '_' + $var_subnet[$i] + "`n" +
                    "Subnet:    " + $var_subnet[$i] + "`n" +
                    "Subnet Mask (bits size):    " + $var_subnet_bits[$i] + "`n" +
                    "`n" +
                    "Device Pools for this Device Mobility Info:`n" + 
                    'DP-' + $var_shortname + '-Jabber' + "`n" +
                    "`n"
                )
            }
        }
    }

    Write-Host (
        "Account settings:`n" +
        'User Management -> User/Phone Add -> Quick User/Phone Add' + "`n" +
        "`n" +
        $var_user_info.Name + "`n" +
        $var_user_info.EmailAddress + "`n" +
        "`n" +
        'Office:    ' + $var_user_info.Office + "`n" +
        "`n" +
        "Extensions:`n" +
        "Line Primary URI/Partition:    " + $var_name_surname + '.jabber@video.heidelbergcement.com' +
        "`nPT-URI `n" +
        "`n" +
        "Personal:`n" + 
        "Directory URI:    " + $var_name_surname + '@video.heidelbergcement.com' + "`n" +
        "`n" +
        "Phone:`n" +
        "Product Type:    Cisco Unified Client Services Framework`n" +
        "Device Protocol:    SIP`n" +
        "Device Name:    CSF" + $var_username.ToUpper() + "`n" +
        "`n" +
        "Resolution message:`n" +
        "Hello,`n" +
        "Cisco Jabber software client account has been created.`n" +
        "In order to login for first time, please follow guide bellow:`n" +
        'http://unite.grouphc.net/uk/it/selfhelp/Self%20Help%20articles%20%20backup/Cisco%20Jabber%20Software%20Client.pdf'
    )
    press_F12_to_continue
    menu
}

function press_F12_to_continue{
    Write-Host ("`n`nPress F12 to continue to menu ...")

    while ($true) {
        if ([console]::KeyAvailable) {
            if ([console]::ReadKey().key.ToString().Trim() -eq 'F12') {
                break
            }
        }
    }  
}
function menu{
    Clear-Host
    Write-Host "Please choose your option by pressing following key`nC ... CPU Utilization`nJ ... Jabber Client`nS ... Server Boot time`nT ... Test running service`nM ... Mailbox permissions`nN ... New PowerShell instance`nQ ... Exit"
    $key_pressed_variable = [System.Console]::ReadKey() 
    switch ($key_pressed_variable.key){
        C {Clear-Host; $global_server_variable = (Read-Host -Prompt 'Please provide server name').Trim(); CPU_Utilization($global_server_variable)}
        N {new_powershell_instance}
        J {Clear-Host;jabber_client}
        S {Clear-Host;get_server_boot_time}
        T {Clear-Host;is_service_running}
        M {Clear-Host;mailbox_permissions}
        Q {Clear-Host;exit}
        default {menu}
    }
} 

Clear-Host
  
menu