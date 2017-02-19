<# User Management Script
 Antonio Abella
 6/24/2016

 Create/delete Active Directory user accounts with no annoying CSV preformatted user lists.
 Just feed it properly formatted names, enter their OU, and let it do all the work. Also support
 creating new Exchange mailboxes for new user accounts. User deletions do no clean up
 disconnected mailboxes because we do that as a scheduled task on the Exchange server.

 Function 'Choose-ADOrganizationalUnit' credited to Mica H
 https://itmicah.wordpress.com/2016/03/29/active-directory-ou-picker-revisited/
#>

# Dot-source file containing Choose-ADOrganizationalUnit
# function.
. .\path\to\ChooseADOrganizationalUnit.ps1

# Hold default password and domain
# Edit these to suit your environment
$defaultpw = (ConvertTo-SecureString "password-here" -AsPlainText -force)
$emaildomain = '@domain.tld'

clear

# Prompt for add or delete
$action = Read-Host -Prompt "`nWould you like to [A]dd or [D]elete users?"

# If adding
if (([string]::Compare($action, 'a', $True) -eq 0) -or ([string]::Compare($action, 'add', $True) -eq 0)) {

    # defining a local clipboard holder for later
    $cliphold = ""

    # Select your OU to bulk-create users for
    Write-Host "Choose an OU to create users for.`n"
    $ou = Choose-ADOrganizationalUnit

    Write-Host "==============================="
    Write-Host "Bulk user creation for $($ou.Name)"
    Write-Host "===============================`n`n"

    # Prompt for names
    Write-Host "Enter names in the format [FirstName] [M] [LastName]. You may"
    Write-Host "choose to omit the middle name. If the first or last names contain"
    Write-Host "a space, use an underscore ( _ ) to represent it.`n"
    Write-Host "When finished, leave blank and press Enter.`n`n"
    $nameArray = @()
    $entry = "temp"

    # Add names to array as they are entered if not empty string
    while ($entry -ne "") {
        $entry = Read-Host -Prompt 'Enter a name in format [First] [M] [Last]'

        if ($entry -ne "") {
            $strip = $entry.trim() -replace '\s+', ' '
            $nameArray += $strip
        }
    }

    clear

    # Print names and indices
    Write-Host "`n`nPlease check the entered names for misspellings or"
    Write-Host "mistaken entries. If any are found, write their"
    Write-Host "index numbers seperated by spaces and press Enter.`n"
    Write-Host "If no misspellings are found, leave blank and press Enter.`n"

    $nameArray | % {$index=0} {
        Write-Host "    $index  $_"
        $index++
    }

    # Loop until all misspellings are taken care of
    while ($true) {

        # Prompt for misspelled indices
        # Use space-delimited string to indicated bad indices
        # Tokenize index string and iterate over resulting array to get bad names and locations to insert fixed names
        Write-Host ""
        $misspells = Read-Host "Misspelled indexes or [Enter]"

        if ($misspells -ne "") {
            Write-Host "`nEnter corrected spelling for each index, or type [delete] to delete.`n"
            $indexlist = $misspells -split " "

            $indexlist | % {

                # Allow entry deletions
                $currentname = $nameArray[$_]
                $fixedname = Read-Host -Prompt "Enter the corrected spelling of [$currentname] or [delete]"
                if ($fixedname -eq "delete"){
                    $nameArray = $nameArray -ne $currentname
                } elseif ($fixedname -ne "") {
                    $nameArray[$_] = $fixedname
                }
            }
        } else {
            clear
            break
        }
        Write-Host "`nCorrected Names:`n"
        $nameArray | % {$index=0} {
            Write-Host "    $index  $_"
            $index++
        }
        $correct = Read-Host -Prompt "`nIs this correct? [Y/n]"
        if (($correct -eq "Y") -or ($correct -eq "y") -or ($correct -eq "")) {
            clear
            break
        }
        clear
        Write-Host "`nCurrent Names:`n"
        $nameArray | % {$index=0} {
            Write-Host "    $index  $_"
            $index++
        }
    }
    # Automated username generation
    # Username convention:
    #   IF AVAILABLE: first initial concatinated with lowercase surname
    #   ELSE:         first two letters of first name concatinated with lowercase surname
    #   ELSE:         first three letters of first name... and so on and so forth
    Write-Host "`nCreating Usernames...`n`n"

    # Establish remote session to domain controller
    # DC must have PowerShell remoting enabled
    $cred = Get-Credential -Credential $null
    $s = New-PSSession -ComputerName domaincontroller.domain.tld -Credential $cred
    Invoke-Command -Session $s -Scriptblock {
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue

        # PS function [string]::IsNullOrWhitespace is not in PS version
        # on our DC, so we make it ourselves
        function StringIsNullOrWhitespace([string] $string) {
            if ($string -ne $null) {
                $string = $string.Trim()
            }
            return [string]::IsNullOrEmpty($string)
        }
        $usrarray = @()
        $userobjects = @()

        # Grab name elements (first, mi, last)
        # split first name to char array, iterate over it
        $using:nameArray | % { $i = 0 }{
            $username = ""
            $splitname = $_ -split " "
            $firstname = $splitname[0] -replace "_",' '
            $minitial = ""
            $lastname = $splitname[-1] -replace "_",' '

            if ($splitname.Count -gt 2) {
                $minitial = $splitname[1]
            }

            # for every char in array starting at index 0, concatinate it to last name
            # check against Get-ADUser to see if username already exists
            # if it does, add next letter in char array and check again
            $fninitarray = @()
            $fnarray = $firstname.ToCharArray()
            foreach ($letter in $fnarray) {
                if ($letter -ne "'") { $fninitarray += ,$letter }
                $fninit = -join $fninitarray
                $username = -join ($fninit, $($lastname.ToLower().split(" -")[-1] -replace "'",''))
                $namecheck = $(Get-ADUser -Filter "samAccountName -like '$username'")
                if (StringIsNullOrWhitespace($namecheck)){
                    $usrarray += ,$username
                    break
                }
            }

            # Create user objects to hold user properties and add to user array
            $userprops = @{FirstName=$firstname;
                           MiddleInit=$minitial;
                           LastName=$lastname;
                           Username=$username}
            $user = New-Object psobject -Property $userprops
            $userobjects += ,$user
            $i++
        }

        # Print names and Usernames
        Write-Host ($userobjects | Format-Table FirstName,LastName,MiddleInit,Username -auto | Out-String)
        $changepw = Read-Host -Prompt "Change password on first login? [y/N]"

        # Create the users
        $continue = Read-Host -Prompt "`nPress [Enter] to create the accounts, or type [Stop] to abort"
        if (([string]::Compare($continue, 'stop', $True) -eq 0) -or ([string]::Compare($continue, 's', $True) -eq 0)) {
            break
        } elseif ($continue -eq "") {
            clear
            Write-Host "Writing users to Active Directory...`n"
            foreach ($userobj in $userobjects) {
                $name = $userobj.FirstName + " " + $userobj.LastName
                if ($userobj.MiddleInit -ne "") {
                    $displayname =  $userobj.FirstName + " " + $userobj.MiddleInit + " " + $userobj.LastName
                    } else {
                        $displayname = $name
                    }
                if (([string]::Compare($changepw, 'n', $True) -eq 0) -or ([string]::Compare($changepw, 'no', $True) -eq 0) -or ($changepw -eq "")) {
                    New-ADUser -AccountPassword $using:defaultpw -ChangePasswordAtLogon $False -PasswordNeverExpires $True -DisplayName $displayname -givenName $userobj.FirstName -surName $userobj.LastName -Initials $userobj.MiddleInit -Enabled $true -Name $displayname -samAccountName $userobj.Username -UserPrincipalName "$($userobj.Username)$using:emaildomain" -Path $using:ou.DistinguishedName
                } elseif (([string]::Compare($changepw, 'y', $True) -eq 0) -or ([string]::Compare($changepw, 'yes', $True) -eq 0)) {
                    New-ADUser -AccountPassword $using:defaultpw -ChangePasswordAtLogon $True -PasswordNeverExpires $False -DisplayName $displayname -givenName $userobj.FirstName -surName $userobj.LastName -Initials $userobj.MiddleInit -Enabled $true -Name $displayname -samAccountName $userobj.Username -UserPrincipalName "$($userobj.Username)$using:emaildomain" -Path $using:ou.DistinguishedName
                }
                Write-Host "Created user $($userobj.Username) - $($userobj.FirstName) $($userobj.LastName)"
            }
        }

        # Prompt for mailbox creation
        $mbox = Read-Host "`n`nDo you want to create mailboxes for these accounts? [y/N]"
        if (([string]::Compare($mbox, 'y', $True) -eq 0) -or ([string]::Compare($action, 'yes', $True) -eq 0)){

            # Edit remote host with your Exchange server. Must have PowerShell sessions with Exchange Management Console forwarding enabled
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchangeserver.domain.tld/PowerShell/ -Authentication Kerberos
            $dbs = Invoke-Command -Session $Session -Scriptblock {
                Get-MailboxDatabase -Status
            }
            $ndbs = $dbs | Where-Object {$_.Name -match 'EMDB'} | Select Name,AvailableNewMailboxSpace
            $mostspace = 0
            $theone = ""
            $spacegb = @()
            $ndbs | % {
                $spacehold = $_.AvailableNewMailboxSpace -split ' '
                if ($spacehold[1] -ne "MB" -and $spacehold[1] -ne "KB") {
                    $spacegb += @(,($_.Name,([double]$spacehold[0]*1024)))
                }
            }
            $spacegb | % {
                if ($($_[1]) -gt $mostspace) {
                    $mostspace = $_[1]
                    $theone = $_[0]
                }
            }
            foreach ($userobj in $userobjects) {

                # Edit with your domain name
                $mailuser = $userobj.Username+"@domain.tld"
                Invoke-Command -Session $Session -Scriptblock {
                    Enable-Mailbox $($args[0]) -Database $($args[1])
                } -argumentlist $mailuser,$theone
            }

            Write-Host "`n`n`nMailboxes have been created."

            Remove-PSSession $Session
        }

        Write-Host "`n"
        $userobjects | Format-Table FirstName,LastName,MiddleInit,Username -auto | Out-String -outvariable clippy

        # Edit with your default new account password
        $clippy += "Passwords are all [default password here]."
        Write-Host "This table has been copied to your clipboard. Please paste it into the ticket resolution."
        Write-Host "Select a monospace font like Courier New for proper formatting."

    }

    $cliphold = Invoke-Command -Session $s -Scriptblock { $clippy }
    $cliphold | clip

    Remove-PSSession $s

# If deleting
} elseif (([string]::Compare($action, 'd', $True) -eq 0) -or ([string]::Compare($action, 'delete', $True) -eq 0) -or ([string]::Compare($action, 'del', $True) -eq 0)){

    # Pick your OU
    Write-Host "Choose an OU to delete users from.`n"
    $ou = Choose-ADOrganizationalUnit

    Write-Host "=================================="
    Write-Host "Bulk user deletion for $($ou.Name)"
    Write-Host "==================================`n`n"

    # Prompt for accounts to delete
    $deletearray = @()
    while ($true) {
        $todelete = Read-Host -Prompt "Enter a user to delete in format [Firstname] [Lastname]"
        if ($todelete -ne "") {
            $strip = $todelete -replace '\s+', ' '
            $deletearray += $strip
        }
        else {
            break
        }
    }
    clear
    try {
        # Establish remote session to domain controller
        # DC must have PowerShell remoting enabled
        $s = New-PSSession -ComputerName domaincontroller.domain.tld -Credential $(Get-Credential -Credential $null)
        Invoke-Command -Session $s -Scriptblock {
            Import-Module ActiveDirectory -ErrorAction SilentlyContinue

            function StringIsNotNullOrWhitespace([string] $string) {
                if ($string -ne $null) {
                    $string = $string.Trim()
                }
                return ![string]::IsNullOrEmpty($string)
            }

            # Self explanatory
            function Get-ADUserLastLogon([string]$userName) {
                $time = 0
                $user = Get-ADUser $userName | Get-ADObject -Properties lastLogontimeStamp
                if ($user.LastLogontimeStamp -gt $time) {
                    $time = $user.LastLogontimeStamp
                }
                $dt = [DateTime]::FromFileTime($time)
                return $dt
            }

            # Parse names and pull associated accounts
            $newuserarray = @()
            $badarray = @()
            $using:deletearray | % {$i = 0} {
                $currentname = $_
                $splitname = $_ -split " "
                $user = $(Get-ADUser -Filter "(givenName -like '$($splitname[0])') -and (surName -like '$($splitname[1])')" -SearchBase $using:ou.DistinguishedName) #| Select Name,GivenName,Surname,DistinguishedName,samAccountName)
                if (StringIsNotNullOrWhitespace($user)) {
                    if ($user.Count -gt 0) {

                        # Prompt for clarification if multiples of same name are found.
                        # Ex. John Doe in FLVS and John C Doe in Client Services
                        Write-Host "`n`nMultiples of the same name have been found."
                        $repeats = @()
                        $user | % { $j=0 } {

                            # Edit below DC= entries with your domain and tld
                            $dname = $_.DistinguishedName.replace("CN=$($_.givenName) $($_.Surname),OU=","").replace(",OU=",",").replace(',DC=domain',"").replace(",DC=tld","")
                            $dnamearray = $dname -split ","
                            $dname = "/$($dnamearray[$dnamearray.Count..0] -join '/')"
                            $ll = Get-ADUserLastLogon($_.samAccountName)
                            $userprops = @{Index=$j;Name=$_.Name;OU=$dname;Account=$_.samAccountName;LastLogon=$ll}
                            $newuser = New-Object psobject -Property $userprops
                            $repeats += ,$newuser
                            $j++
                        }
                        Write-Host "`n" ($repeats | Format-Table Index,Name,OU,Account,LastLogon -auto | Out-String)
                        Write-Host "Enter the index or indices of users to delete. If selecting multiple"
                        Write-Host "indices, write them in a space-seperated list.`n"
                        $delindex = Read-Host -Prompt "Index or indeces to delete"
                        $delindex = $delindex -split " "
                        foreach ($index in $delindex) {
                            $user = $user | Where-Object { $_ -like (Get-ADUser -Filter "samAccountName -eq '$($repeats[$index].Account)'") }
                        }
                    }

                    # Prepare found user objects for display
                    foreach ($u in $user) {

                        # Edit below DC= entries with your domain and tld
                        $dname = $u.DistinguishedName.replace("CN=$($u.givenName) $($u.Surname),OU=","").replace(",OU=",",").replace(',DC=domain',"").replace(",DC=tld","")
                        $dnamearray = $dname -split ","
                        $dname = "/$($dnamearray[$dnamearray.Count..0] -join '/')"
                        $ll = Get-ADUserLastLogon($u.samAccountName)
                        $userprops = @{Index=$i;Name=$u.Name;OU=$dname;Account=$u.samAccountName;LastLogon=$ll}
                        $newuser = New-Object psobject -Property $userprops
                        $newuserarray += ,$newuser
                        $i++
                    }
                } else {
                    $badarray += ,$currentname
                }
            }

            # If any names don't have an account associated (perhaps a mispelling), print them and prompt to edit.
            if ($badarray -ne "") {
                clear
                Write-Host "The following users could not be found:`n"
                $badarray | % {$i=0}{
                    Write-Host "`t$i $_"
                    $i++
                }

                # Add corrected names to user array.
                $fixednames = @()
                Write-Host "`nTo correct an above name, enter a space-seperated list of indices to edit."
                Write-Host "Otherwise, leave blank to discard names.`n"
                $delindex = Read-Host -Prompt "Index or indeces to edit"
                if ($delindex -ne "") {
                    $delindex = $delindex -split " "
                    foreach ($index in $delindex) {
                        $newname = Read-Host -Prompt "Enter the corrected spelling of [$($badarray[$index])]"
                        $fixednames += ,$newname
                    }
                }
                clear

                # Check corrected names for multiples again
                $fixednames | % {$i = $newuserarray.Count} {
                    $currentname = $_
                    $splitname = $_ -split " "
                    $user = $(Get-ADUser -Filter "(givenName -like '$($splitname[0])') -and (surName -like '$($splitname[1])')" -SearchBase $using:ou.DistinguishedName)
                    if (StringIsNotNullOrWhitespace($user)) {
                        if ($user.Count -gt 0) {
                            Write-Host "`n`nMultiples of the same name have been found."
                            $repeats = @()
                            $user | % { $j=0 } {

                                # Edit below DC= entries with your domain and tld
                                $dname = $_.DistinguishedName.replace("CN=$($_.givenName) $($_.Surname),OU=","").replace(",OU=",",").replace(',DC=domain',"").replace(",DC=tld","")
                                $dnamearray = $dname -split ","
                                $dname = "/$($dnamearray[$dnamearray.Count..0] -join '/')"
                                $ll = Get-ADUserLastLogon($_.samAccountName)
                                $userprops = @{Index=$j;Name=$_.Name;OU=$dname;Account=$_.samAccountName;LastLogon=$ll}
                                $newuser = New-Object psobject -Property $userprops
                                $repeats += ,$newuser
                                $j++
                            }
                            Write-Host "`n" ($repeats | Format-Table Index,Name,OU,Account,LastLogon -auto | Out-String)
                            Write-Host "Enter the index or indices of users to delete. If selecting multiple"
                            Write-Host "indices, write them in a space-seperated list.`n"
                            $delindex = Read-Host -Prompt "Index or indeces to delete"
                            if ($delindex -ne "") {
                                $delindex = $delindex -split " "
                                foreach ($index in $delindex) {
                                    $user = $user | Where-Object { $_ -like (Get-ADUser -Filter "samAccountName -eq '$($repeats[$index].Account)'") }
                                }
                            }
                        }
                        foreach ($u in $user) {

                            # Edit below DC= entries with your domain and tld
                            $dname = $u.DistinguishedName.replace("CN=$($u.givenName) $($u.Surname),OU=","").replace(",OU=",",").replace(',DC=domain',"").replace(",DC=tld","")
                            $dnamearray = $dname -split ","
                            $dname = "/$($dnamearray[$dnamearray.Count..0] -join '/')"
                            $ll = Get-ADUserLastLogon($u.samAccountName)
                            $userprops = @{Index=$i;Name=$u.Name;OU=$dname;Account=$u.samAccountName;LastLogon=$ll}
                            $newuser = New-Object psobject -Property $userprops
                            $newuserarray += ,$newuser
                            $i++
                        }
                    } else {
                        $badarray += ,$currentname
                    }
                }
            }

            # CYA warning
            clear
            Write-Host "`n READ CAREFULLY!!"
            Write-Host " ================`n"
            Write-Host "`nThe following users are pending deletion. Examine the list thoroughly for"
            Write-Host "false matches or misspellings. If any are found, provide their indices to"
            Write-Host "remove them from the queue.`n"
            Read-Host -Prompt "Press Enter to continue"

            # Repeatedly prompt to remove a user from the deletion queue until list is confirmed.
            # Don't try hold my script accountable if you fuck up.
            while ($true) {
                Write-Host ($newuserarray | Format-Table Index,Name,OU,Account,LastLogon -auto | Out-String) "`n`n"
                Write-Host "To remove items from the deletion queue, enter their indices as a space-seperated list."
                Write-Host "To proceed with deletions, leave blank."  -backgroundcolor yellow -foregroundcolor black
                $ohshitno = Read-Host -Prompt "`nIndices to remove from the deletion queue"
                if ($ohshitno -ne "") {
                   $ohshitno = $ohshitno -split " "
                    foreach ($index in $ohshitno) {
                        $newuserarray = $newuserarray | Where-Object { $_.Index -ne $index }
                 }
                } else {
                    clear
                 break
                }
            }
            Write-Host "The following users will now be deleted:"
            Write-Host ($newuserarray | Format-Table Index,Name,OU,Account,LastLogon -auto | Out-String) "`n`n"
            Write-Host "This action cannot be reversed!!" -backgroundcolor yellow -foregroundcolor black
            $sureursure = Read-Host -Prompt "Press [Enter] to delete listed users, or type [Stop] to abort"

            if (([string]::Compare($sureursure, 'stop', $True) -eq 0) -or ([string]::Compare($sureursure, 's', $True) -eq 0)){
                Write-Host "`nThe session has been aborted."
                break
            }
            # Delete that shit
            Write-Host "`n"
            $newuserarray | % {
                Get-ADUser -Filter "samAccountName -eq '$($_.Account)'" | Remove-ADUser -Confirm:$false
                Write-Host "User $($_.Name) has been deleted."
            }
        }
    }
    finally {
        Remove-PSSession $s
    }
}
Read-Host -Prompt "`n`nPress enter to exit"
clear
