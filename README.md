# usermanagement.ps1
Create/delete Active Directory user accounts with no annoying CSV preformatted user lists. Just feed it properly formatted names, enter their OU, and let it do all the work. New valid usernames are generated on-demand. Also support creating new Exchange mailboxes for new user accounts. 

User deletions do not clean up disconnected mailboxes because we do that as a scheduled task on the Exchange server. Also does not clean up roaming profile directories, handle security group membership, or Exchange distribution list membership.

Function 'Choose-ADOrganizationalUnit' credited to Mica H: https://itmicah.wordpress.com/2016/03/29/active-directory-ou-picker-revisited/

## Requirements ##
1. Powershell 3.0 (default on Windows 8 and up).
2. MicaH's Choose-ADOrganizationalUnit function in a dot-sourceable file.
2. Script must run from a PC on the same domain the users are being created in.
3. Script will require domain admin creds. Will prompt for credentials via Get-Credential frame.
4. Remote DC and Exchange servers will need PS-Remoting enabled.

Go through the script and edit the following lines to match your environment:

* <code>Line 16</code>: Path to MicaH's Choose-ADOrganizationalUnit file.
* <code>Line 20, 252</code>: Your default user acount password.
* <code>Line 21, 237</code>: Your email domain.
* <code>Line 128, 290</code>: The FQDN of the Domain Controller you will establish a PS Remote session to.
* <code>Line 214</code>: The FQDN of the Exchange server you will establish a PS Remote session to.
* <code>Line 218</code>: A leading substring of your Mailbox Database(s) name(s).
* <code>Line 329, 352, 401, 424</code>: Your domain and tld LDAP object names.

## Usage ##

The script is interactive and guided, just follow the prompts. When asked to list names for account creation or deletion, feel free to copy and paste in a properly-formatted list.

### Name formatting ###

When entering names, ensure you format thusly: 

<code>Firstname M Last_with_spaces</code>

Examples:
* Pedro L. De La Rosa  => <code>Pedro L De_La_Rosa</code>
* Paul  di Resta => <code>Paul di_Resta</code>

Leading/trailing whitespace is trimmed, and consecutive whitespace characters between name elements are condensed to one space.

Apostraphe (single quote) and hyphen characters are permitted.

### Valid username generation ###

The script will check for username availability to generate valid new usernames. Usernames are generated from the concatenation of Firstname leading characters and the final part of a surname. 

Generation rules:
* Try first initial + surname segment, and if invalid try first two letters of firstname + surname segment, and so on.
* Surnames containing spaces or hyphens will use the final "segment" of the surname.
* Apostraphes are ignored in firstnames and surnames.

Examples:
* <code>Pedro L De_La_Rosa</code> => <code>Prosa</code>
* <code>Paul di_Resta</code> => <code>Presta</code>
* <code>Nico Rosberg</code> => <code>Nrosberg</code>
* <code>Daniel Day-Lewis</code> => <code>Dlewis</code>
* <code>Jenson Button</code> when Jbutton is taken => <code>Jebutton</code>
* <code>D'angelo Russel</code> when Drussel is taken => <code>Darussel</code>

### Repeat name conflicts in user deletion ###

When deleting users, you are first prompted to select the OU the users reside in. The search is recursive, so you may specify higher level OUs, or even the entire domain and search the entire directory. As a result, when listing a common Firstname + Lastname for deletion, the search may return multiple users with the same Firstname + Lastname. When such a conflict is detected, all matching users will be listed and you will be prompted to select the intended deletion target.

To assist you in determining the true intended target, the script will print some identifying information:

```
Multiples of the same name have been found.

Index Name            OU           Account    LastLogon
----- ----            --           -------    ---------
    0 Lewis Hamilton  /Users       Lehamilton 9/2/2016 5:07:46 PM
    1 Lewis Hamilton  /TempTest    Lhamilton  12/31/1600 7:00:00 PM

Enter index or indeces to delete: 
```

OU is the Oganizational Unit where the user object resides, and LastLogon is an [estimate (+/- 1 week)](https://blogs.technet.microsoft.com/askds/2009/04/15/the-lastlogontimestamp-attribute-what-it-was-designed-for-and-how-it-works/) of the last time the user autenticated against a Domain Controller. A LastLogon timestamp of <code>12/31/1600 7:00:00 PM</code> indicates that the user account has never been logged in to.

## To-Do ##

1. Prompt to add security groups for new accounts.
2. Prompt after mailbox creation for distribution list membership.
3. Clean up Roaming Profiles and/or folder redirection locations on user deletion.
