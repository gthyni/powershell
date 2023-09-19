# script to dump users from groups in AD
# Copyright G�ran Thyni, AB SL
# Licenced under General Public Licence version 3
#
$PWD=(pwd).Path
$gc = New-Object System.IO.StreamWriter("$PWD\gc-dump.csv", $true)
$gm = New-Object System.IO.StreamWriter("$PWD\gm-dump.csv", $true)

# Write to the file
$gc.WriteLine("Group Policy;count")
$gm.WriteLine("Group Policy;Username;Real name;Expiration date;Enabled;Mobiltfn;Email;Maildomain;Chef/bestallare")


$uidx = 2
$gidx = 2
$count = 
$groupname = "dummy"
$userhash = @{}

$groups = Get-ADGroup -Filter {name -like "AC-TF-*"} -SearchBase 'OU=CGI SG MFA,OU=Grupper,DC=nobel,DC=sl,DC=se'
$groups |  foreach-object { 
    $count = 0; 
    $groupname = $_.name; 
    write-output $groupname; 
    $members = Get-ADGroupMember -Recursive $_
    #$members
    ForEach ($i in $members) {
        $props = Get-ADUser -properties SamAccountName,Name,AccountExpirationDate,Enabled,memberof,mobilephone,emailaddress,manager,othermailbox $i 
        if ($userhash.ContainsKey($props.SamAccountName)) { continue }
        if ($props.EmailAddress -notmatch '@') { $props.EmailAddress = $props.othermailbox[0] }
        $userhash[$props.SamAccountName] = $props.name
        #$props
        #$groupname = ($props.memberof | select-string -pattern 'VPN_[^,]+').Matches.Value -join "/"
        $newgroup = ($props.memberof | select-string -pattern 'AC-TF-[^,]+').Matches.Value -join "/"
        #$groupname
        #$props | select SamAccountName, Name, AccountExpirationDate, Enabled
        $date = $props.AccountExpirationDate
        $date = $date -replace ' \d+:\d+:\d+',''
        $date = $date -replace '(\d+)\/(d+)\/(d+)','$3-$1-$2'
        $mdomain = $props.emailaddress -replace '.+@',''
        #Add-Content -Path ".\gm.csv" -Value 
        $gm.WriteLine("$newgroup;$($props.SamAccountName);$($props.Name);$date;$($props.Enabled);$($props.mobilephone);$($props.emailaddress);$mdomain;$($props.manager)")
        $uidx++; $count++
    }
    $gc.WriteLine("$groupname;$count") 
    $gidx++
 }
# Close the streams when you're done writing to the file
$gc.Close()
$gm.Close()
