# script to dump users from groups in AD
# Copyright G�ran Thyni, AB SL
# Licenced under General Public Licence version 3
#

$use_excel = $false
if ($use_excel) {
    $excel = New-Object -ComObject excel.application
    $excel.visible = $True
    $workbook = $excel.Workbooks.Add()
    $groupwksht= $workbook.Worksheets.Item(1)
    $groupwksht.Name = "Group Set"
    $userwksht= $workbook.Worksheets.add()
    $userwksht.Name = "User Set"

    $groupwksht.Cells.Item(1,1) = 'Group Policy'
    $groupwksht.Cells.Item(1,2) = 'count'

    $userwksht.Cells.Item(1,1) = 'Group Policy'
    $userwksht.Cells.Item(1,2) = 'Username'
    $userwksht.Cells.Item(1,3) = 'Real name'
    $userwksht.Cells.Item(1,4) = 'Expiration date'
    $userwksht.Cells.Item(1,5) = 'Enabled'
} else {
  write-output "Group Policy;count" | out-file "gc.csv"
  write-output "Group Policy;Username;Real name;Expiration date;Enabled" | out-file "gm.csv" 
}
# exit
$uidx = 2
$gidx = 2
$count = 0
$groupname = "dummy"
$userhash = @{}

$groups = Get-ADGroup -Filter {name -like "VPN_*"} 
$groups |  foreach-object { 
    $count = 0; 
    $groupname = $_.name; 
    #write-output ""; 
    write-output $groupname; 
    $members = Get-ADGroupMember -Recursive $_
    #$members
    ForEach ($i in $members) {
        $props = Get-ADUser -properties SamAccountName,Name,AccountExpirationDate,Enabled,memberof $i 
        if ($userhash.ContainsKey($props.SamAccountName)) { continue }
        $userhash[$props.SamAccountName] = $props.name
        #$props
        $groupname = ($props.memberof | select-string -pattern 'VPN_[^,]+').Matches.Value -join "/"
        #$groupname
        #$props | select SamAccountName, Name, AccountExpirationDate, Enabled
        $date = $props.AccountExpirationDate
        $date = $date -replace ' \d+:\d+:\d+',''
        $date = $date -replace '(\d+)\/(d+)\/(d+)','$3-$1-$2'
        if ($use_excel) {
            $userwksht.Cells.Item($uidx,1) = $groupname
            $userwksht.Cells.Item($uidx,2) = $props.SamAccountName
            $userwksht.Cells.Item($uidx,3) = $props.Name
            $userwksht.Cells.Item($uidx,4) = $date
            $userwksht.Cells.Item($uidx,5) = $props.Enabled
        } else {
            Write-Output "$groupname;$($props.SamAccountName);$($props.Name);$date;$($props.Enabled)" | out-file -Append "gm.csv"
        }
        $uidx++; $count++
    }
    if ($use_excel) { 
        $groupwksht.Cells.Item($gidx,1) = $groupname
        $groupwksht.Cells.Item($gidx,2) = $count
    } else {
        write-output "$groupname;$count" | out-file -Append "gc.csv"
    }
    $gidx++
 }
