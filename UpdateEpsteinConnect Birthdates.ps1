
$csv = import-csv -Path C:\temp\dates.csv 

foreach($user in $csv){
    $username = $user.username
    $hiredate1 = $user.hiredate1
    $hiredatefinal = get-date $hiredate1
    $dateofbirth1 = $user.birthdate1
    $dateofbirthfinal = get-date $dateofbirth1
    Set-ADUser $username -replace @{dateofhire1=$hiredatefinal;dateofbirth1=$dateofbirthfinal}
    Get-ADUser -identity $username -Properties name,dateofhire1,dateofbirth1 | select name,dateofhire1,dateofbirth1
}




 