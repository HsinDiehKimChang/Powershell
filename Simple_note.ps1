###Written by Hsin-Dieh (Kim) Chang
###2019/06/18
###Version 1


$i=1

do
{
    
Write-Host "$i ticket(s):"

$INC = Read-Host "Please input INC number"
$supName = Read-Host "Please input support team name"

Write-Host "BPPM alert : please help to check $INC Thanks"

$Primary = Read-Host "Please input Primary name"
Write-Host "paged $supName ($Primary) for $INC / BPPM"

$stamp=get-date
Write-Host "$stamp paged $supName ($Primary)"

$input = Read-Host "Type q to quit or type any key to continue"
$i++

}until ($input -eq 'q')
