$wshell = New-Object -ComObject wscript.shell;

$wshell.AppActivate('chrome.exe')

$count= 0
sleep 2


while($count -ne 10000){
Sleep 2
$wshell.SendKeys('2')

$count++

Write-Output $count
}
