start microsoft-edge:https://www.tickcounter.com/countdown/4381721/expiration-date

$wshell = New-Object -ComObject wscript.shell;

$wshell.AppActivate('edge') 

while ($true)
{
    Sleep 1
    $wshell.SendKeys('{f11}')
    exit 
}

exit