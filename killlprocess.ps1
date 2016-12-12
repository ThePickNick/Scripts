$MinTime = 0
$Processname = "prime95"

sleep 5
while ($true){


$CpuCores = (Get-WMIObject Win32_ComputerSystem).NumberOfLogicalProcessors
$Samples = (Get-Counter "\Process($Processname*)\% Processor Time").CounterSamples

$CPU = $Samples.CookedValue / $CpuCores

write-host "Processor % "+ $CPU
Write-Host "Minutes "+ $MinTime

    if(($CPU -ge 30) -And ($MinTime -ge 3)){

        Write-Host "Killing process"
        sleep 2
        Stop-Process -processname $Processname -Force
        break

    }
    if ($CPU -le 30){
        $MinTime = 0
    }
$MinTime++
sleep 5


}

return $CPU