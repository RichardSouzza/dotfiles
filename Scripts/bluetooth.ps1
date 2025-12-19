$currentTime = Get-Date
$lunchTime = Get-Date -Hour 13 -Minute 00 -Second 00
$workTime =  Get-Date -Hour 14 -Minute 00 -Second 00
$bluetoothAdapter = "Realtek Bluetooth Adapter"

if ( ($currentTime -ge $lunchTime) -and ($currentTime -lt $workTime) ) {
    Get-PnpDevice -Class Bluetooth -FriendlyName $bluetoothAdapter | Disable-PnpDevice -Confirm:$false
}
else {
    Get-PnpDevice -Class Bluetooth -FriendlyName $bluetoothAdapter | Enable-PnpDevice -Confirm:$false
}
