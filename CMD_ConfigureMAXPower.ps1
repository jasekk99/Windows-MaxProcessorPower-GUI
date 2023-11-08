using namespace System.Windows.Forms
using namespace System.Drawing

Add-Type -AssemblyName System.Windows.Forms


$activeScheme = cmd /c "powercfg /getactivescheme"
$regEx = '(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}'
$asGuid = [regex]::Match($activeScheme,$regEx).Value

$SubProcessor_Guid = '54533251-82be-4824-96c1-47b60b740d00'
$procthrottlemax_Guid = 'bc5038f7-23e0-4960-96da-33abaf5935ec'

$currentPROCTHROTTLEMAXValue_AC = $null

#Function for getting current PROCTHROTTLEMAX values
function Get-CurrentPROCTHROTTLEMAX([string]$mode) {
    Switch($mode){
        'AC'{
            $currentPROCTHROTTLEMAXValue = powercfg /query $asGuid $SubProcessor_Guid $procthrottlemax_Guid

            $currentPROCTHROTTLEMAXValue_AC = $currentPROCTHROTTLEMAXValue -like '*Index der aktuellen Wechselstromeinstellung:*'
            $currentPROCTHROTTLEMAXValue_AC = $currentPROCTHROTTLEMAXValue_AC.Split(':')
            $currentPROCTHROTTLEMAXValue_AC = $currentPROCTHROTTLEMAXValue_AC[1]
            $currentPROCTHROTTLEMAXValue_AC = $currentPROCTHROTTLEMAXValue_AC.Trim()
            $currentPROCTHROTTLEMAXValue_AC = [System.Convert]::ToInt64($currentPROCTHROTTLEMAXValue_AC, 16)
            $currentPROCTHROTTLEMAXValue_AC
        }
        'DC'{
            $currentPROCTHROTTLEMAXValue = powercfg /query $asGuid $SubProcessor_Guid $procthrottlemax_Guid

            $currentPROCTHROTTLEMAXValue_DC = $currentPROCTHROTTLEMAXValue -like '*Index der aktuellen Gleichstromeinstellung:*'
            $currentPROCTHROTTLEMAXValue_DC = $currentPROCTHROTTLEMAXValue_DC.Split(':')
            $currentPROCTHROTTLEMAXValue_DC = $currentPROCTHROTTLEMAXValue_DC[1]
            $currentPROCTHROTTLEMAXValue_DC = $currentPROCTHROTTLEMAXValue_DC.Trim()
            $currentPROCTHROTTLEMAXValue_DC = [System.Convert]::ToInt64($currentPROCTHROTTLEMAXValue_DC, 16)
            $currentPROCTHROTTLEMAXValue_DC
        }
        default{
            $currentPROCTHROTTLEMAXValue = powercfg /query $asGuid $SubProcessor_Guid $procthrottlemax_Guid

            $currentPROCTHROTTLEMAXValue_AC = $currentPROCTHROTTLEMAXValue -like '*Index der aktuellen Wechselstromeinstellung:*'
            $currentPROCTHROTTLEMAXValue_AC = $currentPROCTHROTTLEMAXValue_AC.Split(':')
            $currentPROCTHROTTLEMAXValue_AC = $currentPROCTHROTTLEMAXValue_AC[1]
            $currentPROCTHROTTLEMAXValue_AC = $currentPROCTHROTTLEMAXValue_AC.Trim()
            $global:currentPROCTHROTTLEMAXValue_AC = [System.Convert]::ToInt64($currentPROCTHROTTLEMAXValue_AC, 16)

            $currentPROCTHROTTLEMAXValue_DC = $currentPROCTHROTTLEMAXValue -like '*Index der aktuellen Gleichstromeinstellung:*'
            $currentPROCTHROTTLEMAXValue_DC = $currentPROCTHROTTLEMAXValue_DC.Split(':')
            $currentPROCTHROTTLEMAXValue_DC = $currentPROCTHROTTLEMAXValue_DC[1]
            $currentPROCTHROTTLEMAXValue_DC = $currentPROCTHROTTLEMAXValue_DC.Trim()
            $global:currentPROCTHROTTLEMAXValue_DC = [System.Convert]::ToInt64($currentPROCTHROTTLEMAXValue_DC, 16)
                
        }
    }
}

$currentPROCTHROTTLEMAXValue_AC = Get-CurrentPROCTHROTTLEMAX('AC')
$currentPROCTHROTTLEMAXValue_DC = Get-CurrentPROCTHROTTLEMAX('DC')

switch ($currentPROCTHROTTLEMAXValue_AC) {
    {$_ -ge 1 -and $_ -le 25} {
        $currPowerColor_AC = "Red"
    }
    {$_ -ge 26 -and $_ -le 50}{
        $currPowerColor_AC = "Orange"
    }
    {$_ -ge 51 -and $_ -le 75}{
        $currPowerColor_AC = "Yellow"
    }
    {$_ -ge 76 -and $_ -le 100}{
        $currPowerColor_AC = "Green"
    }
    default{
        $currPowerColor_AC = "Green"
    }
}

switch ($currentPROCTHROTTLEMAXValue_DC) {
    {$_ -ge 1 -and $_ -le 25} {
        $currPowerColor_DC = "Red"
    }
    {$_ -ge 26 -and $_ -le 50}{
        $currPowerColor_DC = "Orange"
    }
    {$_ -ge 51 -and $_ -le 75}{
        $currPowerColor_DC = "Yellow"
    }
    {$_ -ge 76 -and $_ -le 100}{
        $currPowerColor_DC = "Green"
    }
    default{
        $currPowerColor_DC = "Green"
    }
}

Write-Host "Current CPU power settings:" -ForegroundColor DarkGray


Write-Host "    Power (AC):     " -ForegroundColor DarkGray -NoNewline
Write-Host $currentPROCTHROTTLEMAXValue_AC -ForegroundColor $currPowerColor_AC

Write-Host "    Battery (DC):   " -ForegroundColor DarkGray -NoNewline
Write-Host $currentPROCTHROTTLEMAXValue_DC -ForegroundColor $currPowerColor_DC


do{
$BATorAC = Read-Host "Do you want to change power on Battery or AC? [A]C / [B]attery"
    if(-not ($BATorAC -match '^[A-Ba-b]$')){
        Write-Host "Invalid letter, please input either A or B" -BackgroundColor red
        [Console]::ResetColor()
    }
} while (-not ($BATorAC -match '^[A-Ba-b]$'))

do{
$UserInput_Percent = Read-Host "Power %?"
    if(-not ($UserInput_Percent -match '^\d+$' -and $UserInput_Percent -ge 1 -and $UserInput_Percent -le 100)){
        Write-Host "Invalid number, please input a number in range 1-100!" -BackgroundColor red
        [Console]::ResetColor()
    }
} while (-not ($UserInput_Percent -match '^\d+$' -and $UserInput_Percent -ge 1 -and $UserInput_Percent -le 100))

switch ($BATorAC)
    {
        "b" {powercfg -setdcvalueindex $asGuid $SubProcessor_Guid $procthrottlemax_Guid $UserInput_Percent}
        "B" {powercfg -setdcvalueindex $asGuid $SubProcessor_Guid $procthrottlemax_Guid $UserInput_Percent}
        "a" {powercfg -setacvalueindex $asGuid $SubProcessor_Guid $procthrottlemax_Guid $UserInput_Percent}
        "A" {powercfg -setacvalueindex $asGuid $SubProcessor_Guid $procthrottlemax_Guid $UserInput_Percent}
    }