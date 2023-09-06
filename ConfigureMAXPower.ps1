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



#POWERSHELL GUI
$maxProgressSteps = 100


##Create window that contains all the GUI elements
$form = [Form] @{
  Text = "TRANSFER RATE"; Size = [Size]::new(600, 200); StartPosition = 'CenterScreen'; TopMost = $true; MinimizeBox = $false; MaximizeBox = $false; FormBorderStyle = 'FixedSingle'
}

##Set title and size of the form
###Set title
$form.Text ='Set Processor MAX - ©Jack Green'
###Set width
$form.Width = 800
###Set height
$form.Height = 300

Write-Host ('test')

#Create Progress Bars
$barAC = New-Object Windows.Forms.ProgressBar
$barAC.Dock = [Windows.Forms.DockStyle]::Top

$barDC = New-Object Windows.Forms.ProgressBar
$barDC.Dock = [Windows.Forms.DockStyle]::Bottom

# Create a timer to update the progress bar
$timer = New-Object Windows.Forms.Timer
$timer.Interval = 2000  # Milliseconds (adjust as needed)
$timer.Add_Tick({
    # Call your function to get the progress value (replace with your function)
    $progressValueAC = Get-CurrentPROCTHROTTLEMAX('AC')
    $progressValueDC = Get-CurrentPROCTHROTTLEMAX('DC')

    # Update the progress bar value
    $barAC.Value = $progressValueAC
    $barDC.Value = $progressValueDC


    # Check if the progress is complete and stop the timer if needed
    <#if ($progressValueAC -ge 100) {
        $timer.Stop()
    }#>
})

# Add the progress bar to the form
$form.Controls.Add($barAC)
$form.Controls.Add($barDC)


# Enable the timer when the form loads.
$form.add_Load({
  $timer.Enabled = $true
})







if ($UserInput_ACPercent){
    #Set Processor MAX in % (On AC Power)
    powercfg -setacvalueindex $asGuid $SubProcessor_Guid $procthrottlemax_Guid $UserInput_ACPercent
    }
if ($UserInput_BattPercent){
    #Set Processor MAX in % (On Battery Power)
    powercfg -setdcvalueindex $asGuid $SubProcessor_Guid $procthrottlemax_Guid $UserInput_BattPercent
    }

##automatically stretch the form if the elements on the form are outside the form boundaries
$form.AutoSize = $true
##Display the form on the screen
$form.ShowDialog()

# Getting here means that the form was closed.
# Clean up.
$timer.Dispose(); $form.Dispose()