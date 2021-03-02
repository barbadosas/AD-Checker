Add-Type -AssemblyName System.Windows.Forms

[System.Windows.Forms.Application]::EnableVisualStyles()

#GUI forms

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(412,429)
$Form.text                       = "Quick AD Checker"
$Form.TopMost                    = $false

$compnameTxtbox                  = New-Object system.Windows.Forms.TextBox
$compnameTxtbox.multiline        = $false
$compnameTxtbox.width            = 307
$compnameTxtbox.height           = 20
$compnameTxtbox.location         = New-Object System.Drawing.Point(48,57)
$compnameTxtbox.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Computer name"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(145,26)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$result                          = New-Object system.Windows.Forms.TextBox
$result.multiline                = $true
$result.width                    = 346
$result.height                   = 220
$result.location                 = New-Object System.Drawing.Point(34,148)
$result.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$pcButton                         = New-Object system.Windows.Forms.Button
$pcButton.text                    = "Check PC"
$pcButton.width                   = 102
$pcButton.height                  = 30
$pcButton.location                = New-Object System.Drawing.Point(64,100)
$pcButton.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$closeButton                     = New-Object system.Windows.Forms.Button
$closeButton.text                = "Close"
$closeButton.width               = 102
$closeButton.height              = 30
$closeButton.location            = New-Object System.Drawing.Point(144,385)
$closeButton.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$userButton                   = New-Object system.Windows.Forms.Button
$userButton.text              = "Check User"
$userButton.width             = 102
$userButton.height            = 30
$userButton.location          = New-Object System.Drawing.Point(227,99)
$userButton.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form.controls.AddRange(@($compnameTxtbox,$Label1,$result,$pcButton,$closeButton,$userButton))

$pcButton.Add_Click({ check_pc })
$userButton.Add_Click({ check_user })
$closeButton.Add_Click({ closeForm })

#Function for Check PC Button
function check_pc(){

    $result.text = ''
    $PC = $compnameTxtbox.text
    $search = [adsisearcher]"(&(ObjectCategory=Computer)(Name=$PC))"
    $users = $search.FindAll()

    foreach($user in $users) {

        $DN = $user.Properties['DisplayName']
        $WIN = $user.Properties['OperatingSystemVersion']
        $OC = $user.Properties['ObjectClass']
        $MB = $user.Properties['ManagedBy']
        $OU = $user.Properties['distinguishedname']
        $LL = $user.Properties['lastlogontimestamp']

        #Converts OS build to version

        if ($WIN -eq "10.0 (18363)"){
            $WIN = '1909'
        }

        elseif ($WIN -eq '10.0 (17134)'){
            $WIN = '1803'}

        else {$WIN = $WIN + '(Not 1909)'
        }
    }

    #check if primary user is set

     if ($null -ne $MB){

        $MB_split = $MB.split(",") | Select-String -Pattern 'CN'
        $MB_str = $MB_split.ToString()
        $MB_trimmed = $MB_str.Substring(3)
        }

    else {$MB_trimmed = $MB + 'Not Set'}

    #split OU
    $OU2 = $OU.split(",") | Select-String -Pattern 'OU'

    #convert unix time to date time

    $LLT = (Get-Date '1601-01-01').AddDays([long]::parse($LL)*100/86400/1000/1000/1000)


  $result.text += "DisplayName:    $DN"
  $result.text += "`r`nLast Logon:    $LLT"
  $result.text += "`r`nWindows Version:    $WIN"
  $result.text += "`r`nPrimary User:    $MB_trimmed"
  $result.text += "`r`nMachine OU:    $OU2"
  $result.text += "`r`nObjectClass:    $OC"
}
#Function for Check User Button
function check_user(){

    $result.text = ''
    $USER = $compnameTxtbox.text
    $search = [adsisearcher]"(&(ObjectCategory=Person)(ObjectClass=User)(cn=$USER))"
    $users = $search.FindAll()

        #Get user AD objects
    foreach($user in $users) {
        $DisplayName = $user.Properties['DisplayName']
        $DistinguishedName = $user.Properties['DistinguishedName']
        $EmailAddress = $user.Properties['mail']
        $HomeDirectory = $user.Properties['HomeDirectory']
        $managedObjects = $user.Properties['managedObjects']
        $department = $user.Properties['extensionattribute7']
        $lastlogon = $user.Properties['lastlogon']

        $Managed_split = $managedObjects.split(",") | Select-String -Pattern 'CN'
        $Managed_split_trimmed = ($Managed_split -join " ")
        $Managed_split_cleaned = ($Managed_split_trimmed.ToString().Split())

        $DN_split = $DistinguishedName.split(",") | Select-String -NotMatch 'DC'

        #unix time to date time
        $last_logon_time = (Get-Date '1601-01-01').AddDays([long]::parse($lastlogon)*100/86400/1000/1000/1000)


        $result.text += "DisplayName:    $DisplayName"
        $result.text += "`r`nDepartment:    $department"
        $result.text += "`r`nLast Logon:    $last_logon_time"
        $result.text += "`r`nDistinguishedName:    $DN_split"
        $result.text += "`r`nEmailAddress:    $EmailAddress"
        $result.text += "`r`nHomeDirectory:    $HomeDirectory"
        $result.text += "`r`nManaged Objects:    $Managed_split" # if new line per object is not needed
        $result.text += "`r`nManaged Objects Ordered:"

    try{
        $ErrorActionPreference = 'Stop'

    foreach ($mo in $Managed_split_cleaned){
        $mo2 = $mo.Substring(3)
        $result.text += "`r`n$mo2"}
        }
    catch { "Error occured" }
        }

}

function closeForm(){$Form.close()}
[void]$Form.ShowDialog()