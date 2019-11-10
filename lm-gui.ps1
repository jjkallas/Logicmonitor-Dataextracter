<#
LM Extraction GUI 
by jack kallas

Command-line: LM-Export-Data deviceName *startTime *stopTime *fileName

creates output file: deviceName.csv or fileName.csv

#>
####################################### START OF BACKEND

Function LM-Create-Devlist{ # creates dated list of devices named: mm-dd-yyyy-devices.txt
    $list = LM-Get-All-Devices
    $date = Get-Date -Format "MM-dd-yyyy"
    $txtFile = $date + "-devices.txt"
    
    if ( (Test-Path $txtFile) -eq $true ){
        rm $txtFile   # delete .txt if already created today
    }

    $header = "Date: " + $date + " | Total devices: " + $list.total
    Write-Output $header | Out-File $txtFile

    $i = 1
    Foreach ( $dev in $list.items ){
        $line = [String]$i + ". " + $dev.displayName
        Write-Output $line | Out-File $txtFile -Append
        $i = $i + 1
    }
}

Function LM-Get-Datasources{
    [Cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage=
        "Please enter a device ID! :)")]
        $deviceId,
        [Parameter(Mandatory=$False)]
        $queryString = '?size=100'    #100 datasources returned, max is 1000
    )

    $accessId = ''
    $accessKey = ''
       
    <# Use TLS 1.2 #>
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    <# request details #>
    $httpVerb = 'GET'
    $resourcePath = '/device/devices/' + $deviceId + '/devicedatasources' 
   
    $url = 'https://bhsg.logicmonitor.com/santaba/rest' + $resourcePath + $queryString

    <# Get current time in milliseconds #>
    $epoch = [Math]::Round((New-TimeSpan -start (Get-Date -Date "1/1/1970") -end (Get-Date).ToUniversalTime()).TotalMilliseconds)

    <# Concatenate Request Details #>
    $requestVars = $httpVerb + $epoch + $resourcePath

    <# Construct Signature #>
    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    $hmac.Key = [Text.Encoding]::UTF8.GetBytes($accessKey)
    $signatureBytes = $hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($requestVars))
    $signatureHex = [System.BitConverter]::ToString($signatureBytes) -replace '-'
    $signature = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($signatureHex.ToLower()))

    <# Construct Headers #>
    $auth = 'LMv1 ' + $accessId + ':' + $signature + ':' + $epoch
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization",$auth)
    $headers.Add("Content-Type",'application/json')

    <# Make Request #>
    $response = Invoke-RestMethod -Uri $url -Method Get -Header $headers 

    <# Print status and body of response #>
    $status = $response.status
    $body = $response.data

    $body
}

Function LM-Get-All-Devices{
    [Cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        $queryString = '?size=500'  # request 500 devices, but there should be much less (~200)
    )

    $accessId = ''
    $accessKey = ''
       
    <# Use TLS 1.2 #>
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    <# request details #>
    $httpVerb = 'GET'
    $resourcePath = '/device/devices'

    $url = 'https://bhsg.logicmonitor.com/santaba/rest' + $resourcePath + $queryString

    <# Get current time in milliseconds #>
    $epoch = [Math]::Round((New-TimeSpan -start (Get-Date -Date "1/1/1970") -end (Get-Date).ToUniversalTime()).TotalMilliseconds)

    <# Concatenate Request Details #>
    $requestVars = $httpVerb + $epoch + $resourcePath

    <# Construct Signature #>
    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    $hmac.Key = [Text.Encoding]::UTF8.GetBytes($accessKey)
    $signatureBytes = $hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($requestVars))
    $signatureHex = [System.BitConverter]::ToString($signatureBytes) -replace '-'
    $signature = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($signatureHex.ToLower()))

    <# Construct Headers #>
    $auth = 'LMv1 ' + $accessId + ':' + $signature + ':' + $epoch
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization",$auth)
    $headers.Add("Content-Type",'application/json')

    <# Make Request #>
    $response = Invoke-RestMethod -Uri $url -Method Get -Header $headers 

    <# Print status and body of response #>
    $status = $response.status
    $body = $response.data

    $body
}

Function LM-Get-Device{
    [Cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage=
        "Please enter a device name! :)")]
        $deviceName,
        [Parameter(Mandatory=$False)]
        $queryParam = ''
    )

    $queryString = '?filter=displayName:' + $deviceName
    $queryString = $queryString + $queryParam

    $accessId = ''
    $accessKey = ''
       
    <# Use TLS 1.2 #>
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    <# request details #>
    $httpVerb = 'GET'
    $resourcePath = '/device/devices/' 
    $url = 'https://bhsg.logicmonitor.com/santaba/rest' + $resourcePath + $queryString

    <# Get current time in milliseconds #>
    $epoch = [Math]::Round((New-TimeSpan -start (Get-Date -Date "1/1/1970") -end (Get-Date).ToUniversalTime()).TotalMilliseconds)

    <# Concatenate Request Details #>
    $requestVars = $httpVerb + $epoch + $resourcePath

    <# Construct Signature #>
    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    $hmac.Key = [Text.Encoding]::UTF8.GetBytes($accessKey)
    $signatureBytes = $hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($requestVars))
    $signatureHex = [System.BitConverter]::ToString($signatureBytes) -replace '-'
    $signature = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($signatureHex.ToLower()))

    <# Construct Headers #>
    $auth = 'LMv1 ' + $accessId + ':' + $signature + ':' + $epoch
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization",$auth)
    $headers.Add("Content-Type",'application/json')

    <# Make Request #>
    $response = Invoke-RestMethod -Uri $url -Method Get -Header $headers 

    <# Print status and body of response #>
    $status = $response.status
    $body = $response.data 
    
    $body
}

Function LM-Get-Data{
    [Cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage=
        "Please enter a device ID! :)")]
        $deviceId,
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage=
        "Please enter a datasource ID! :)")]
        $dataSourceId,
        [Parameter(Mandatory=$False)]
        $queryString = ''
    )

    $accessId = ''
    $accessKey = ''
       
    <# Use TLS 1.2 #>
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    <# request details #>
    $httpVerb = 'GET'
    $resourcePath = '/device/devices/' + $deviceId + '/devicedatasources/' + $dataSourceId + '/data' 
    $url = 'https://bhsg.logicmonitor.com/santaba/rest' + $resourcePath + $queryString

    <# Get current time in milliseconds #>
    $epoch = [Math]::Round((New-TimeSpan -start (Get-Date -Date "1/1/1970") -end (Get-Date).ToUniversalTime()).TotalMilliseconds)

    <# Concatenate Request Details #>
    $requestVars = $httpVerb + $epoch + $resourcePath

    <# Construct Signature #>
    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    $hmac.Key = [Text.Encoding]::UTF8.GetBytes($accessKey)
    $signatureBytes = $hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($requestVars))
    $signatureHex = [System.BitConverter]::ToString($signatureBytes) -replace '-'
    $signature = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($signatureHex.ToLower()))

    <# Construct Headers #>
    $auth = 'LMv1 ' + $accessId + ':' + $signature + ':' + $epoch
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization",$auth)
    $headers.Add("Content-Type",'application/json')

    <# Make Request, retry in 60 secs if failed (usually throttled) #>
    $stoploop = $false
    do {
        try {
            $response = Invoke-RestMethod -Uri $url -Method Get -Header $headers 
            $stoploop = $true
            }
        catch {
            Write-Host "Request exceeded rate limit, retrying in 60 seconds..."
            Start-Sleep -Seconds 60
		}
	}
    While ($stoploop -eq $false)
            
    $response

}

Function LM-Get-All-Data{
    [Cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage=
        "Please enter a device ID! :)")]
        $deviceId,
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage=
        "Please enter a datasource ID! :)")]
        $dataSourceId,
        [Parameter(Mandatory=$True)]
        $start,
        [Parameter(Mandatory=$True)]
        $end
    )

    For ($i=$start; $i -le $end; $i = $i + 60000){        # 60000 epoch difference is 500 results (the result limit per API request), if dataqueries are 2 minutes apart
        $query = '?format=csv&start=' + $i + '&end=' + ($i + 60000)
        LM-Get-Data $deviceId $dataSourceId $query
    }

}

Function LM-Export-Data{
    [Cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage=
        "Please enter a device name! :)")]
        $deviceName,
        [Parameter(Mandatory=$False)]
        $start = '',
        [Parameter(Mandatory=$False)]
        $stop = '',
        [Parameter(Mandatory=$False)]
        $queryString = '?format=csv',
        [Parameter(Mandatory=$False)]
        $fileName = ''
    )

    $deviceInfo = LM-Get-Device $deviceName

    $deviceId = $deviceInfo.items.id
    $deviceCreatedOn = $deviceInfo.items.createdOn
    $deviceUpdatedOn = $deviceInfo.items.updatedOn

    #  set default time range and filename if needed
    if ( $start -eq '' ){
        $start = $deviceCreatedOn
        $stop = $deviceUpdatedOn
    }
    if ( $fileName -eq '' ){
        $fileName = $deviceName + '.csv'
    }
    else {
        $fileName = $fileName + '.csv'
    }

    $deviceDatasources = LM-Get-Datasources $deviceId

    Write-Host $start -ForegroundColor DarkYellow
    Write-Host $stop -ForegroundColor DarkYellow

    $percentDone = 0

    Write-Progress "Generating LM extract for $deviceName" -PercentComplete 0

    Foreach ($source in $deviceDatasources.items){ 
        LM-Get-All-Data $deviceId $source.id $start $stop | 
            ConvertFrom-Csv -Header Datasource, Timestamp, Val0, Val1, Val2, Val3, Val4, Val5, Val6, Val7, Val8, Val9, Val10,
             Val11, Val12, Val13, Val14, Val15, Val16, Val17, Val18, Val19, Val20, Val21, Val22, Val23, Val24, Val25, Val26, Val27, Val28, Val29, Val30 |
            Export-Csv $fileName -Append -NoTypeInformation

        $percentDone++

        Write-Progress "Generating LM extract for $deviceName" -PercentComplete ( ($percentDone * 100) / $deviceDatasources.total )`
            -Status ("current source: " + $source.dataSourceName)

        Start-Sleep -m 1000  #1 second delay to avoid API throttle (set to 2000 if hitting request limits)
    }
    Write-Progress "Done" -Completed

}

####################################### END OF BACKEND

####################################### START OF FRONTEND

$gui = New-Object System.Windows.Forms.Form
$gui.Text = 'BH Data Extractor'
$gui.Height = 200
$gui.Width = 500
$gui.AutoScale = $true
$gui.StartPosition = "CenterScreen"
$gui.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")

#### Background images START

$img = [System.Drawing.Image]::FromFile( 'C:\beaconHill1800.jpg' )
$picBox = New-Object Windows.Forms.PictureBox
$picBox.Width = $img.Size.Width
$picBox.Height = $img.Size.Height
$picBox.Image = $img

$bhLogo = [System.Drawing.Image]::FromFile( 'C:\bhLogo.png' )
$bhBox = New-Object Windows.Forms.PictureBox
$bhBox.Width = $bhLogo.Size.Width
$bhBox.Height = $bhLogo.Size.Height
$bhBox.Image = $bhLogo

#### Background images END

#### Input section START

$findB = New-Object System.Windows.Forms.Button
$findB.Location = New-Object System.Drawing.Point( ((($gui.Width / 2) + $gui.Width / 4) - 130), (($gui.Height / 2) ) )
$findB.Text = "Find Device"
$findB.Add_Click(
{
    Write-host "looking..."
    $selectedDevice = LM-Get-Device ($deviceBox.Text)
    $startI.Text = ($selectedDevice.items[0].createdOn)
    $stopI.Text = ($selectedDevice.items[0].updatedOn)
}
)

$extractB = New-Object System.Windows.Forms.Button
$extractB.Location = New-Object System.Drawing.Point( ((($gui.Width / 2) + $gui.Width / 4) - 50), (($gui.Height / 2) ) )
$extractB.Text = "Extract"
$extractB.Add_Click(
{
    $startVal = [int] $startI.Text
    $stopVal = [int] $stopI.Text
    if ($stopVal -gt $startVal ) {
        Try {
            LM-Export-Data $deviceBox.Text $startVal $stopVal -fileName $fileI.Text
            Write-Host "Extract completed." 
        }
        Catch {
            Write-Host "Received error: " -NoNewline -ForegroundColor White
            Write-Error $_
            Write-Host "Check device name."
        }

    } 
    else {
        if ( $stopVal -eq '' ) {
            Write-Host "Please enter a time range"
        }
        else {
            Write-Host "Invalid dates: Stop date must be after start date"
        }
    }
}
)

#### Device Input Start

$deviceL = New-Object System.Windows.Forms.Label
$deviceL.Location = New-Object System.Drawing.Point( ((($gui.Width / 2) + $gui.Width / 4) - 136), (($gui.Height / 2) - 57 ) )
$deviceL.Text = "Device:"
$deviceL.AutoSize = $true

$deviceBox = New-Object System.Windows.Forms.ComboBox
$deviceBox.Width = 180
$deviceBox.Location = New-Object System.Drawing.Point( ((($gui.Width / 2) + $gui.Width / 4) - 90), (($gui.Height / 2) - 60 ) )
$deviceBox.Sorted = $true
$deviceBox.Add_Textchanged(
{
    $fileI.Text = $deviceBox.Text
}
)

$deviceList = LM-Get-All-Devices
Foreach ($device in $deviceList.items){ 
    $deviceBox.Items.Add( $device.displayName ) | Write-Debug #redirect output to debug to clear the spam
} 

#### Device Input End

$fileL = New-Object System.Windows.Forms.Label
$fileL.Location = New-Object System.Drawing.Point( ((($gui.Width / 2) + $gui.Width / 4) - 155), (($gui.Height / 2) - 32 ) )
$fileL.Text = "Output csv:"
$fileL.AutoSize = $true

$fileI = New-Object System.Windows.Forms.TextBox
$fileI.Location = New-Object System.Drawing.Point( ((($gui.Width / 2) + $gui.Width / 4) - 90), (($gui.Height / 2) - 35 ) )
$fileI.Text = $deviceBox.Text
$fileI.Width = 165

#### Input section END

#### Epochs section START

$epochsL = New-Object System.Windows.Forms.Label
$epochsL.Location = New-Object System.Drawing.Point( (($gui.Width / 4) - 35), (($gui.Height / 2) - 70 ) )
$epochsL.Text = "Epochs" 
$epochsL.AutoSize = $true

$startL = New-Object System.Windows.Forms.Label
$startL.Location = New-Object System.Drawing.Point( (($gui.Width / 4) - 75), (($gui.Height / 2) - 47 ) )
$startL.Text = "Start:"
$startL.AutoSize = $true

$startI = New-Object System.Windows.Forms.TextBox
$startI.Location = New-Object System.Drawing.Point( (($gui.Width / 4) - 45), (($gui.Height / 2) - 50 ) )
$startI.Width = 80
$startI.Text = ''

$stopL = New-Object System.Windows.Forms.Label
$stopL.Location = New-Object System.Drawing.Point( (($gui.Width / 4) - 75), (($gui.Height / 2) - 22 ) )
$stopL.Text = "Stop:"
$stopL.AutoSize = $true

$stopI = New-Object System.Windows.Forms.TextBox
$stopI.Location = New-Object System.Drawing.Point( (($gui.Width / 4) - 45), (($gui.Height / 2) - 25 ) )
$stopI.Width = 80
$stopI.Text = ''
 
#### Epochs section END

$gui.Controls.Add( $extractB )
$gui.Controls.Add( $findB )
$gui.Controls.Add( $deviceL )
$gui.Controls.Add( $deviceBox )
$gui.Controls.Add( $fileL )
$gui.Controls.Add( $fileI )
$gui.Controls.Add( $epochsL )
$gui.Controls.Add( $startL )
$gui.Controls.Add( $startI )
$gui.Controls.Add( $stopL )
$gui.Controls.Add( $stopI )

$gui.Controls.Add( $picBox )      # or comment this out, and use below lines for BH logo

#$gui.Controls.Add( $bhBox )
#$gui.BackColor = "White"           # WhiteSmoke/White/

Clear
Write-Host
Write-Host
Write-Host "            Beacon Hill SG            " -ForegroundColor Blue -BackgroundColor White
Write-Host "                                      " -BackgroundColor Gray
Write-host "         " $deviceList.total "devices found" "          " -ForegroundColor Blue -BackgroundColor Gray
Write-Host "                                      " -BackgroundColor Gray
Write-Host 

$gui.ShowDialog()

####################################### END OF FRONTEND
