Import-Module PDFTools

$DebugPreference = 'Continue'

$errs = $null

$errs = @{}

# deafult path to script location using DIMS evn.variable

$configPath = "$env:DIMS\settings\dims-gui\config.xml"

# get .xml config

[xml]$config = Get-Content $configPath

$formPath = $config.Settings.Form.FormPath

# import files names and paths

[System.Collections.ArrayList]$errorsOutGrid = @()

function Get-AdDn {

    <#
        .SYNOPSIS
            Функция используется для преобразования строки DN
            стандартного вида в формат вида my.domain.ru/L1/L2/.../Ln .
        
        .INPUTS
            В качестве аргумента на вход принимается строка, 
            содержащяя полное значение DN, в составе всех RDN.

            OU=L3, OU=L3,OU=L4,DC=my,DC=domain,DC=r

        .OUTPUTS
            my.domain.ru/L1/L2/L3

        .EXAMPLE
            Get-AdDn -dpStr 'OU=Ln,OU=...,OU=L3,OU=L4,DC=my,DC=domain,DC=ru'
    #>

    Param(
    [string]$dpStr
    )

    $dpStr = $dpStr -replace '(OU=)|(DC=)'

    #$domainStr,$dpVrb,$ouVrb = $null

    #$ouVrb = [array]@()

    $arrStr = $dpStr.Split(",")

    foreach($w in $arrStr) {

        if($w -like 'main'){

            $domainStr += $w

            $domainStr += '.'

            $arrStr = $arrStr -ne $w

        }

        if($w -like 'russianpost'){

            $domainStr += $w

            $domainStr += '.'

            $arrStr = $arrStr -ne $w

        }

        if($w -like 'ru'){

            $domainStr += $w

            $arrStr = $arrStr -ne $w

        }

    }

    [array]::Reverse($arrStr)

    foreach($w in $arrStr) {

        $ouVrb += '/' + $w  

    }

    $dpVrb = $domainStr + $ouVrb
    
    return $dpVrb

}

function Get-ProcessData{

    Param(
    [int]$flag,
    [int]$processid
    )

    Try {

        switch($flag){
        
        1 {$data = gwmi Win32_Process | ? {$_.ProcessID -eq $processid} -ErrorAction Stop -ErrorVariable $err | select WS}
        2 {$data = (gwmi Win32_Process | ? {$_.ProcessID -eq $processid} -ErrorAction Stop -ErrorVariable $err).GetOwner()| select User}
        
        }

        return $data
    
    }

    Catch{
    
        Write-Debug "WMI error: $_"

        Write-Output "[$(Get-TimeStamp -type 1)] Process WMI error: $_" | Out-File $logPath -Append

        if ($err.length -ne 0){
            
            Write-Debug "$err"

            Write-Output "[$(Get-TimeStamp -type 1)] Process WMI error-variable: $err" | Out-File $logPath -Append
        
        } else {

            Write-Debug "Length of WMI error variable: 0"

            Write-Output "[$(Get-TimeStamp -type 1)] Process WMI errorvariable length equal 0" | Out-File $logPath -Append
            
        }
    }

}


function Get-TimeStamp {

    <#
        .SYNOPSIS
            Функция используется для получения значения
            временной метки в заданном формате. В контексте 
            скрипта, применяется для вывода временной
            метки в лог файл или консоль в режиме отладки,
            а также в структуре имен файлов.

        .INPUTS
            На вход принимается целочисленное значение индекса,
            соотвествующего формату вывода в контексте switch

        .OUTPUTS
            Варианты вывода, в соответствие заданным индексам:
            
            1. MM-dd-yyyy
            2. HH:mm:ss
            3. HH-mm

        .EXAMPLE
            Get-TimeStamp -type 1
    #>

    Param(
    [int]$type
    )
    switch ($type) {
        0 {return "{0:MM-dd-yyyy}" -f (Get-Date)}
        1 {return "{0:HH:mm:ss}" -f (Get-Date)}
        2 {return "{0:HH-mm}" -f (Get-Date)}
    }    
}

function Get-ID{

    # функция используется для получения уникального идентификатора сессии с целью записи файла отчета

    $unqId = 'ID'+ $(-join ((65..90) + (97..122) | Get-Random -Count 5 | % {[char]$_})).ToString().ToUpper() + "$(1..100 | Get-Random)"

    return $unqId
}

$logName =  "DIMS-GUI-log-$(Get-TimeStamp -type 0)-$(Get-TimeStamp -type 2).log"
$logPath = "$($config.Settings.exportPath.logPath)" + $logName


function Get-Xaml{

    Param(
    [string]$path
    )

    [xml]$Global:xmlWPF = Get-Content -Path $path

    Try{

        Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,system.windows.forms
    
    } 
    
    Catch{

        Throw "Failed to load Windows Presentation Framework assemblies."

        Write-Output "[$(Get-TimeStamp -type 1)] XAML load error: Failed to load Windows Presentation Framework assemblies." | Out-File $logPath -Append
    
    }

    $Global:xamlGUI = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $xmlWPF))

    $xmlWPF.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | 
    
    %{

        Set-Variable -Name ($_.Name) -Value $xamlGUI.FindName($_.Name) -Scope Global

    }

}

function Get-NicData {

    Param(
    [int]$flag,
    [int]$index,
    [string]$computername
    )

    try{
    
        switch($flag){

        1{$getNetAdapter = gwmi win32_networkadapterconfiguration -cn $computername -f "Index = $index" -ErrorAction Stop -ErrorVariable $err}
        2{$getNetAdapter = gwmi win32_networkadapter -cn $computername -f "Index = $index" -ErrorAction Stop -ErrorVariable $err}

        }
        
        return $getNetAdapter
    }

    Catch{
    
        Write-Debug "WMI error: $_"
        
        Write-Output "[$(Get-TimeStamp -type 1)] Get NIC WMI error: $_" | Out-File $logPath -Append


        if ($err.length -ne 0){
            
            Write-Debug "$err"

            Write-Output "[$(Get-TimeStamp -type 1)] Get NIC WMI error-variable: $err" | Out-File $logPath -Append
        
        } else {

            Write-Debug "Length of WMI error variable: 0"

            Write-Output "[$(Get-TimeStamp -type 1)] NIC WMI errorvariable length equal 0" | Out-File $logPath -Append
            
        }
    }

}

function Get-PcData {

    Param(
    [string]$type,
    [string]$computername
    )

    Try{
    
        switch($type){

        "CS" {$data = gwmi Win32_ComputerSystem -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "CPU" {$data = gwmi Win32_Processor -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "BIOS" {$data = gwmi Win32_BIOS -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "OS" {$data = gwmi Win32_OperatingSystem -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "MOBO" {$data = gwmi Win32_Baseboard -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "RAM" {$data = gwmi Win32_PhysicalMemory -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "USB" {$data = gwmi Win32_UsbDevice -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "HDD" {$data = gwmi Win32_DiskDrive -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "NIC" {$data = gwmi Win32_NetworkAdapterConfiguration -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "PRINTER" {$data = gwmi Win32_Printer -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "DISPLAY" {$data = gwmi WmiMonitorID -Namespace root\wmi -cn $computername -ErrorAction Stop -ErrorVariable $err}

        }
        
        return $data
    }

    Catch{
    
        Write-Debug "WMI error: $_"

        Write-Output "[$(Get-TimeStamp -type 1)] Get PC data WMI error: $_" | Out-File $logPath -Append

        if ($err.length -ne 0){
            
            Write-Debug "$err"

            Write-Output "[$(Get-TimeStamp -type 1)] PC data WMI error-variable: $err" | Out-File $logPath -Append
        
        } else {

            Write-Debug "Length of WMI error variable: 0"

            Write-Output "[$(Get-TimeStamp -type 1)] PC data WMI errorvariable length equal 0" | Out-File $logPath -Append
            
        }
    }

}

function Get-NicPrimaryIndex {
    
    <#
        .SYNOPSIS
            Функция используется для получения значения
            индекса сетевого адаптера по умолчанию . В контексте 
            скрипта, применяется для сбора информации о 
            TCP/IP конфигурации хоста средствами WMI.

        .INPUTS
            На вход принимается строкое значение имени целевого хоста

        .OUTPUTS
            Целочисленное десятичное значение индекса в составе сообщения отладки:
            "Primary NIC is 1"
            "Default GW: 192.168.1.1"

        .EXAMPLE
            Get-NicPrimaryIndex -computername TEST01
    #>

    Param(
    [string]$computername
    )
   
    Try{
    
    # Get active NIC index

        $getNicIndex = (gwmi win32_networkadapter -cn $computername -f "netconnectionstatus = 2" -ErrorAction Stop -ErrorVariable $err).Index

        Write-Debug "Test NIC $getNicIndex"

        if($getNicIndex -gt 0){
        
            Write-Debug "Selected host have more than 1 active NIC"

            foreach($index in $getNicIndex){

                Write-Debug "Test NIC $index"
                
                $getNicGW = (gwmi win32_networkadapterconfiguration -cn $computername -f "Index = $index").DefaultIPGateway

                if($getNicGW.length -gt 0){

                    $nicPrimaryIndex = $index
                
                    Write-Debug "Primary NIC is $nicPrimaryIndex"

                    Write-Debug "Default GW: $getNicGW"

                    return [int]$nicPrimaryIndex
                
                } else{
                
                    Write-Debug "GW is empty for NIC $index"
                
                }
                            
            }
        
        } else {
        
            Write-Debug "GW NIC $getNicIndex"

            return [int]$nicPrimaryIndex = $index
        
        }
    
    }

    Catch{

        Write-Debug $err
        
        $errMsg | Out-File $logPath -Append

        return $errMsg
    }

}

function Get-RamType {
    
    param(
    [string]$speed
    )

    switch($speed){

        {$_ -lt 1333}{$type = 'DDR2'}
        
        {$_ -ge 1333 -and $_ -le 2000}{$type = 'DDR3'}

        {$_ -gt 2000}{$type = 'DDR4'}       
        
    }
    
    return $type
}

Get-Xaml -path $formPath

$wsUp = @()
$wsDwn = @()
$progBarCounter = $null

# Active Directory datasource

$adWksSrch = $config.Settings.adSearchBase.wksDN
$adUsrSrch = $config.Settings.adSearchBase.usrDN

$adWksDnVerbose = Get-AdDn -dpStr $adWksSrch
$adUsrDnVerbose = Get-AdDn -dpStr $adUsrSrch

# get data from Active Directory

Try {

    $getAdWksEn = Get-ADComputer -SearchBase $adWksSrch -Filter {Enabled -Eq "True"} -ErrorVariable $adErr -ErrorAction Stop | Select -Exp Name
    $getAdUsrEn = Get-ADuser -SearchBase $adUsrSrch -Filter {Enabled -Eq "True"} -ErrorVariable $adErr -ErrorAction Stop | Select -Exp Name
    $getAdWksDis = Get-ADComputer -SearchBase  $adWksSrch -Filter {Enabled -Eq "False"} -ErrorVariable $adErr -ErrorAction Stop | Select -Exp Name
    $getAdUsrDis = Get-ADuser -SearchBase $adUsrSrch -Filter {Enabled -Eq "False"} -ErrorVariable $adErr -ErrorAction Stop | Select -Exp Name

}

Catch{

     Write-Debug $adErr
        
     $adErr | Out-File $logPath -Append

     return $adErr
}


function Decode {

    If ($args[0] -is [System.Array]) {

        [System.Text.Encoding]::ASCII.GetString($args[0])

    }
    Else {

        Write-Debug "Argument is not an array."

    }

}

# ping scan

Foreach ($wks in $getAdWksEn){

$progBarCounter++

[int]$pct = ($progBarCounter/$($getAdWksEn.count))*100

Write-Progress -Activity “Scanning workstations..”`
-status “Running on $wks. Loading $pct %” `
-percentComplete $pct

    $echoToWs = Test-Connection -ComputerName $wks -Count 1 -Quiet

        If ($echoToWs) {
    
            $wsUp += $wks  
        
         }

         elseif (!$echoToWs){
        
            $wsDwn += $wks
            
         }
}

Write-Progress -Activity “Scanning workstations..” -Completed

foreach ($ws in $wsUp) {

     $wksCbox.Items.add($ws) | Out-null

     Write-Debug "WKS ComboBox list filled successfully. Item $ws added"

}

if($wksCbox.Items.Count -gt 0){

    $scanBtn.IsEnabled = 'True'

    $usbDevBtn.IsEnabled = 'True'

    $prntDevBtn.IsEnabled = 'True'

    $hddBtn.IsEnabled = 'True'

    $ramBtn.IsEnabled = 'True'

    $dspBtn.IsEnabled = 'True'

    $nicBtn.IsEnabled = 'True'

} else {

    $errMsg = '"Error: ComboBox is empty."'

    Write-Debug $errMsg

    Write-Output "[$(Get-TimeStamp -type 1)] $errMsg" | Out-File $logPath -Append
}

# output ad counters 

$wksDn.Text = $adWksDnVerbose
$usrDn.Text = $adUsrDnVerbose
$enWksTxtb.Text = $getAdWksEn.Count
$disWksTxtb.Text = $getAdWksDis.Count
$enUsrTxtb.Text = $getAdUsrEn.Count
$disUsrTxtb.Text = $getAdUsrDis.Count
$totalWks.Text = $getAdWksEn.Count + $getAdWksDis.Count
$totalUsr.Text = $getAdUsrEn.Count + $getAdUsrDis.Count

# ping scan results output

$upHstTxtb.Text = $wsUp.Count
$dwnHstTxtb.Text = $wsDwn.Count

Write-debug "Process ID: $PID"

Write-Output "[$(Get-TimeStamp -type 1)] Process ID: $PID" | Out-File $logPath -Append

[int]$procMemory = [math]::Round($(Get-ProcessData -processid $PID -flag 1).WS/1MB,2)

Write-Debug "Process memory: $procMemory"

Write-Output "[$(Get-TimeStamp -type 1)] Process memory: $procMemory" | Out-File $logPath -Append

[string]$procOwner = (Get-ProcessData -processid $PID -flag 2).User

Write-Debug "Process owner: $procOwner"

Write-Output "[$(Get-TimeStamp -type 1)] Process owner: $procOwner" | Out-File $logPath -Append

$procData.Content = "Process owner: $procOwner | Memory: $procMemory Mb | PID: $PID"


$scanBtn.Add_Click({

    $global:tempName = $null
    $global:tempPath = $null
    $global:getUsbDevices = $null
    $procMemory = $null

    $sessionId = Get-ID

    Write-Debug "DIMS scan ID: $sessionId"
    
    Write-Output "[$(Get-TimeStamp -type 1)] DIMS scan ID: $sessionId" | Out-File $logPath -Append

    [int]$procMemory = [math]::Round($(Get-ProcessData -processid $PID -flag 1).WS/1MB,2)

    [System.Collections.ArrayList]$global:reportPrnt = @()

    $global:tempName = "report-$($wksCbox.SelectedItem)-$(Get-TimeStamp -type 0)-$(Get-TimeStamp -type 2)-$sessionId.pdf"

    write-debug "Report name: $tempName"

    Write-Output "[$(Get-TimeStamp -type 1)] Report name: $tempName" | Out-File $logPath -Append

    $global:tempPath = "$($config.Settings.exportPath.tempPath)" + $tempName

    write-debug "Report path: $tempPath"

    Write-Output "[$(Get-TimeStamp -type 1)] Report path: $tempPath" | Out-File $logPath -Append
    
    $pingScn = Test-Connection -cn $wksCbox.SelectedItem -Count 3 -Quiet

    Write-Debug "Host $($wksCbox.SelectedItem) ICMP status is $pingScn"

    Write-Output "[$(Get-TimeStamp -type 1)] Host $($wksCbox.SelectedItem) ICMP status is $pingScn" | Out-File $logPath -Append

    if($pingScn){
    
        $icmpTxtbUpdt.Text = 'UP'

        $nicPrimaryIndex = Get-NicPrimaryIndex -computername $wksCbox.SelectedItem

        $getNicPrimary = Get-NicData -flag 1 -computername $wksCbox.SelectedItem -index $nicPrimaryIndex

        $getNicPrimaryDev = Get-NicData -flag 2 -computername $wksCbox.SelectedItem -index $nicPrimaryIndex

        #otput primary NIC properties

        $connTypeTxtbUpdt.Text = $getNicPrimaryDev.AdapterType
        $ipAddrTxtbUpdt.Text = $getNicPrimary.IPAddress
        $dhcpTxtbUpdt.Text = $getNicPrimary.DHCPEnabled

        if($getNicPrimary.DHCPEnabled){

            $leaseStatus = $getNicPrimary.ConvertToDateTime($getNicPrimary.DHCPLeaseObtained)

        } else{
        
            $leaseStatus = 'IP is static'
        
        }
 
        $gwTxtbUpdt.Text = $getNicPrimary.DefaultIPGateway
        $macTxtbUpdt.Text = $getNicPrimary.MACAddress
        $nicMdTxtbUpdt.Text = $getNicPrimaryDev.Name
        $leaseTxtbUpdt.Text = $leaseStatus
        
        # get workstation data

        $getCs = Get-PcData -type CS -computername $wksCbox.SelectedItem
        $getCpu = Get-PcData -type CPU -computername $wksCbox.SelectedItem
        $getBios = Get-PcData -type BIOS -computername $wksCbox.SelectedItem
        $getOsVer = Get-PcData -type OS -computername $wksCbox.SelectedItem
        $getMobo = Get-PcData -type MOBO -computername $wksCbox.SelectedItem
        $getRam = Get-PcData -type RAM -computername $wksCbox.SelectedItem

        # ram calculations 

        $getRamSumm = $([math]::Round($getCs.TotalPhysicalMemory / 1GB, 2)).ToString() + " GB"
        $getRamType = Get-RamType -speed $($getRam.Speed)

        # output workstation data
        
        Write-Debug "BIOS serial length:  $($getBios.SerialNumber.Length)"

        if(($getBios.SerialNumber.Length -eq 0)`
        -or ($getBios.SerialNumber -contains " ")) {
            
            Write-Debug "BIOS serial is empty"

            $biosSn = "Not defined"

        } else {

            $biosSn = $getBios.SerialNumber

        }

        $mdTxtbUpdt.Text = $getCs.Model
        $snTxtbUpdt.Text = $biosSn  # getBios.SerialNumber
        $osTxtbUpdt.Text = $getOsVer.Caption
        $cpuTxtbUpdt.Text = $getCpu.Name
        $coresTxtbUpdt.Text = $getCpu.NumberOfCores
        $moboTxtbUpdt.Text = $getMobo.Product
        $ramSummTxtbUpdt.Text = $getRamSumm
        $ramTypeTxtbUpdt.Text = $getRamType
        $ramBanksTxtbUpdt.Text = @($getRam).Length

        # build report array

        $reportPrnt.Add("Hostname: $($wksCbox.SelectedItem)")
        $reportPrnt.Add('')
        
        $reportPrnt.Add('NIC status')
        $reportPrnt.Add('--------------------------------')
        $reportPrnt.Add("Connection type: $($getNicPrimaryDev.AdapterType)")
        $reportPrnt.Add("Speed: $([math]::Round($($getNicPrimaryDev.Speed) / 1MB, 2)), Mbps")
        $reportPrnt.Add("IPv4: $($getNicPrimary.IPAddress)")
        $reportPrnt.Add("Mask: $($getNicPrimary.IPSubnet)")
        $reportPrnt.Add("Default gateway: $($getNicPrimary.DefaultIPGateway)")
        $reportPrnt.Add("DNS: $($getNicPrimary.DNSServerSearchOrder)")
        $reportPrnt.Add("MAC: $($getNicPrimary.MACAddress)")
        $reportPrnt.Add("DHCP status: $($getNicPrimary.DHCPEnabled)")
        $reportPrnt.Add("Lease obtained: $($leaseStatus)")
        $reportPrnt.Add("Model: $($getNicPrimaryDev.Name)")
        $reportPrnt.Add('')

        $reportPrnt.Add('PC')
        $reportPrnt.Add('--------------------------------')
        $reportPrnt.Add("Model:$($getCs.Model)")
        $reportPrnt.Add("BIOS SN:$($getBios.SerialNumber)")
        $reportPrnt.Add("OS:$($getOsVer.Caption)")
        $reportPrnt.Add("OS Version:$($getOsVer.Version)")
        $reportPrnt.Add("OS Serial:$($getOsVer.SerialNumber)")
        $reportPrnt.Add("CPU Model:$($getCpu.Name)")
        $reportPrnt.Add("CPU Cores:$($getCpu.NumberOfCores)")
        $reportPrnt.Add("RAM type:$getRamType")
        $reportPrnt.Add("RAM Size:$getRamSumm")
        $reportPrnt.Add("RAM banks:$(@($getRam).Length)")
        
        $procData.Content = "Process owner: $procOwner | Memory: $procMemory Mb | PID: $PID | Session ID: $sessionId"
        $pdfRepBtn.IsEnabled = 'True'

        Write-Output "[$(Get-TimeStamp -type 1)] Host $($wksCbox.SelectedItem) scanned successfully." | Out-File $logPath -Append

    }

})

$pdfRepBtn.Add_Click({

    # create PDF report
    
    if(Test-Path $tempPath){

        Write-Debug "Report file already exists: $tempPath"

        Write-Output "[$(Get-TimeStamp -type 1)] Report file already exists: $tempPath" | Out-File $logPath -Append
    
        ii $tempPath

    } else{
    
      $reportPrnt | Out-PTSPDF -Path $tempPath -Open
    
      Write-Debug "Report file successfully created: $tempPath"

      Write-Output "[$(Get-TimeStamp -type 1)] Report file successfully created: $tempPath" | Out-File $logPath -Append

    }

})

$usbDevBtn.Add_Click({

    # get USB devices

    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()

    $getUsbDevices = Get-PcData -type USB -computername $wksCbox.SelectedItem

    $getUsbDevices | select Name, Manufacturer, PNPDeviceID, Service, Status | ogv -Title "USB devices on $($wksCbox.SelectedItem)"

    Write-Debug "Get USB data from $($wksCbox.SelectedItem) successfully!"

    Write-Output "[$(Get-TimeStamp -type 1)] Get USB data from $($wksCbox.SelectedItem) successfully!" | Out-File $logPath -Append

    Write-Debug "Query execution time: $stopTimeMin sec."

    Write-Output "[$(Get-TimeStamp -type 1)] Query execution time: $stopTimeMin sec." | Out-File $logPath -Append

})

$hddBtn.Add_Click({

    # get HDD

    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()

    $getHdd = Get-PcData -type HDD -computername $wksCbox.SelectedItem

    $getHdd | select Model, @{n='Size, GB'; e={[math]::Round($_.size/1GB, 2)}}, Caption | ogv -Title "HDD on $($wksCbox.SelectedItem)"

    Write-Debug "Get HDD data from $($wksCbox.SelectedItem) successfully!"

    Write-Output "[$(Get-TimeStamp -type 1)] Get HDD data from $($wksCbox.SelectedItem) successfully!" | Out-File $logPath -Append

    Write-Debug "Query execution time: $stopTimeMin sec."

    Write-Output "[$(Get-TimeStamp -type 1)] Query execution time: $stopTimeMin sec." | Out-File $logPath -Append

})

$ramBtn.Add_Click({

    # get RAM

    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()

    $getRam = Get-PcData -type RAM -computername $wksCbox.SelectedItem

    $getRam | select @{n='Bank'; e={$_.BankLabel}}, `
    @{n='Speed'; e={$_.ConfiguredClockSpeed}}, `
    @{n='Size, GB'; e={[math]::Round($_.capacity/1GB, 2)}}, `
    @{n='Voltage'; e={$_.ConfiguredVoltage/1000}} | 
    ogv -Title "RAM on $($wksCbox.SelectedItem)"

    $stopTimeMin = [math]::Round($stopWatch.Elapsed.TotalMinutes,2)
    
    Write-Debug "Get RAM data from $($wksCbox.SelectedItem) successfully!"

    Write-Output "[$(Get-TimeStamp -type 1)] Get RAM data from $($wksCbox.SelectedItem) successfully!" | Out-File $logPath -Append

    Write-Debug "Query execution time: $stopTimeMin sec."

    Write-Output "[$(Get-TimeStamp -type 1)] Query execution time: $stopTimeMin sec." | Out-File $logPath -Append

})

$dspBtn.Add_Click({

    # get DISPLAY

    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()

    $getDisplay = Get-PcData -type DISPLAY -computername $wksCbox.SelectedItem

    $getDisplay | select @{n='Vendor'; e={Decode $_.ManufacturerName -notmatch 0}}, ` 
    @{n='Product Code'; e={Decode $_.ProductCodeID -notmatch 0}}, `
    @{n='Serial Number'; e={Decode $_.SerialNumberID -notmatch 0}}, `
    @{n='Model name'; e={Decode $_.UserFriendlyName -notmatch 0}} |
    ogv -Title "DISPLAY on $($wksCbox.SelectedItem)"

    $stopTimeMin = [math]::Round($stopWatch.Elapsed.TotalMinutes,2)

    Write-Debug "Get DISPLAY data from $($wksCbox.SelectedItem) successfully!"

    Write-Output "[$(Get-TimeStamp -type 1)] Get DISPLAY data from $($wksCbox.SelectedItem) successfully!" | Out-File $logPath -Append

    Write-Debug "Query execution time: $stopTimeMin sec."

    Write-Output "[$(Get-TimeStamp -type 1)] Query execution time: $stopTimeMin sec." | Out-File $logPath -Append
})

$prntDevBtn.Add_Click({

    # get PRINTER

    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()

    $getPrinter = Get-PcData -type PRINTER -computername $wksCbox.SelectedItem

    $getPrinter | select Status, PrinterStatus, DetectedErrorState, Name, DriverName, `
    Default, Network, PortName, Shared, Sharename, Location | 
    ogv -Title "PRINTER on $($wksCbox.SelectedItem)"

    $stopTimeMin = [math]::Round($stopWatch.Elapsed.TotalMinutes,2)
    
    Write-Debug "Get PRINTER data from $($wksCbox.SelectedItem) successfully!"

    Write-Output "[$(Get-TimeStamp -type 1)] Get PRINTER data from $($wksCbox.SelectedItem) successfully!" | Out-File $logPath -Append

    Write-Debug "Query execution time: $stopTimeMin sec."

    Write-Output "[$(Get-TimeStamp -type 1)] Query execution time: $stopTimeMin sec." | Out-File $logPath -Append

})

$nicBtn.Add_Click({

    # get NIC

    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()

    $getNic = Get-PcData -type NIC -computername $wksCbox.SelectedItem

    $getNic | select index, @{n='DNS Hostname'; e={$_.dnshostname}}, `
    @{n='IPv4 Address'; e={$_.ipaddress[0]}}, `
    @{n='Subnet'; e={$_.ipsubnet[0]}}, ` 
    @{n='Gateway'; e={$_.defaultipgateway[0]}}, `
    @{n='Connection metric'; e={$_.IPConnectionMetric}}, `
    @{n='MAC-address'; e={$_.macaddress}}, `
    @{n='Name'; e={$_.description}}, `
    @{n='Service'; e={$_.servicename}}, `
    @{n='DHCP enabled?'; e={$_.dhcpenabled}}, `
    @{n='DHCP server'; e={$_.dhcpserver}}, `
    @{n='DHCP lease obtained'; e={$_.dhcpleaseobtained}}, `
    @{n='DHCP lease expired'; e={[datetime]::ParseExact($_.dhcpleaseexpired,"yyyyMMddHHmmss",$null)}}, `
    @{n='Domain suffix search order'; e={$_.dnsdomainsuffixsearchorder}}, 
    @{n='DNS servers search order'; e={$_.dnsserversearchorder}} |
    ogv -Title "NIC on $($wksCbox.SelectedItem)"

    $stopTimeMin = [math]::Round($stopWatch.Elapsed.TotalMinutes,2)
    
    Write-Debug "Get NIC data from $($wksCbox.SelectedItem) successfully!"

    Write-Output "[$(Get-TimeStamp -type 1)] Get NIC data from $($wksCbox.SelectedItem) successfully!" | Out-File $logPath -Append
    
    Write-Debug "Query execution time: $stopTimeMin sec."

    Write-Output "[$(Get-TimeStamp -type 1)] Query execution time: $stopTimeMin sec." | Out-File $logPath -Append

})

$xamlGUI.ShowDialog() | Out-Null
