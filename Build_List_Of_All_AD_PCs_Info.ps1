#Note: Requires ImportExcel Module
#Run: Install-Module ImportExcel or Install-Module ImportExcel -scope CurrentUser
#Note: Update Line 7 *Domain* for your domain

Import-Module ActiveDirectory

Get-ADComputer -filter * -SearchBase "CN=Computers,DC=*Domain*,DC=local" |
    Select-Object -ExpandProperty Name |
    Sort-Object |
    Set-Content C:\temp\AD_Computer_Names.txt

$ComputerList = Get-Content "C:\temp\AD_Computer_Names.txt"

$Output = foreach ($Computer in $ComputerList) {
    try {
        Test-WSMan -ComputerName $Computer -ErrorAction Stop > $null

        $Session = New-CimSession -ComputerName $Computer
        $System = Get-CimInstance -CimSession $Session -ClassName Win32_ComputerSystem -Property Model, TotalPhysicalMemory
        $Proc = Get-CimInstance -CimSession $Session -ClassName Win32_Processor -Property Name
        $BIOS = Get-CimInstance -CimSession $Session -ClassName Win32_BIOS -Property SerialNumber
        $OS = Get-CimInstance -CimSession $Session -ClassName Win32_OperatingSystem -Property Caption

        $Hashtable = @{
            'Name'            = $Computer
            'Model'           = $System.Model
            'Serial Number'   = $BIOS.SerialNumber
            'OS'              = $OS.Caption
            'Memory Capacity' = $System.TotalPhysicalMemory / 1GB
            'CPU'             = $Proc.Name
            'Status'          = 'Online'
        }
    }
    catch {
        $Hashtable = @{
            'Name'            = $Computer
            'Model'           = $null
            'Serial Number'   = $null
            'OS'              = $null
            'Memory Capacity' = $null
            'CPU'             = $null
            'Status'          = 'Offline'
        }
    }
    finally {
        [PSCustomObject]$Hashtable
        if ($Session) { Remove-CimSession $Session -ErrorAction SilentlyContinue }
    }
}

$Output | Export-Excel -Path C:\Temp\MyReport.xlsx
