
#Requires -Version 3
function Get-UptimeDays
{
    <#
    .Synopsis
       Calculates the system uptime of a local or remote computer
    .DESCRIPTION
       This function calculates the system uptime of a local or remote computer and returns total uptime, as well as the last boot time.  It will attempt to connect remotely via WinRM first, and fall back on DCOM if WinRM was unsuccessful.
    .EXAMPLEget
       Get-Uptime
       Returns the system uptime of the current computer
    .EXAMPLE
       Get-Uptime -ComputerName SERVER01
       Returns the system uptime of SERVER01
    .EXAMPLE
       Get-Content C:\computers.txt | Get-Uptime
       Returns the system uptime of all computers in a text file at C:\computers.txt. This example assumes there is one computer per line.
    .EXAMPLE
       Get-ADComputer -Filter * -SearchBase "OU=Servers,DC=example,DC=com" | Select -ExpandProperty Name | Get-Uptime
       Returns the system uptime of all computers in the Servers OU in the example.com domain. This example requires the ActiveDirectory module for Windows PowerShell3.
    .INPUTS
       [String[]] List of computers to check
    .OUTPUTS
       [PSCustomObject[]]
    #>
    [CmdletBinding()] param([parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][Alias('hostname')] $ComputerName)
    
    <# param(
        [Parameter(Position = 0,
                   ValueFromPipelineByPropertyName)]
        [Alias('Name','CN','__Server')]
        [String[]] $ComputerName = @($env:COMPUTERNAME) #>
          
    

    begin
    {
        [DateTime] $now = Get-Date
    }

    process
    {
        foreach ($c in $ComputerName)
        {
            Write-Verbose "Testing connection to computer [$c]"
            if (Test-Connection -ComputerName $c -Quiet -BufferSize 16 -Count 2)
            {
                $success = $false
                try
                {
                    [DateTime] $boot = Get-CimInstance -ComputerName $c -ClassName Win32_OperatingSystem -Property LastBootUpTime -ErrorAction Stop | Select-Object -ExpandProperty LastBootUpTime
                    $user = (Get-CimInstance -ComputerName $c -ClassName Win32_ComputerSystem).Username
                    $success = $true
                } catch {
                    Write-Verbose "Get-CimInstance from computer [$c] failed. Falling back on Get-WmiObject..."
                    try
                    {
                        $wmi = Get-WmiObject -ComputerName $c -Class Win32_OperatingSystem -Property LastBootUpTime
                        [DateTime] $boot = $wmi.ConvertToDateTime($wmi.LastBootUpTime)
                        $user = (Get-WmiObject -ComputerName $c -Class win32_computersystem).Username
                        $success = $true
                    } catch {
                        Write-Error "Unable to obtain a remote reference to the WMI class Win32_OperatingSystem on the remote system [$c]"
                    }
                }

                if ($success)
             {
                $obj = [PSCustomObject] @{
                    'ComputerName'   = $c;
                    'User'           = $user;
                    'LastBootUpTime' = $boot;
                    'Uptime'         = $now - $boot;


                }
                $d = $obj.Uptime.days
                [int]$obj.Uptime = [int]"${d}"


                    Write-Output $obj
                }
            } else {
                Write-Error "Unable to access computer [$c] remotely."
            }
        }
    }
}
