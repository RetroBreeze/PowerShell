<# Mapped Printer Script created by Shem Gyll #>

$cn = gethostlist
$output = ''
$finLoc = ''

#makes sure RR is not disabled
#Set-Service -Name RemoteRegistry -StartUpType Automatic
#Start-Service -Name RemoteRegistry

#basically illegal variable name
$disconnecteds = @()

<#$spiceLoc = getSpiceData -query "select name,location from devices where type='Computer'"
$locationArr =@()
#$locationObj = ''


foreach($prop in $spiceLoc)
{
        $spiceName = $prop.name
        $spiceLocation = $prop.location

        $locationObj = New-Object -TypeName psobject
        $locationObj | Add-Member -MemberType NoteProperty -Name hostName -Value $spiceName
        $locationObj | Add-Member -MemberType NoteProperty -Name locationName -Value $spiceLocation
        $locationArr += $locationObj
}
#> 

$result = foreach($comp in $cn){

    $testnet = ''
    #Tries to ping the computer, then stores results in variable $testnet
    try
    {
        $testnet = Test-Connection $comp -Count 2 -ErrorAction Stop
    }
    #If unable to ping the computer, it displays that it is not connected.
    catch
    {
        Write-Host "$comp is not connected" -ForegroundColor Red
        $disconnecteds += $comp
    }

    #If ping was successful, the script moves on, if not it loops back to the next computer
    if($testnet)
    {
       # $s = New-PSSession -ComputerName $comp -ErrorAction SilentlyContinue
        #Invoke-Command -Session $s -ScriptBlock{
        $spiceLoc = getSpiceData -query "select location from devices where name='$comp'"
        
        foreach($prop in $spiceLoc)
        {
            $locN = $prop.location
        }

        #$spiceLocation = $prop.location

        Invoke-Command -ComputerName $comp -ErrorAction SilentlyContinue -ScriptBlock {

       # $name = $comp

            #gets list of SIDS to recurse through to get the user
            $sids = Get-ChildItem Registry::HKEY_USERS -Exclude ".Default","*Classes*" | Select-Object Name -ExpandProperty Name
            #gets the users mapped printers by recursing through sids to determine who has mapped printers
    
            foreach($sid in $sids)
                {
                  #Get-ChildItem Registry::$sid\Printers
                  if (Test-Path -Path Registry::$sid\Printers\Connections)  
                    {
                        $goodsid = $sid

                        #trims string
                        $unsid = $sid.Substring(11)

                        #gets username from sid
                        $objSID = New-Object System.Security.Principal.SecurityIdentifier($unsid)
                        $objUser = $objSID.Translate( [System.Security.Principal.NTAccount])

                        #Might want to UN for some reason
                        $un = $objUser.Value

                        #Stores results into $output
                        $output = Get-ChildItem Registry::$goodsid\Printers\Connections
                        
                        $uName = $un
                      
                        #Since $pName can contain multiple strings as an array, we create an object for each string in the array,
                        #then output that obj as a column in our CSV. Also, we convert commas to backslash and remove "MDL\"
                        #from usernames for readability.
                        
                        $uName = $uName.Substring(4)
                        $uName = $uName.ToUpper()

                        #$locN = $locationObj | Select-Object -ExpandProperty locationName | Where-Object -EQ hostName
                       
                       $locN = $locationArr.GetValue($ENV:COMPUTERNAME) #| Select-Object name,locationName | Where-Object{$_.Properties -eq $comp}

                        foreach($o in $output)
                        {
                        $pName = $o.PSChildName

                        $pName = $pName.Replace(",","\")
                        $pSimpleName = $o.PSChildName
                        $pSimpleName.Replace(",,mdlfps,", "")

                       $obj = New-Object -TypeName PSobject
                        $obj | Add-Member -MemberType NoteProperty -Name Username -Value $uName
                        $obj | Add-Member -MemberType NoteProperty -Name Hostname -Value $ENV:COMPUTERNAME
                        $obj | Add-Member -MemberType NoteProperty -Name Location -Value $Using:locN #$Using:locationArr | Select-Object -ExpandProperty locationName | Where-Object {hostName -eq $ENV:COMPUTERNAME} 
                        $obj | Add-Member -MemberType NoteProperty -Name PrinterName -Value $pName
                        $obj
                        }
                        
                }
            } 

        }
    }
}

$result | Select-Object Username,Hostname,Location,PrinterName
sleep 1
"`n"
#Prompts user for path. If path is entered then we export results to $path\MappedPrinters.csv. Otherwise,
#we use C:\Temp\

$path = Read-Host "Enter path for MappedPrinters.csv (Enter for C:\Temp\)"
if($path -eq $null)
{
    "Saving to C:\Temp\MappedPrinters.csv..."
    $result | Select-Object Username,Hostname,Location,PrinterName | Export-Csv C:\Temp\MappedPrinters.csv -NoTypeInformation 
}
if($path -eq "")
{
    "Saving to C:\Temp\MappedPrinters.csv..."
    $result | Select-Object Username,Hostname,Location,PrinterName | Export-Csv C:\Temp\MappedPrinters.csv -NoTypeInformation 
}
Else
{
 $result | Select-Object Username,Hostname,Location,PrinterName | Export-Csv ($path + "MappedPrinters.csv") -NoTypeInformation
 "Saving to $path"+"MappedPrinters.csv..."
}
"`n"
"Done!"