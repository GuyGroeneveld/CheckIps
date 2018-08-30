<#
  .SYNOPSIS
  This script retrieves known Office 365 subnets and check which established https connections are related

  .DESCRIPTION
  Microsoft published the known IP Subnet as described in this link
  https://support.office.com/en-us/article/managing-office-365-endpoints-99cab9d4-ef59-4207-9f2b-3728eb46bf9a?ui=en-US&rs=en-US&ad=US#ID0EABAAA=2._Proxies

  The recommandation is to bypass the proxy for Exchange/Sharepoint/Teams/Skype. If this is done, we should see on PCs established connection to IP addresses
  that are part of the published subnets.

  This script currently works only for IPv4 source and destination address.

  At first run we get the known subnets and save it to a file with the last version number we got.
  Every new run, we check if we are still at the current version, if not we get the new one.

  We get the established connection on port 443 and then check if they relate to one of the published subnets

  .EXAMPLE
  To get established connection to known O365 subnets and having the known subnets saved in the current directory

  .\CheckIP.ps1
  
  
  .EXAMPLE
  To get report on only IP address related to known subnets and activity logged in checkips.log in the current directory

  .\CheckIP.ps1 -KnownIPOnly -LogEnabled

    
  .PARAMETER KnownIPOnly
  Switch parameter. Specifying it makes the oiutput with only matching addresses

  .PARAMETER ForceUploadList
  Switch parameter to force the upload of the list
  
  .PARAMETER LogEnabled
  The name of a text file to write the mailbox name being scanned. This will appear also in the console with the Verbose mode
  #>


[cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline=$True,Mandatory=$false,Position=0)]
        $Path = (Get-Location).Path,
        [switch]$LogEnabled=$false,
        [switch]$KnownIPOnly=$false,
        [switch]$ForceUploadList=$false
        )

$ScriptVersion = 1.2

if ($LogEnabled)
{
    $Log = $Path + "\" + "CheckIp.log"
}
else
{
    $log = $null
}

<#Function to create log entry in file $LogName#>
Function Write-LogEntry 
{
    param(
        [string] $LogName=$Null ,
        [string] $LogEntryText
         )

    [string]$currentdaytime = get-date -Format G
    [string]$logstring = "$currentdaytime : $LogEntryText"


    if ($LogName -NotLike $Null) 
    {
        
        $logstring | Out-File -FilePath $LogName -append
        Write-Verbose -Message $logstring
    }
    Else
    {
        Write-Verbose -Message $logstring
    }
}

Function Convert-IPv4AddrToBinary
{
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline=$True,Mandatory=$false,Position=0)]
        $Address,
        [string]$LogName=$log
        )
    $AddrInBinary = $null
    $AddrArray = ($Address.split("/"))[0].split(".")
    for ($i = 0; $i -lt 4 ; $i++)
    {
        If ($AddrArray[$i] -gt 0 )
        {
            $addrconverted = [convert]::tostring($AddrArray[$i],2)
            if ($addrconverted.length -lt 8)
            {
                $ZeroToAdd = 8 - $addrconverted.length
                for ($w = 0 ; $w -lt $ZeroToAdd ; $w++)
                {
                    $addrconverted = "0" + $addrconverted

                }
            }
            $AddrInBinary = $AddrInBinary + $addrconverted
        }
        Else
        {
            $AddrInBinary = $AddrInBinary + "00000000"
        }
    }
    return $AddrInBinary
}




Function Get-ServiceAreaIPs
{
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline=$True,Mandatory=$false,Position=0)]
        $ServiceArea,
        $iplist,
        [string]$LogName=$log
        )
    
    $HttpsIps = $iplist |  Where-Object { $_.TcpPorts -like "*443*" -and $_.ips -ne $null }
    $HttpsIps = $HttpsIps | Sort-Object serviceArea
    
    $HttpsIPList = [System.Collections.ArrayList]@()

    For ($ipall = 0 ; $ipall -lt $HttpsIps.count ; $ipall++)
    {
         $ipss = $HttpsIps[$ipall].ips.split(",")

         for ($xx = 0; $xx -lt $ipss.count ; $xx++)
         {
        
            [int]$Mask = ([string]$ipss[$xx]).split("/")[1]
                   
            <#Sometimes there are same IPs with different Urls
            for example
            111.221.112.0/21   {*.outlook.com, *.outlook.office.com, autodiscover-*.outlook.com}
            111.221.112.0/21   {outlook.office.com, outlook.office365.com} 
            So we merge the URLs and remove the dupplicate      
            #>
        
            $URLs = $HttpsIps[$ipall].urls
            $Category = $HttpsIps[$ipall].category

            if (($HttpsIPList -ne $null) -and $HttpsIPList.Definition.contains($ipss[$xx]))
            {
                
                $toremove = $HttpsIPList | where { $_.definition -eq $ipss[$xx]}
                
                
                $URLs = ($URLs + $toremove.urls)
                
                switch ($toremove.category)
                {
                    "optimize" 
                    {
                        $Category = "Optimize"
                    }
                    "Allow" 
                    {
                        if ($Category -eq "Optimize" -or $Category -eq "Default")
                        {
                            #$category is higher and doesn't change
                        }
                    }
                    "Default" 
                    {
                        if ($Category -eq "Allow")
                        {
                            $Category = $toremove.category
                        }
                    }
                }
                
                $HttpsIPList.Remove($toremove) 
            }   
            
            $AddrInBinary = $null

            If (([string]$ipss[$xx]).contains("."))
            {
                $AddrInBinary = Convert-IPv4AddrToBinary $ipss[$xx]
             } 
                        
            if ($AddrInBinary)
            {
                $WKLIPs = New-Object -type PSObject
                $WKLIPs | Add-Member -MemberType NoteProperty -Name "ServiceArea" -Value $HttpsIps[$ipall].serviceArea
                $WKLIPs | Add-Member -MemberType NoteProperty -Name "Mask" -Value $Mask
                $WKLIPs | Add-Member -MemberType NoteProperty -Name "MaskInBinary" -Value $AddrInBinary
                $WKLIPs | Add-Member -MemberType NoteProperty -Name "Definition" -Value $ipss[$xx]
                $WKLIPs | Add-Member -MemberType NoteProperty -Name "serviceAreaDisplayName" -Value $HttpsIps[$ipall].serviceAreaDisplayName
                $WKLIPs | Add-Member -MemberType NoteProperty -Name "urls" -Value $URLs
                $WKLIPs | Add-Member -MemberType NoteProperty -Name "required" -Value $HttpsIps[$ipall].required
                $WKLIPs | Add-Member -MemberType NoteProperty -Name "category" -Value $Category
                [void]$HttpsIPList.Add($WKLIPs)

            }
        }
    
    }
    
    Write-LogEntry -LogEntryText ("There are " + $HttpsIPList.Count + " different known subnets") -LogName $LogName
    
    return $HttpsIPList
}

Function Get-HttpsConnections
{
        [cmdletbinding()]
        Param (
            [parameter(ValueFromPipeline=$True,Mandatory=$false,Position=0)]
            [string]$LogName=$log
            )
    $connections = Get-NetTCPConnection -State Established -RemotePort 443

    Write-LogEntry ("There are " + $connections.count + " established https connections") -LogName $LogName

    $ConnectionList = [System.Collections.ArrayList]@()
    
    for ($conn = 0; $conn -lt $connections.count; $conn++)
    {
        $RemoteAddrInBinary = $null
        $remoteaddress = $connections[$conn].RemoteAddress
        If ($remoteaddress.contains("."))
        {
            $RemoteAddrInBinary = Convert-IPv4AddrToBinary $remoteaddress
        } 
        if ($remoteaddress.contains(":"))
        {
            #$RemoteAddrInBinary = Convert-IPv6AddrToBinary $remoteaddress
            
        }

        if ($RemoteAddrInBinary)
        {                  
            $process = Get-Process -id $connections[$conn].owningprocess
            $ConnectionDetails = New-Object -type PSObject
            $ConnectionDetails | Add-Member -MemberType NoteProperty -Name "Process" -Value $process.Name
            $ConnectionDetails | Add-Member -MemberType NoteProperty -Name "Description" -Value $process.Description
            $ConnectionDetails | Add-Member -MemberType NoteProperty -Name "RemoteAddress" -Value $RemoteAddress
            $ConnectionDetails | Add-Member -MemberType NoteProperty -Name "LocalAddress" -Value $connections[$conn].LocalAddress
            $ConnectionDetails | Add-Member -MemberType NoteProperty -Name "LocalPort" -Value $connections[$conn].LocalPort
            $ConnectionDetails | Add-Member -MemberType NoteProperty -Name "BinaryAddress" -Value $RemoteAddrInBinary
            [void]$ConnectionList.Add($ConnectionDetails)
        }

    }
    return $ConnectionList
}

Function Get-SubnetMatchs
{
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline=$True,Mandatory=$false,Position=0)]
        $IPs,
        $Subnets,
        [string]$LogName=$log
        )

    
    $AddressSubnetMatchList = [System.Collections.ArrayList]@()
    [string]$daytime = get-date -Format G
    [int]$countmatch = 0
    For ($x = 0 ; $x -lt $IPs.count ; $x++)
    {
       [bool]$foundmatch = $false
       
       For ($m = 0; $m -lt $Subnets.count ; $m++)
       {
   
            $substring = $Subnets[$m].mask
        
            if ($Subnets[$m].MaskInBinary.substring(0,$substring) -eq $IPs[$x].BinaryAddress.substring(0,$substring))
            { 
                                
                $AddressSubnetMatch = New-Object -type PSObject
                $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "Time" -value $daytime
                $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "Process" -Value $IPs[$x].process
                $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "localaddress" -Value $IPs[$x].localaddress
                $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "LocalPort" -Value $IPs[$x].LocalPort
                $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "Remoteip" -Value $IPs[$x].remoteaddress
                $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "subnet" -Value $Subnets[$m].definition
                $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "ServiceArea" -Value $Subnets[$m].ServiceArea
                $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "Required" -Value $Subnets[$m].required
                $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "Category" -Value $Subnets[$m].category
                $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "URLs" -Value ($Subnets[$m].Urls -join ",")
                [void]$AddressSubnetMatchList.Add($AddressSubnetMatch)
                $foundmatch = $True
                $countmatch++
            }
        
       }

       if ($foundmatch -eq $false -and $KnownIPOnly -eq $false)
       {
            $AddressSubnetMatch = New-Object -type PSObject
            $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "Time" -value $daytime
            $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "Process" -Value $IPs[$x].process
            $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "localaddress" -Value $IPs[$x].localaddress
            $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "LocalPort" -Value $IPs[$x].LocalPort
            $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "Remoteip" -Value $IPs[$x].remoteaddress
            $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "subnet" -Value "N/A"
            $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "ServiceArea" -Value "N/A"
            $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "Required" -Value "N/A"
            $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "Category" -Value "N/A"
            $AddressSubnetMatch | Add-Member -MemberType NoteProperty -Name "URLs" -Value "N/A"
            [void]$AddressSubnetMatchList.Add($AddressSubnetMatch) 
       }

    }

    Write-LogEntry -LogEntryText "There are $countmatch connections to known subnets" -LogName $LogName

    return $AddressSubnetMatchList
}


Function Get-ServiceAreaRestData
{
[cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline=$True,Mandatory=$false,Position=0)]
        $WorkPath = (Get-Location).Path,
        [string]$LogName=$log
        )
    $version = (Invoke-RestMethod -uri "https://endpoints.office.com/version/Worldwide?ClientRequestId=e0792bff-483a-4046-9b71-97293aebecb2").latest
    $DestinationPath = $WorkPath + "\" + "Endpoind_version_" + $version + "_script_version_" + $ScriptVersion + ".xml"
    if (!(Test-Path $DestinationPath) -or $ForceUploadList -eq $true)
    {
        #The file does not exist or is not the last version
        Write-LogEntry -LogEntryText "Retrieving ServiceArea IPs for version $version as we couldn't find it in $WorkPath" -LogName $LogName
        $EndPoints = Invoke-RestMethod -uri "https://endpoints.office.com/endpoints/Worldwide?noipv6&ClientRequestId=e0792bff-483a-4046-9b71-97293aebecb2"

        $ServiceAreaIps = Get-ServiceAreaIPs -iplist $EndPoints
        $ServiceAreaIps | Export-Clixml -Path $DestinationPath -Force
    }
    Else
    {
        Write-LogEntry -LogEntryText "Found current version of ServiceArea IPs in $WorkPath" -LogName $LogName
        $ServiceAreaIps = Import-Clixml -Path $DestinationPath
    }

    return $ServiceAreaIps
}

Write-LogEntry -LogEntryText "**** Starting script run *****" -LogName $Log

$listeips = Get-ServiceAreaRestData -WorkPath $Path
$activeips = Get-HttpsConnections

Get-SubnetMatchs -IPs $activeips -Subnets $listeips

