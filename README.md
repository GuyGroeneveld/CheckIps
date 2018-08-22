# CheckIps
This script retrieves known Office 365 subnets and check which established https connections are related

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

  .PARAMETER LogEnabled
  The name of a text file to write the mailbox name being scanned. This will appear also in the console with the Verbose mode
