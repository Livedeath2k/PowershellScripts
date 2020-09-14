<# 
.SYNOPSIS Define the external host, the internal host and the Exchange server.  If not set it reads them from localhost.
.DESCRIPTION Define the external host, the internal host and the Exchange server that needs to be modified. If it's not set it will read the localhost and set the external url to internal url. 
.NOTES It can run from remote shell. Except Edge Servers. 
.COMPONENT Microsoft.Exchange.Management.PowerShell.SnapIn required. 
.Parameter InternalURI = internal servername like srv001.internal.mydomain.com
.Parameter ExternalURI = external servername like portal.mydomain.com. 
.Parameter Identity = The Exchange servername like MyMailmonster.
.Author = Lars Boos
#>
function set-exchangeuri {
$ComputerSystem = [System.Net.Dns]::GetHostByName(($env:computerName))
$ComputerSystem = $ComputerSystem | Select HostName
$identity = $env:computername

param(
    [Parameter(Mandatory = $true,ParameterSetName = "Internal Servername")]
    [string]$InternalURI=$ComputerSystem,  
    [Parameter(Mandatory = $true,ParameterSetName = "External Servername")]
    [string]$ExternalURI=$InternalURI,
    [Parameter(Mandatory = $true,ParameterSetName = "Identity of Exchange Server")]
    [string]$Identity,
    [Parameter(ParameterSetName = "Protection Standard")]
    [ValidateSet('Basic', 'NTLM', 'Negotiate')]
    [string]$AuthType="Negotiate"
)
 
#OWA
$owain = "https://" + "$InternalURI" + "/owa"
$owaex = "https://" + "$ExternalURI" + "/owa"
Get-OwaVirtualDirectory -Server $env:computername | Set-OwaVirtualDirectory -internalurl $owain -externalurl $owaex
 
#ECP
$ecpin = "https://" + "$InternalURI" + "/ecp"
$ecpex = "https://" + "$ExternalURI" + "/ecp"
Get-EcpVirtualDirectory -server $env:computername| Set-EcpVirtualDirectory -internalurl $ecpin -externalurl $ecpex
 
#EWS
$ewsin = "https://" + "$InternalURI" + "/EWS/Exchange.asmx"
$ewsex = "https://" + "$ExternalURI" + "/EWS/Exchange.asmx"
Get-WebServicesVirtualDirectory -server $env:computername | Set-WebServicesVirtualDirectory -internalurl $ewsin -externalurl $ewsex -confirm:$false -force
 
#ActiveSync
$easin = "https://" + "$InternalURI" + "/Microsoft-Server-ActiveSync"
$easex = "https://" + "$ExternalURI" + "/Microsoft-Server-ActiveSync"
Get-ActiveSyncVirtualDirectory -Server $env:computername  | Set-ActiveSyncVirtualDirectory -internalurl $easin -externalurl $easex
 
#OfflineAdressbuch
$oabin = "https://" + "$InternalURI" + "/OAB"
$oabex = "https://" + "$ExternalURI" + "/OAB"

Get-OabVirtualDirectory -Server $env:computername | Set-OabVirtualDirectory -internalurl $oabin -externalurl $oabex
 
#MAPIoverHTTP
$mapiin = "https://" + "$InternalURI" + "/mapi"
$mapiex = "https://" + "$ExternalURI" + "/mapi"
Get-MapiVirtualDirectory -Server $env:computername| Set-MapiVirtualDirectory -externalurl $mapiex -internalurl $mapiin
 
#Outlook Anywhere (RPCoverhTTP)

Get-OutlookAnywhere -Server $env:computername| Set-OutlookAnywhere -externalhostname $ExternalURI -internalhostname $InternalURI -ExternalClientsRequireSsl:$true -InternalClientsRequireSsl:$true -ExternalClientAuthenticationMethod $AuthType
 
#Autodiscover SCP
$autodiscover = "https://" + "$InternalURI" + "/Autodiscover/Autodiscover.xml"

Get-ClientAccessService $env:computername | Set-ClientAccessService -AutoDiscoverServiceInternalUri $autodiscover

#Output
write-host "OWA URL:" $owain
write-host "OWA URL:" $owaex
write-host "ECP URL:" $ecpin
write-host "ECP URL:" $ecpex
write-host "EWS URL:" $ewsin
write-host "EWS URL:" $ewsex
write-host "ActiveSync URL:" $easin
write-host "ActiveSync URL:" $easex
write-host "OAB URL:" $oabin
write-host "OAB URL:" $oabex
write-host "MAPI URL:" $mapiin
write-host "MAPI URL:" $mapiex
write-host "OA Hostname:" $InternalURI
write-host "OA Hostname:" $ExternalURI
write-host "Autodiscover URL:" $autodiscover
}
set-exchangeuri
