<#

Get-CsCertificate information

#>


<#function save-certdetails {
Param (
[string]$certdetails
[string]$filename
)

$path = Split-Path ($MyInvocation.MyCommand.Path)

$certdetails | out-file $path\$filename+".txt"

}
#>


$options = ("Front End Pool", "Edge Server Internal", "Edge Server External", "OAuth Cert", "Office Web Apps Farm")
$x = 0

foreach ($item in $options) {
write-host "$x $item"
$x++
}

[int]$choice = Read-Host "Select cert type  :"

if ($choice -eq 0) {

$cert = Get-CsCertificate | ?{$_.Use -eq "Default"} 

$validfrom = $cert.NotBefore

$validto = $cert.NotAfter
$subject = ($cert.Subject).Split(",")
$thumbprint = $cert.Thumbprint
$serial = $cert.SerialNumber

$sans = ($cert.alternativenames).Split(",") | Sort-Object

$details = "`n`n`n---SUBJECT---`n"
$details += $subject[0]+"`n`n"
$details += "---VALIDITY---`n"
$details += "Valid From:  $validfrom`n"
$details += "Valid To:  $validto`n`n"
$details += "---ID---`n"
$details +="Thumbprint: $thumbprint`n"
$details +="Serial Number: $serial`n`n"
$details +="---SUBJECT ALTERNATIVE NAMES---`n"

foreach ($san in $sans) {
$details += $san + "`n"
}
$details



}

if ($choice -eq 1) {

$cert = Get-CsCertificate | ?{$_.Use -eq "Internal"} 

$validfrom = $cert.NotBefore

$validto = $cert.NotAfter
$subject = ($cert.Subject).Split(",")
$thumbprint = $cert.Thumbprint
$serial = $cert.SerialNumber

$sans = ($cert.alternativenames).Split(",") | Sort-Object

$details = "`n`n`n---SUBJECT---`n"
$details += $subject[0]+"`n`n"
$details += "---VALIDITY---`n"
$details += "Valid From:  $validfrom`n"
$details += "Valid To:  $validto`n`n"
$details += "---ID---`n"
$details +="Thumbprint: $thumbprint`n"
$details +="Serial Number: $serial`n`n"
$details +="---SUBJECT ALTERNATIVE NAMES---`n"
foreach ($san in $sans) {
$details += $san + "`n"
}
$details

}

if ($choice -eq 2) {

$cert = Get-CsCertificate | ?{$_.Use -eq "AccessEdgeExternal"} 

$validfrom = $cert.NotBefore

$validto = $cert.NotAfter
$subject = ($cert.Subject).Split(",")
$thumbprint = $cert.Thumbprint
$serial = $cert.SerialNumber

$sans = ($cert.alternativenames).Split(",") | Sort-Object

$details = "`n`n`n---SUBJECT---`n"
$details += $subject[0]+"`n`n"
$details += "---VALIDITY---`n"
$details += "Valid From:  $validfrom`n"
$details += "Valid To:  $validto`n`n"
$details += "---ID---`n"
$details +="Thumbprint: $thumbprint`n"
$details +="Serial Number: $serial`n`n"
$details +="---SUBJECT ALTERNATIVE NAMES---`n"
foreach ($san in $sans) {
$details += $san + "`n"
}
$details
}

if ($choice -eq 3) {

$cert = Get-CsCertificate | ?{$_.Use -eq "OAuthTokenIssuer"} 

$validfrom = $cert.NotBefore

$validto = $cert.NotAfter
$subject = ($cert.Subject).Split(",")
$thumbprint = $cert.Thumbprint
$serial = $cert.SerialNumber

$sans = ($cert.alternativenames).Split(",") | Sort-Object

$details = "`n`n`n---SUBJECT---`n"
$details += $subject[0]+"`n`n"
$details += "---VALIDITY---`n"
$details += "Valid From:  $validfrom`n"
$details += "Valid To:  $validto`n`n"
$details += "---ID---`n"
$details +="Thumbprint: $thumbprint`n"
$details +="Serial Number: $serial`n`n"
$details +="---SUBJECT ALTERNATIVE NAMES---`n"
foreach ($san in $sans) {
$details += $san + "`n"
}
$details
}

if ($choice -eq 4) {

$friendlyname =  (Get-OfficeWebAppsFarm).CertificateName
$cert = Get-ChildItem -Path cert:\LocalMachine\My | ?{$_.FriendlyName -match $friendlyname}

$validfrom = $cert.NotBefore

$validto = $cert.NotAfter
$subject = ($cert.Subject).Split(",")
$thumbprint = $cert.Thumbprint
$serial = $cert.SerialNumber

$sans = ($cert.Extensions | Where-Object {$_.Oid.FriendlyName -eq "subject alternative name"}).Format(1) | Sort-Object

$details = "`n`n`n---SUBJECT---`n"
$details += $subject[0]+"`n`n"
$details += "---VALIDITY---`n"
$details += "Valid From:  $validfrom`n"
$details += "Valid To:  $validto`n`n"
$details += "---ID---`n"
$details +="Thumbprint: $thumbprint`n"
$details +="Serial Number: $serial`n`n"
$details +="---SUBJECT ALTERNATIVE NAMES---`n"
foreach ($san in $sans) {
$san = $san -Replace "DNS Name=", ""
$details += $san + "`n"
}
$details

}