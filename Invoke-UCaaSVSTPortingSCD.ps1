
<#

.SYNOPSIS
    Applicable VST Toolkit script for searching DID range and produce a porting event SCD.
    Resulting CSV file will be appended if the same site name is used.

.DESCRIPTION
	Searches for Lync/Skype objects with matching DID range.
	
    Built-in paramters for use:
    -rangeDigits            -   Mandatory paramater, specify pattern to match (start range)
    -endRange               -   Optional paramter, specify end range pattern to NOT match
    -PortVoicePolicy        -   Mandatory paramter, specify future voice policy for porting event(must be valid VP name)
    -siteName               -   Mandatory paramter, IMPORTANT - Specify site name as RgsWorkFlows with no DID will also be found using the site naming convention

.EXAMPLE
    Below example specifies a number range to match, the future voice policy and the site name
    .\Invoke-UCaaSVSTPortingSCD.ps1 -rangeDigits 16104966 -PortVoicePolicy VP-AMER-CHI-GSIP -siteName USMO
	
.EXAMPLE
    Below example specified a number range to match with an end range to stop searching past
    .\Invoke-UCaaSVSTPortingSCD.ps1 -rangeDigits 16104966 -endRange 161049665 -PortVoicePolicy VP-AMER-CHI-GSIP -siteName USMO
	
.INPUTS
	None
	
.OUTPUTS
	[PS Object]

.NOTES
	Copyright Applicable 2018
	AUTHOR: Adnan Khan <adnan.khan@applicable.com>
	
	Change Log
    v1.0 - 19 November 2018 - A Khan - Initial - Peer Tech Approval - M Jameel. 
    v1.1 - 6th December 2018 - A Khan - Added commands to check users post VoicePolicy application for verification purposes
                                        - Corrected rollback scripts
                                        - Changed validation on digit range to 64-bit integer to accommodate long digit range
                                        
#>

#Requires -Version 3.0

[CmdletBinding()]

Param (

    [Parameter(Mandatory = $True)]
    [ValidateScript( {
        Try {
            [int64]$_
        }
        Catch {
            Throw "$_ is not a valid digit range."
        }
    })]
    [string]$rangeDigits,


    [Parameter(Mandatory = $False)]
    [ValidateScript( {
        Try {
            [int64]$_
        }
        Catch {
            Throw "$_ is not a valid digit range."
        }
    })]
    [string]$endRange,

    [Parameter(Mandatory = $True)]
    [ValidateScript( {
        $vp = Get-CsVoicePolicy -Identity $_ -ea SilentlyContinue
    

        If ($vp) {
            $True
        }
        else {
            Throw "Cannot find the Voice Policy $_ because it does not exist."
        }
    })]
    [string]$PortVoicePolicy,

    [Parameter(Mandatory = $True)]
    [ValidateNotNullOrEmpty()]
   
    [string]$siteName


)


Function Invoke-Search { 

    Param (
        [string]$ObjectType,
        [string]$PrivateLine
    )


    $r = Foreach ($obj in $result)
    {
        
        [PSCustomObject]@{
            PSTypeName         = 'Applicable.UCaaS.VST.SkypeObjects'
            Displayname        = $obj.Displayname
            SipAddress         = $obj.SipAddress
            RegistrarPool      = $obj.RegistrarPool
            LineURI            = $obj.LineURI
            ConferencingPolicy = $obj.ConferencingPolicy
            DialPlan           = $obj.DialPlan
            LegacyVoicePolicy  = $obj.VoicePolicy.FriendlyName
            PortVoicePolicy    = $PortVoicePolicy
            ObjectType         = $ObjectType
            AssignVoicePolicy  = "Get-" + $ObjectType + " -Identity " + $obj.SipAddress + " | Grant-CsVoicePolicy -PolicyName " + $PortVoicePolicy
            PostCheck     = "Get-" + $ObjectType + " -Identity " + $obj.SipAddress + "| Select-Object Displayname,Sipaddress,Registrarpool,LineURI,ConferencingPolicy,DialPlan,VoicePolicy | Export-Csv c:\temp\"+ $siteName +"_USER_POST_CHECK.csv -NoTypeInformation -Append"
            RollBackScript     = "Get-" + $ObjectType + " -Identity " + $obj.SipAddress + "| Grant-CsVoicePolicy -PolicyName " + $obj.VoicePolicy.FriendlyName
            
        
        }
    }

    
    if ($r) {

        $r | export-csv -NoTypeInformation -Append -Path c:\temp\$fileName"_DID_PORT_SCD.csv"
   

    }
    else {
            if ($PrivateLine) {
                Write-Warning "No users with PrivateLine numbers matched given range"
            } else {

        Write-Warning  "No objects found with given number range for object type $ObjectType"
        }
    } 
}


[string]$fileName = $siteName.ToUpper()


$expr = "Get-CsUser -Filter {lineuri -like '*$rangeDigits*'} -resultsize unlimited" # Construct FILTER string for expression invoked below

Write-Progress -Activity "Searching objects in Active Directory....." -Status "User LineURI search"
$result = Invoke-Expression $expr

if ($endRange) {
    $result = $result | Where-Object {$_.lineuri -notmatch $endRange}

}
Invoke-Search -ObjectType CsUser


$expr = "Get-CsUser -Filter {privateline -like '*$rangeDigits*'}" # Construct FILTER string for expression invoked below
Write-Progress -Activity "Searching objects in Active Directory....." -Status "User PrivateLine search"
$result = Invoke-Expression $expr

if ($endRange) {
    $result = $result | Where-Object {$_.privateline -notmatch $endRange}

}

Invoke-Search -ObjectType CsUser -PrivateLine PrivateLine


$expr = "Get-CsAnalogDevice -Filter {lineuri -like '*$rangeDigits*'}" # Construct FILTER string for expression invoked below
Write-Progress -Activity "Searching objects in Active Directory....." -Status "Analog Devices search"
$result = Invoke-Expression $expr

if ($endRange) {
    $result = $result | Where-Object {$_.lineuri -notmatch $endRange}

}

Invoke-Search -ObjectType CsAnalogDevice


$expr = "Get-CsCommonAreaPhone -Filter {lineuri -like '*$rangeDigits*'}" # Construct FILTER string for expression invoked below
Write-Progress -Activity "Searching objects in Active Directory....." -Status "Common Area Phones search"
$result = Invoke-Expression $expr

if ($endRange) {
    $result = $result | Where-Object {$_.lineuri -notmatch $endRange}

}

Invoke-Search -ObjectType CsCommonAreaPhone


$expr = "Get-CsExUmContact -Filter {lineuri -like '*$rangeDigits*'}" # Construct FILTER string for expression invoked below
Write-Progress -Activity "Searching objects in Active Directory....." -Status "Exchange UM Contact search"
$result = Invoke-Expression $expr

if ($endRange) {
    $result = $result | Where-Object {$_.lineuri -notmatch $endRange}

}


Invoke-Search -ObjectType CsExUmContact


$expr = "Get-CsDialinConferencingAccessNumber -Filter {lineuri -like '*$rangeDigits*'}" # Construct FILTER string for expression invoked below
Write-Progress -Activity "Searching objects in Active Directory....." -Status "Dial In Conferencing Access Number search"
$result = Invoke-Expression $expr

if ($endRange) {
    $result = $result | Where-Object {$_.lineuri -notmatch $endRange}

}

Invoke-Search -ObjectType CsDialInConferencingAccessNumber


Write-Progress -Activity "Searching objects in Active Directory....." -Status "CS Application search"
$result = Get-CsApplicationEndpoint | Where-Object {$_.lineuri -like "*$rangeDigits*" -Or $_.DisplayName -match $sitename}

if ($endRange) {
    $result = $result | Where-Object {$_.lineuri -notmatch $endRange}

}

Invoke-Search -ObjectType CsApplicationEndPoint

