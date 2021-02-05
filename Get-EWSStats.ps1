<#

.SYNOPSIS

Created by: https://ingogegenwarth.wordpress.com/
Version:    42 ("What do you get if you multiply six by nine?")
Changed:    12.01.2017

Retrieves statistics from CAS server for specific user from EWS logs.

.DESCRIPTION

The Get-EWStats.ps1 script is enumerating all CAS servers in the current AD site and parse all EWS log files within the given time range for the given mailbox or logonaccount.

The output will be CSV file.

.PARAMETER EmailAddress

the given address is used for filtering. If omitted all entries will be reported.

.PARAMETER EmailAddresses

similar to the parameter EmailAddress, but for multiple addresses is filtered.

.PARAMETER StartDate

this is used for filtering the logfiles to be parsed. The format must be yyMMdd. If omitted current date will be used.

.PARAMETER EndDate

this is used for filtering the logfiles to be parsed. The format must be yyMMdd. If omitted current date will be used.

.PARAMETER Logparser

this is used for the path to LogParser.exe

.PARAMETER ADSite

here you can define in which ADSite is searched for Exchange server. If omitted current AD site will be used.

.PARAMETER Outpath

where the output will be found. If omitted $env:temp will be used.

.PARAMETER SpecifiedServers

the enumerated server will be filtered for the given server. Useful if you want to parse the logs on only a subset of servers.

.PARAMETER Localpath

which folder to parse when the logfiles have been copied to a local path. You can't mix Exchange 2010 with Exchange 2013 RCA logs!

.PARAMETER Exchange2013

looks only for Exchange 2013 and newer mailbox server.

.PARAMETER OneFile

by default the script will generate for each day a dedicated file, while searching across multiple days. Using this switch the script will create only one file, but includes the day.

.EXAMPLE 

# collect all entries for a given address with defined outpath
.\Get-EWSStats.ps1 -EmailAddress donald@entenhausen.com -startdate 130213 -enddate 130214 -Outpath c:\temp

.NOTES

You should run this script in the same AD site where the servers are due to latency.

#>

param(
    [CmdletBinding()]
    [parameter( Mandatory=$false, Position=0)]
    [string]$Emailaddress,

    [parameter( Mandatory=$false, Position=1)]
    [array]$Emailaddresses,

    [parameter( Mandatory=$false, Position=2)]
    [int]$StartDate="$((get-date).ToString("yyMMdd"))",

    [parameter( Mandatory=$false, Position=3)]
    [int]$EndDate="$((get-date).ToString("yyMMdd"))",

    [parameter( Mandatory=$false, Position=4)]
    [ValidateScript({If (Test-Path $_ -PathType leaf) {$True} Else {Throw "Logparser could not be found!"}})]
    [string]$Logparser="C:\Program Files (x86)\Log Parser 2.2\LogParser.exe",

    [parameter( Mandatory=$false, Position=5)]
    [string[]]$ADSite,

    [parameter( Mandatory=$false, Position=6)]
    [ValidateScript({If (Test-Path $_ -PathType container) {$True} Else {Throw "$_ is not a valid path!"}})]
    [string]$Outpath = $env:temp,

    [parameter( Mandatory=$false, Position=7)]
    [array]$SpecifiedServers,

    [parameter( Mandatory=$false, Position=8)]
    [ValidateScript({If (Test-Path $_ -PathType container) {$True} Else {Throw "$_ is not a valid path!"}})]
    [string]$Localpath,

    [parameter( Mandatory=$false, Position=9)]
    [switch]$Exchange2013=$true,

    [parameter( Mandatory=$false, Position=10)]
    [switch]$OneFile
)

# check for elevated PS
if (! ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning -Message "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    break
}

# function to get the Exchangeserver from AD site
function GetExchServer {
    [CmdLetBinding()]
    #http://technet.microsoft.com/en-us/library/bb123496(v=exchg.80).aspx on the bottom there is a list of values
    param([array]$Roles,[string[]]$ADSites
        )
    Process {
        $valid = @("2","4","16","20","32","36","38","54","64","16385","16439","16423")
        foreach ($Role in $Roles){
            if (!($valid -contains $Role)) {
                Write-Output -InputObject "Please use the following numbers: MBX=2,CAS=4,UM=16,HT=32,Edge=64 multirole servers:CAS/HT=36,CAS/MBX/HT=38,CAS/UM=20,E2k13 MBX=54,E2K13 CAS=16385,E2k13 CAS/MBX=16439, E2K19=16423"
                return
            }
        }
        function GetADSite {
            param([string]$Name)
            if ($null -eq $Name) {
                [string]$Name = ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()).GetDirectoryEntry().Name
            }
            $FilterADSite = "(&(objectclass=site)(Name=$Name))"
            $RootADSite= ([ADSI]'LDAP://RootDse').configurationNamingContext
            $SearcherADSite = New-Object -TypeName System.DirectoryServices.DirectorySearcher -ArgumentList ([ADSI]"LDAP://$RootADSite")
            $SearcherADSite.Filter = "$FilterADSite"
            $SearcherADSite.pagesize = 1000
            $ResultsADSite = $SearcherADSite.FindOne()
            $ResultsADSite
        }
        $Filter = "(&(objectclass=msExchExchangeServer)(|"
        foreach ($ADSite in $ADSites){
            $Site=''
            $Site = GetADSite -Name $ADSite
            if ($null -eq $Site) {
                Write-Verbose -Message "ADSite $($ADSite) could not be found!"
            }
            else {
                Write-Verbose -Message "Add ADSite $($ADSite) to filter!"
                $Filter += "(msExchServerSite=$((GetADSite -Name $ADSite).properties.distinguishedname))"
            }
        }
        $Filter += ")(|"
        foreach ($Role in $Roles){
            $Filter += "(msexchcurrentserverroles=$Role)"
        }
        $Filter += "))"
        $Root= ([ADSI]'LDAP://RootDse').configurationNamingContext
        $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher -ArgumentList ([ADSI]"LDAP://$Root")
        $Searcher.Filter = "$Filter"
        $Searcher.pagesize = 1000
        $Results = $Searcher.FindAll()
        if ("0" -ne $Results.Count) {
            $Results
        }
        else {
            Write-Verbose -Message "No server found!"
        }
    }
}

$Path               = $null
[array]$LogFiles    = $null
[array]$TempPath    = $null
[string]$Logsfrom   = $null

if (!($Localpath)) {
        # get servers
        if ($Exchange2013) {
            [array]$servers = GetExchServer -Roles 54,16439,16423 -ADSites $ADSite
        }
        else {
            [array]$servers = GetExchServer -Roles 4,20,36,38 -ADSites $ADSite
        }
        if ($SpecifiedServers) {
            $Servers = $Servers | Where-Object -FilterScript {$SpecifiedServers -contains $_.Properties.name}
        }
        if ($Servers) {
            Write-Output -InputObject "Found the following Exchange servers:", $($Servers | ForEach-Object -Process {$_.Properties.name})
            foreach ($Server in $Servers) {
                [array]$TempPath += "\\" + $Server.Properties.name + "\" + ($Server.Properties.msexchinstallpath -as [string]).Replace(":","$") + "\Logging\Ews"
            }
        }
        else {
            Write-Output -InputObject "No server found!"
            break
        }
}
else {
    Write-Output -InputObject "Using the following path:", $Localpath
    [array]$TempPath = $Localpath
    $ADSite = "localfiles"
}

# validate all path
foreach ($Path in $TempPath) { 
    if (Test-Path -LiteralPath $Path) {
    [array]$ValidPath += $Path
    }
}
# get all items in final path
if ($ValidPath) {
    foreach ($Item in $ValidPath) {
        if (Test-Path -LiteralPath $Item){
            $LogFiles += Get-ChildItem -LiteralPath $Item -Filter "*.log"
        }
    }
}
else {
    Write-Output -InputObject "No logs found!"
    break
}

# filter and sort files
if (!($Localpath)) {
    if (($StartDate.ToString().Length -gt 6) -or ($EndDate.ToString().Length -gt 6)) {
        if (($StartDate.ToString().Length -gt 6) -and ($EndDate.ToString().Length -gt 6)) {
            $LogFiles = $LogFiles | Where-Object -FilterScript {$_.name.substring(6,8) -ge $StartDate -and $_.name.substring(6,8) -le $EndDate}
        }
        elseIf (($StartDate.ToString().Length -gt 6) -and ($EndDate.ToString().Length -eq 6)) {
            $LogFiles = $LogFiles | Where-Object -FilterScript {$_.name.substring(6,8) -ge $StartDate -and $_.name.substring(6,6) -le $EndDate}
        }
        else {
            $LogFiles = $LogFiles | Where-Object -FilterScript {$_.name.substring(6,6) -ge $StartDate -and $_.name.substring(6,8) -le $EndDate}
        }
    }
    else {
        $LogFiles = $LogFiles | Where-Object -FilterScript {$_.name.substring(6,6) -ge $StartDate -and $_.name.substring(6,6) -le $EndDate}
    }
}

if ($LogFiles) {
    $LogFiles | ForEach-Object -Process {$Logsfrom += "'" + $_.fullname +"',"}
    $Logsfrom = $Logsfrom.TrimEnd(",")
    Write-Output -InputObject "Logs to be parsed:"
    $LogFiles |Select-Object -Property fullname | Sort-Object -Property fullname
}
else {
    Write-Output -InputObject "No logs found!"
    break
}

if ($Exchange2013) {
    if ($Emailaddresses) {
        $stamp = "EWS2013_multiple_users_" + ($ADSite -join "_") + "_" + $(Get-Date -Format HH-mm-ss)
        if ($Onefile) {
            $stamp = $stamp + "_From_" + $StartDate.ToString() + "_To_" + $EndDate.ToString()
        }
    }
    else {
        $stamp = "EWS2013_" + $Emailaddress + "_" + ($ADSite -join "_") + "_" + $(Get-Date -Format HH-mm-ss)
        if ($Onefile) {
            $stamp = $stamp + "_From_" + $StartDate.ToString() + "_To_" + $EndDate.ToString()
        }
    }
$query_EWS = @"
SELECT  Day,Time AS TimeUTC,AuthenticatedUser,AuthenticationType,SoapCommand,ImpersonatedUser,ProxyAsUser,ActAsUser,ErrorCode,ServerHostName,FrontEndServer,Connections,HangConnections,
EndADBudget%,EndBudgetCAS%,EndBudgetRPC%,EndBudgetFindCount%,EndBudgetSubscriptions,
ThrottlingDelay,ThrottlingPolicy,TotalDCRequestLatency,TotalMBXRequestLatency,HttpPipelineLatency,TotalRequestTime,REVERSEDNS(ClientIpAddress) AS ClientHostName,
UserAgent,ThrottlingRequestType,AuthenticationErrors,GenericErrors

"@
}
else {
    $stamp = "EWS2010_" + $Emailaddress + "_" + ($ADSite -join "_") + "_" + $(Get-Date -Format HH-mm-ss)
$query_EWS = @"
SELECT  Day,Time AS TimeUTC,AuthenticatedUser,AuthenticationType,SoapCommand,ImpersonatedUser,ErrorCode,ServerHostName,Connections,HangConnections,
EndADBudget%,EndBudgetCAS%,EndBudgetRPC%,EndBudgetFindCount%,EndBudgetSubscriptions,
ThrottlingDelay,ThrottlingPolicy,TotalDCRequestLatency,TotalMBXRequestLatency,TotalRequestTime,REVERSEDNS(ClientIpAddress) AS ClientHostName,
UserAgent,ThrottlingRequestType,AuthenticationErrors,GenericErrors

"@
}

$query_EWS += @"
USING
TO_STRING(TO_TIMESTAMP(EXTRACT_PREFIX(REPLACE_STR([#Fields: datetime],'T',' '),0,'.'), 'yyyy-MM-dd hh:mm:ss'),'yyMMdd') AS Day,
TO_TIMESTAMP(EXTRACT_PREFIX(TO_STRING(EXTRACT_SUFFIX([#Fields: datetime],0,'T')),0,'.'), 'hh:mm:ss') AS Time,
AuthenticatedUser AS User,
SoapAction AS SoapCommand,
EndBudgetConnections AS Connections,
EndBudgetHangingConnections AS HangConnections,
SUBSTR(EndBudgetAD,ADD(LAST_INDEX_OF(EndBudgetAD, '/'),1)) AS EndADBudget%,
SUBSTR(EndBudgetCAS,ADD(LAST_INDEX_OF(EndBudgetCAS, '/'),1)) AS EndBudgetCAS%,
SUBSTR(EndBudgetRPC,ADD(LAST_INDEX_OF(EndBudgetRPC, '/'),1)) AS EndBudgetRPC%,
SUBSTR(EndBudgetFindCount,ADD(LAST_INDEX_OF(EndBudgetFindCount, '/'),1)) AS EndBudgetFindCount%

"@

if ($Onefile) {
$query_EWS += @"
INTO $outpath\$stamp.csv
FROM 
"@
}
else {
$query_EWS += @"
INTO $outpath\*_$stamp.csv
FROM 
"@
}

$query_EWS += $Logsfrom
if ($Emailaddress) {        
$query_EWS += @"

Where ((User LIKE '%$Emailaddress%') OR (ImpersonatedUser LIKE '%$Emailaddress%'))
"@
}
elseif ($Emailaddresses) {

[string]$QueryString= ""
foreach ($Address in $Emailaddresses) {
    $QueryString += "((User LIKE '%$Address%') OR (ImpersonatedUser LIKE '%$Address%')) OR " 
}

#build string from multiple addresses
$QueryString = $QueryString.Substring("0",($QueryString.LastIndexOf(")")+1))

$query_EWS += @"

WHERE $QueryString
"@

}
$query_EWS += @"

ORDER BY Day,Time
"@
Set-Content -Value $query_EWS -Path $outpath\query.txt -Force
Write-Output -InputObject 'Start query!'
& $Logparser file:$outpath\query.txt -i:CSV -o:csv -nSkipLines:5 -e:100 -iw:on -dtlines:0
Write-Output -InputObject 'Query done!'
# clean query file
Get-ChildItem -LiteralPath $outpath -Filter query.txt | Remove-Item -Confirm:$false | Out-Null