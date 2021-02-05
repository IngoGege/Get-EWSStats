# Get-EWSStats
 Retrieves statistics from CAS server for specific user from EWS logs.

## Required Parameters

### -EmailAddress

the given address is used for filtering. If omitted all entries will be reported.

## Optional Parameters

### -EmailAddresses

similar to the parameter EmailAddress, but for multiple addresses is filtered.

### -StartDate

this is used for filtering the logfiles to be parsed. The format must be yyMMdd. If omitted current date will be used.

### -EndDate

this is used for filtering the logfiles to be parsed. The format must be yyMMdd. If omitted current date will be used.

### -Logparser

this is used for the path to LogParser.exe

### -ADSite

here you can define in which ADSite is searched for Exchange server. If omitted current AD site will be used.

### -Outpath

where the output will be found. If omitted $env:temp will be used.

### -SpecifiedServers

the enumerated server will be filtered for the given server. Useful if you want to parse the logs on only a subset of servers.

### -Localpath

which folder to parse when the logfiles have been copied to a local path. You can't mix Exchange 2010 with Exchange 2013 RCA logs!

### -Exchange2013

looks only for Exchange 2013 and newer mailbox server.

### -OneFile

by default the script will generate for each day a dedicated file, while searching across multiple days. Using this switch the script will create only one file, but includes the day.

## Examples

# collect all entries for a given address with defined outpath
```
.\Get-EWSStats.ps1 -EmailAddress donald@entenhausen.com -startdate 130213 -enddate 130214 -Outpath c:\temp
```

## NOTES

You should run this script in the same AD site where the servers are due to latency.
