#!/usr/bin/pwsh
#
#   Script to read hosts from excel file or csv and add DNS records
#

Param(
    [switch] $DryRun,
    [string] $DnsServer = "localhost",
    [string] $RemoteHost = "",
    [string] $Auth = "",
    [string] $AuthPwd  = "",
    [string] $AuthPwdFile  = "",
    [Parameter(Mandatory=$TRUE, ParameterSetName="csv")] [string] $CsvFile,
    [Parameter(Mandatory=$TRUE, ParameterSetName="excel")] [string] $ExcelFile,
    [Parameter(ParameterSetName="excel")][string] $ExcelSheetName = "Hosts",
    [Parameter(Mandatory=$TRUE, ParameterSetName="excel")][string] $ExcelRangeNames
)

#
# functions
#

# Override the default Write-Error function because the default is... dumb
function Write-Error($message) {
    [Console]::ForegroundColor = 'red'
    [Console]::Error.WriteLine("ERROR: $message")
    [Console]::ResetColor()
}

# read an excel file and return a hashtable of hosts
function readExcel($file) {
    # variable to store the hosts from the named ranges we want
    $hosts = @()

    # debug info
    Write-Debug "Reading hosts from sheet '$ExcelSheetName' in ranges named like '$ExcelRangeNames' on '$ExcelFile'"

    # open excel file
    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open($file)

    # foreach range in the workbook
    foreach ($name in $workbook.Names) {
        # we only care about specific named ranges that match the $ExcelRangeNames regex
        if ($name.Name -match $ExcelRangeNames) {
            # Debug info
            Write-Debug "Getting hosts from range $($name.Name)"

            # foreach row in the range
            foreach ($row in $workbook.WorkSheets.Item($ExcelSheetName).Range($name.Name).Rows) {
                # column format:
                # 1 - Name / Description
                # 2 - IP address
                # 3 - FQDN
                # 4 - CNAME (Alias)
            
                # check if we got a name, some rows can be empty
                if ($row.Cells(1,1).Text) {
                    $hosts += @{
                        name = $row.Cells(1,1).Text;
                        ip = $row.Cells(1,2).Text;
                        fqdn = $row.Cells(1,3).Text;
                        cname = $row.Cells(1,4).Text;
                    }
                }
            }
        }
    }

    # quit excel
    $excel.Quit()

    # return the hosts we read from excel
    return $hosts
}

# read hosts from CSV file
function readCSV($file) {
    # variable to store the hosts from the named ranges we want
    $hosts = @()

    # read hosts from csv and store in hosts array
    Import-Csv $file | `
    ForEach-Object {
        $hosts += @{
            name = $_.name;
            ip = $_.ip;
            fqdn = $_.fqdn;
            cname = $_.cname;
        }
    }

    # return the hosts we read from CSV
    return $hosts
}

# wrapper for destructrive commands
# if -DryRun flag was passed, just print the command instead of executing it
function maybeDo($cmd) {
    #check -DryRun flag
    if ($DryRun) {
        # print the command
        [Console]::ForegroundColor = 'cyan'
        [Console]::Error.WriteLine("DRY RUN: $cmd")
        [Console]::ResetColor()
    } else {
        # if debug, write destructive commands
        Write-Debug "Executing destructive command: $cmd"

        # create scriptblock from $cmd
        $script = [scriptblock]::Create($cmd)

        #execute the command
        Invoke-Command -ScriptBlock $script
    }
}

# get the forward host name for the fqdn pards
# returns forward host name if valid
# returns $FALSE if invalid
function getFwdHost($fqdnParts) {
    # validate input
    if ($fqdnParts.Count -lt 3) {
        return $FALSE
    }

    # return the host component (everything except the domain name and tld)
    return $($fqdnParts[0..$($fqdnParts.Count - 3)]) -join '.'
}

# get the forward zone name for the fqdn parts
# assumes $fqdnParts has already been validated by getFwdHost
# returns forward lookup zone
function getFwdZone($fqdnParts) {
    return $($fqdnParts[$($fqdnParts.Count - 2)..$($fqdnParts.Count - 1)]) -join '.'
}

# get the ptr host name for the ip address parts
# returns ptr host name if valid
# returns $FALSE if invalid
function getPtrHost($ipParts) {
    # validate input
    if ($ipParts.Count -ne 4) {
        return $FALSE
    }

    # return the host component (last octet of the ip address)
    return $ipParts[3]
}

# get the ptr zone name for the ip address parts
# assumes $ipParts has already been validated by getPtrHost
# returns ptr lookup zone
function getPtrZone($ipParts) {
    # build the ptr zone from the ip address (first 3 octets)
    $ptrZone = $ipParts[0..$($ipParts.Count-2)]

    # reverse the array to get it in ptr zone format
    [array]::Reverse($ptrZone)

    # add in-addr.arpa suffix
    $ptrZone += "in-addr.arpa"

    # join the parts into a single string and return
    return $ptrZone -join '.'
}

# validate that this zone exists
# returns $TRUE if zone exists
# returns $FALSE if zone doesn't exist
function checkZone($zone) {
    # try to lookup zone by name
    try {
        Get-DnsServerZone -ComputerName $DnsServer -Name $zone -ErrorAction 'Stop'
    } catch {
        Write-Error $_
        return $FALSE
    }

    # if we got here the zone exists
    return $TRUE
}

function lookupARecordsByIP($ip, $zone) {
    # try looking up records where ip matches
    try {
        $records = @(Get-DnsServerResourceRecord -ComputerName $DnsServer -ZoneName "$zone" -RRType "A" -ErrorAction 'Stop' | where {$_.RecordData.IPv4Address -eq "$ip"})
        
        # if we found existing records, warn
        if ($records.Count -gt 0) {
            Write-Warning "The following A records also point to '$ip'"
            # foreach record
            foreach ($r in $records) {
                Write-Warning "$($r.HostName) --> $ip"
            }
            Write-Warning "While this is allowed, it is generally not recommended unless necessary"
        }
    } catch {
        # failed a record lookup by ip
        Write-Error "Failed looking up records with IP '$ip' in '$zone'"
        Write-Error $_
        return $FALSE
    }

    # if we got here we succeeded looking up records for this ip
    return $TRUE
}

# function to get the dns zone parts (everything except the host name)
function lookupAndRemoveOldRecords($name, $zone, $type) {
    # array to hold records we looked up
    $records = @()

    # try looking up existing records with this name
    try {
        # lookup records
        $records = @(Get-DnsServerResourceRecord -ComputerName $DnsServer -ZoneName "$zone" -RRType "$type" -Name "$name" -ErrorAction 'Stop')
    } catch {
    }

    # if we found any records, need to remove them
    if ($records.Count -gt 0) {
        # warn we found some existing records
        Write-Warning "Found $($records.Count) existing record(s) for '$name.$zone'"
        try {
            Write-Warning "Removing all '$type' record(s) for '$name.$zone'"
            # remove any existing records with this name
            maybeDo "Remove-DnsServerResourceRecord -ComputerName $DnsServer -ZoneName $zone -RRType $type -Name $name -Confirm:`$FALSE -Force -ErrorAction 'Stop'"
        } catch {
            Write-Error "Failed removing old '$type' record(s) for '$name.$zone'"
            Write-Error $_
            return $FALSE
        }

        # check for and remove any cname records pointing to this host if we're adding an a record
        if ($type -eq "A") {
            # if we're adding an a record, find and remove any cname records pointing to it
            foreach ($r in $(Get-DnsServerResourceRecord -ComputerName $DnsServer -ZoneName "$zone" -RRType "CName" | where {$_.RecordData.HostNameAlias -eq "$name.$zone."})) {
                # warn user we're removing this cname record
                Write-Warning "Removing existing CNAME record '$($r.HostName).$zone' --> '$($r.RecordData.HostNameAlias)'"
                
                # try removing the cname record
                try {
                    $r | Remove-DnsServerResourceRecord -ComputerName $DnsServer -ZoneName "$zone" -Confirm:$FALSE -Force -ErrorAction 'Stop'
                } catch {
                    Write-Error "Failed removing CNAME record"
                    Write-Error $_
                    return $FALSE
                }
            }
        }
    }
    
    # if we got here we succeeded looking up & removing old records
    return $TRUE
}

# function to cleanly exit (close any open sessions, etc)
function CleanExit {
    # is there a $RemoteHost?
    if ($RemoteHost) {
        Remove-PSSession $RemoteHost
    }
    
    # remove the DnsServer module, it's possible we remoted this module
    # doesn't hurt to do this if we're local, it'll just be reloaded automagically if any DnsServer module commands are entered
    Remove-Module DnsServer

    # bye bye
    exit
}

#
# application begin
#

# stop execution on any uncaught errors
$ErrorActionPreference = 'Stop'

# don't run if we're not on PS7
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Error "This script requires PowerShell 7 or higher"
}

# if $RemoteHost is present, create session on that host
if ($RemoteHost) {
    # did we get a user with $Auth?
    if ($Auth) {
        # check for $AuthPwd (this takes precedence)
        if ($AuthPwd -and $AuthPwdFile) {
            Write-Debug "Using -AuthPwd from arguments instead of -AuthPwdFile."
        # otherwise, check for $AuthPwdFile
        } elseif ($AuthPwdFile) {
            # debug
            Write-Debug "Reading password from '$AuthPwdFile'."
            
            # TODO: CHECK FILE PERMS AND WARN IF GLOBALLY READABLE
            # TODO: Linux: (Get-ChildItem $AuthPwdFile).UnixMode -match "^.{4}r.*" or "^.{7}r.*"
            # TODO: Windows: ...... windows file permissions dumb.....
            
            # read file
            $AuthPwd = Get-Content $AuthPwdFile
        }

        # check if we got an $AuthPwd
        if ($AuthPwd) {
            # debug
            Write-Debug "Converting -AuthPwd to SecureString."
            
            # convert password to a secure string
            $AuthPwdSString = ConvertTo-SecureString $AuthPwd -AsPlainText -Force
            
            # set $Auth to a new PSCredential object
            $AuthPSCred = New-Object System.Management.Automation.PSCredential ($Auth, $AuthPwdSString)
        }
        
        # default to value from -Auth
        $AuthCred = $Auth
        
        # but check if we have an $AuthPSCred and use that instead
        if ($AuthPSCred) {
            $AuthCred = $AuthPSCred
        }

        # TODO: wrap w/ try
        # create remote session
        $RemoteSession = New-PSSession -ComputerName $RemoteHost -Authentication Negotiate -Credential $AuthCred
    # otherwise, check if we're not on windows (!windows needs $Auth)
    } elseif (-not ($IsWindows)) {
        Write-Error "-Auth needed if not running on Windows."
        exit
    } else {
        # TODO: wrap w/ try
        # no need to pass -Authentication Negotiate here, we're on windows and have ntlm
        $RemoteSession = New-PSSession -ComputerName $RemoteHost
    }
    
    # at this point we should have a $RemoteSession, if not, there's a problem
    if ($RemoteSession.State -ne "Opened") {
        Write-Error "Failed creating remote session."
        exit
    }
    
    # import DnsServer module
    Import-Module -PSSession $RemoteSession DnsServer
}

# check for DnsServer module
if (-not (Get-Module | Where-Object {$_.Name -eq 'DnsServer'})) {
    Write-Error "DnsServer module not available - try specifying a -RemoteHost with the DnsServer module installed."
    exit
}

# test connection to $DnsServer
try {
    Get-DnsServer -ComputerName $DnsServer | Out-Null
} catch {
    Write-Error "Could not connect to '$DnsServer'."
    Write-Error $_
    if ($DnsServer -ne "localhost") {
        Write-Error "It looks like we're trying to connect to a remote DNS server. Make sure PowerShell Remoting is enabled by running 'Enable-PSRemoting' on the target."
    }
    exit
}

# variable to store the hosts to add to DNS
[array] $hosts = @()

# read input file
if ($ExcelFile) {
    $hosts = readExcel $ExcelFile
} elseif ($CsvFile) {
    $hosts = readCSV $CsvFile
} else {
    Write-Error "Please specify an Excel or CSV file!"
    exit
}

# one of the file vars will be empty, showing only one on the console
Write-Host "Attempting to add $($hosts.Count) host(s) from '$ExcelFile$CsvFile'..."

# added records counters
$addedRecords = @{
    a = 0
    ptr = 0
    cname = 0
}

# foreach host
foreach ($h in $hosts) {
    # write info for the current host to the console
    Write-Host "Adding '$($h.name)' IP: '$($h.ip)' FQDN: '$($h.fqdn)' CNAME: '$($h.cname)' to DNS..."

    # split the fqdn
    $fqdnParts = $h.fqdn -split '\.'

    # try to get the fqdn host (this will also validate the input, move along on fail)
    if (-not ($fqdnHost = getFwdHost $fqdnParts )) {
        Write-Error "FQDN validation failed"
        continue
    }

    # get fqdn zone
    $fqdnZone = getFwdZone $fqdnParts

    # debug info
    Write-Debug "FQDN Host: '$fqdnHost'  Zone: '$fqdnZone'"

    # check fqdn zone, move along on fail
    if (-not $(checkZone $fqdnZone)) {
        continue
    }

    # split the ip
    $ipParts = $h.ip -split '\.'

    # try to get the ptrHost (this will also validate the input, move along on fail)
    if (-not ($ptrHost = getPtrHost $ipParts)) {
        Write-Error "IP address validation failed"
        continue
    }

    # get ptrZone
    $ptrZone = getPtrZone $ipParts

    # debug info
    Write-Debug "PTR Host: '$ptrHost'  Zone: '$ptrZone'"

    # check ptr zone, move along on fail
    if (-not $(checkZone $ptrZone)) {
        continue
    }

    # if there's a cname for this host
    if ($h.cname) {
        # split the cname
        $cnameParts = $h.cname -split '\.'

        # try to get the cname host (this will also validate the input, move along on fail)
        if (-not ($cnameHost = getFwdHost $cnameParts)) {
            Write-Error "CNAME validation failed"
            continue
        }

        # get the cname zone
        $cnameZone = getFwdZone $cnameParts

        # debug info
        Write-Debug "CNAME Host: '$cnameHost'  Zone: '$cnameZone'"

        # check cname zone, move along on fail
        if (-not $(checkZone $cnameZone)) {
            continue
        }
    }

    #
    # add a record
    #

    # lookup and remove old a records in this zone, move along on fail
    if (-not (lookupAndRemoveOldRecords $fqdnHost $fqdnZone "A")) {
        continue
    }

    # there could be a cname record with this name too, need to remove that
    if (-not (lookupAndRemoveOldRecords $fqdnHost $fqdnZone "Cname")) {
        continue
    }

    # before we add this a record, check if other records exist with this ip
    if (-not (lookupARecordsByIP $h.ip $fqdnZone)) {
        continue
    }

    # try adding a record
    try {
        maybeDo "Add-DnsServerResourceRecordA -ComputerName $DnsServer -ZoneName $fqdnZone -Name $fqdnHost -IPv4Address $($h.ip) -ErrorAction 'Stop'"
    } catch {
        Write-Error "Failed adding A record"
        Write-Error $_
        # move along
        continue
    }

    #
    # add ptr record
    #

    # lookup and remove old ptr records in this zone, move along on fail
    if (-not (lookupAndRemoveOldRecords $ptrHost $ptrZone "Ptr")) {
        continue
    }

    # try adding ptr record
    try {
        maybeDo "Add-DnsServerResourceRecordPtr -ComputerName $DnsServer -ZoneName $ptrZone -Name $ptrHost -PtrDomainName $fqdnHost.$fqdnZone -ErrorAction 'Stop'"
    } catch {
        Write-Error "Failed adding PTR record"
        Write-Error $_
        # move along
        continue
    }

    # if there's a cname for this host
    if ($h.cname) {
        #
        # add cname record - don't need to remove old cname records (will overwrite existing record)
        #

        # try adding cname record
        try {
            maybeDo "Add-DnsServerResourceRecordCName -ComputerName $DnsServer -ZoneName $cnameZone -Name $cnameHost -HostNameAlias $fqdnHost.$fqdnZone -ErrorAction 'Stop'"
        } catch {
            Write-Error "Failed adding CNAME record"
            Write-Error $_
            # move along
            continue
        }
    }

    # debug
    Write-Debug "Success!"

    # increment records counter
    # $addedRecords++
}

# how many hosts did we add?
Write-Host "Added $addedRecords records to DNS"

# calculate delta between hosts we tried to add and hosts we actually added
#$hostAddDelta = $hosts.Count - $addedRecords

# was there a delta?
if ($hostAddDelta -gt 0) {
    Write-Warning "$hostAddDelta host(s) were skipped due to problems"
}

# call function to cleanly exit
CleanExit
