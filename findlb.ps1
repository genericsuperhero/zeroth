# -------------------------------------------------------------------------------------
#
# Script to find load balancers for a list of servers using tracert and a 
# list of load balancer hosts
#
# Note: Use "Set-ExecutionPolicy Unrestricted" to allow script execution
#
# Usage: <scriptpath>\findlb.ps1 [<filepath>\]<filename>
#  e.g.: .\findlb.ps1 servers.txt
#
# Input: <scriptpath>\hosts.lb
#
# Output: [<filepath>\]<filename>.lb.csv
#         [<filepath>\]<filename>.lb.log
#
# -------------------------------------------------------------------------------------

# Constants
# ---------
$maxhops = 15
$timeout = 200

# Function to get computer IP address
# -----------------------------------

Function GetComputerIPAddresses
{
    # Get IP address from Network Adapter Configuration
    $adapters =  Get-WmiObject Win32_NetworkAdapterConfiguration -Filter IPEnabled=True
    
    $IPAddresses = "Unknown-IP"
    $IPAddresses = [string]$adapters.IPAddress

    # Check for multiple IP addresses
    If ($IPAddresses -eq "") 
    {
        ForEach ($adapter in $adapters)
        {
            $IPAddresses += [string]$adapter.IPAddress + " "
        }
    }
    $IPAddresses = [String]$IPAddresses.Trim() -Replace(" ", ",")
    
    # Remove IPv6 addresses
    $IPAddressArray = $IPAddresses.Split(",")
    $IPAddresses = ""
    ForEach ($IPAddress in $IPAddressArray)
    {
            If ( !$IPAddress.Contains(":") )
            {
                $IPAddresses += $IPAddress + " "
            }
    }
    $IPAddresses = [String]$IPAddresses.Trim() -Replace(" ", ",")
    
    Return ([String]$IPAddresses)
}

# Function to get destination name and IP address
# -----------------------------------------------

Function GetDestination ([String]$dest)
{
    # Determine the hostname and IP address
    If ( [System.Net.IPAddress]::TryParse($dest, [ref] $null) )
    {
        $destIP = $dest
        $destname = "Unknown-Name"
        $ErrorActionPreference = "SilentlyContinue"
        $destname = [System.Net.Dns]::GetHostByAddress($dest).Hostname
        $ErrorActionPreference = "Continue"
    }
    Else
    {
        $destname = $dest
        $destIP = "Unknown-IP"
        $ErrorActionPreference = "SilentlyContinue"
        $destIP = ([System.Net.Dns]::GetHostAddresses($dest) | select IPAddressToString).IPAddressToString
        $ErrorActionPreference = "Continue"
        
        # Check for multiple DNS entries for the destination
        If ($destIP -eq $null) {
            $destIP = "Multiple-IP"
        }
    }
    
    Return ([String]$destname, [String]$destIP)
}

# Function to check hosts file for IP in trace
# -------------------------------------------

Function CheckHostsfile ([Boolean]$foundlb, [String]$result)
{
    # For each line in the hosts file
    ForEach ($hostsline in $hostsfile)
    {
        # Ignore comments and short lines in the hosts file
        If ( !($hostsline.StartsWith("#")) -and ($hostsline.Length -gt 10) )
        {
            # Parse host ip and name in hosts file
            $hostsIP = $hostsline.Trim().Split("`t ")[0]
            $hostsname = $hostsline.Trim().Split("`t ")[-1]

            # Match load balancer IP in the trace
            $traceIP = $traceline.Trim().Split(" ")[-1]
            If ( $traceIP -eq $hostsIP )
            {
                $foundlb = $True
                $result += " -x-> $hostsname [$hostsIP]"
               
                # write the per load balancer output
                $string = '"' + $linecount + '","' + $sourcename + '","' + $sourceIP + '","' + $hostsname + '","' + $hostsIP + '","' + $destname + '","' + $destIP + '"'
                $string | Out-File -FilePath $outputcsvfile -Append
            }
        }
    }
    
    Return ([Boolean]$foundlb, [String]$result)
}

# Function to check trace for load balancers
# -------------------------------------

Function CheckTrace ([Boolean]$unreachable, [Boolean]$maxhoptrace, [Boolean]$foundlb, [String]$result)
{   
    # For each line in the trace
    ForEach ($traceline in $trace)
    {
        # Check for Destination host unreachable
        If ( $traceline.Contains("Destination host unreachable") )
        {
            $unreachable = $True
            Write-Host -ForegroundColor "yellow" "Warning: Destination host unreachable"
        }
        Else
        {                              
            # Ignore short lines and lines that begin with "Tracing"
            If ( ($traceline.Length -gt 2) -and !($traceline.StartsWith("Tracing")) )
            {
                $hopcount += 1
                                        
                If ($hopcount -gt $maxhops)
                {
                    $maxhoptrace = $True
                    Write-Host -ForegroundColor "yellow" "Warning: Trace reached maximum hops"
                }
                Else
                {
                    # Call function to check hostsfile
                    $return = CheckHostsfile $foundlb $result
                    $foundlb = [Boolean]$return[0]
                    $result = [String]$return[1]
                }
            }
        }
    }
    
    Return ([Boolean]$unreachable, [Boolean]$maxhoptrace, [Boolean]$foundlb, [String]$result)
}
        
# Main program
# ------------

# Check arguements
If ($Args.Count -lt 1)
{
    Write-Host -ForegroundColor Red "Usage: <scriptpath>\findlb.ps1 [<filepath>\]<filename>"
    Write-Host -ForegroundColor Red " e.g.: .\findlb.ps1 servers.txt"
    Exit
}

# Test to ensure the input file exists
$inputfilename = $Args[0]
If (!(Test-Path $inputfilename))
{
    Write-Host "$inputfilename was not found"
    Exit
}

# Load the input file contents
$inputfile = Get-Content -Path $inputfilename

# Specify the load balancer hosts file
$hostsfilename = $MyInvocation.MyCommand.Path.Replace($MyInvocation.MyCommand,"hosts.lb")

# Test to ensure the hosts file exists
If (!(Test-Path $hostsfilename))
{
    Write-Host "$hostsfilename was not found"
    Exit
}

# Load the hosts file contents
$hostsfile = Get-Content -Path $hostsfilename

# Create the output log file
$outputlogfile = $inputfilename + ".lb.log"
Write-Host "Log file: $outputlogfile"
Out-File -InputObject "Find Load Balancer Log" -FilePath $outputlogfile
Get-Date | Out-File -FilePath $outputlogfile -Append

# Create the output csv file
$outputcsvfile = $inputfilename + ".lb.csv"
Write-Host "CSV file: $outputcsvfile"
$string = '"Line","Source","Source IP","Load balancer","Load balancer IP","Destination","Destination IP"'
$string | Out-File -FilePath $outputcsvfile

# Get the computer name and IP address
$sourcename = $env:ComputerName
$sourceIP = GetComputerIPAddresses
Write-Host "Source: $sourcename [$sourceIP]"

$linecount = 0

# For each line in the input file
ForEach ($dest in $inputfile)
{
    $linecount += 1
    Write-Host "Line[$linecount]: $dest"
    
    # Ignore anything after first whitespace
    $dest = $dest.Trim().Split("`t ")[0]

    # Ignore comments and short lines in the input file
    If ( !($dest.StartsWith("#")) -and ($dest.Length -gt 1) )
    {  
        # Get destination name and IP address
        $return = GetDestination $dest
        $destname = [String]$return[0]
        $destIP = [String]$return[1]

        # process if destination is known
        If ( ($destIP -ne "Unknown-IP") -and ($destIP -ne "Multiple-IP") )
        {
            # Tracert to the host
            $command = "tracert -d -h $maxhops -w $timeout"
            Write-Host "Running: $command $destIP"
            $trace = Invoke-Expression "$command $destIP"

            $hopcount = 0
            $maxhoptrace = $False
            $unreachable = $False
            $foundlb = $False
            
            $result = "$sourcename [$sourceIP]"
            
            # Call function to check trace
            $return = CheckTrace $unreachable $maxhoptrace $foundlb $result
            $unreachable = [Boolean]$return[0]
            $maxhoptrace = [Boolean]$return[1]
            $foundlb = [Boolean]$return[2]
            $result = [String]$return[3]
            
            # Format the output for success condition
            $arrow = "--->"
            $color = "green"

            # Format the output for warning conditions
            If ($maxhoptrace -or $unreachable)
            {
                $arrow = "-?->"
                $color = "yellow"
            }
            
            # Format the output for found load balancer
            If ($foundlb)
            {
                $color = "red"
            }

            # Display the string
            $string = "$result $arrow $destname [$destIP]"
            Write-Host -ForegroundColor $color "Line[$linecount]: $string"
        }
        Else
        {
            # Construct the output string for unknown destination
            $arrow = "-?->"
            $color = "yellow"
            
            # Display the output string
            $string = "$sourcename [$sourceIP] $arrow $destname [$destIP] - $destname did not resolve"
            Write-Host -ForegroundColor $color "Line[$linecount]: $string"
        }
        
        # Write the output string to the file
        $string = "Line[$linecount]: $string"
        $string | Out-File -FilePath $outputlogfile -Append
    }
}
