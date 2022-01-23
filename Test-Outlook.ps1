Function ConvertFrom-Hexa {
    <#
    .SYNOPSIS
    This function is designed to convert a hex value to a string

    .DESCRIPTION
    This function is designed to convert a hex value to a string

    .EXAMPLE
    Move
    
    .INPUTS
    None. This function does not support piped input. # <-- Example (when the function does not accept piped input)

    .OUTPUTS
    Task status: Unknown, Disabled, Queued, Ready or Running

    .NOTES
    No exceptions are returned by this function

    .LINK
    None
    #>

    param (
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $false,
            ValueFromPipelineByPropertyName = $false,
            ValueFromRemainingArguments = $false,
            HelpMessage = "The hex value that is to-be converted to a string"
        )]
        [ValidateNotNullOrEmpty()]
        [string] $HexString
    )

    ($hexstring.Split(",",[System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object {$_ -gt '0'} | ForEach-Object {[char][int]"$($_)"}) -join ''

}

Function Remove-Object {
    param (
        $Path
    )

	[PSCustomObject] $local:Results = New-Object PSObject -Property @{
        Status = $null
        Details = $null
    }

    try {
        Remove-Item $Path -Force -Confirm:$false -Recurse -ErrorAction Stop -WhatIf
		$local:Results.Status = "Success"
		Write-Host $Path
    } catch {
        if($_ -like "*does not exist*") {
            $local:Results.Status = "Success"
            Write-Host "$Path Error: $_"
        } else {
            $local:Results.Status = "Failure"
            Write-Host "$Path Error: $_"
        }
    }

    Write-Host "$($local:Results.Details)"
	Return $local:Results
}

$script:sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value

[PSCustomObject] $local:OutlookProfiles  = ((Get-ChildItem ("Registry::HKEY_USERS\{0}\Software\Microsoft\Office\16.0\Outlook\Profiles" -f $script:sid) -Recurse ))

# Local variables
[string] $local:FoundAddress = [string]::Empty
[array] $local:OSTLocations = @()
[int] $local:numPSTsFound = 0

# Capture and log all configured Outlook profile details
foreach($local:Key in $local:OutlookProfiles) {

    # Capture all Outlook profiles that are configured
    [PSCustomObject] $local:KeyProps = ((Get-Item -path "REGISTRY::$local:Key").property)

    # Foreach profile found, capture its associated information
    foreach($local:SubKey in $local:KeyProps) {

        # Capture email addresses of the configured Outlook profiles
        if($local:SubKey -eq "Account Name") {

            # Locate the property that contains the profile's email address and log it
            $local:FoundAddress = Get-ItemProperty -path "REGISTRY::$local:Key" | Select-Object -ExpandProperty $local:SubKey 
            if($local:FoundAddress -like "*@*") {

                #Output email address to file and log it
                if(-not ([string]::IsNullOrEmpty($local:FoundAddress))) {
                    # $local:FoundAddress | Out-File (Get-UserLogFilePath -Name "OutlookProfiles") -Append
                }

                $local:logAddresses += $local:FoundAddress
            }
        }
        if($local:SubKey -eq "001f6700") {

            # Local variables
            [string] $local:NewProfileHex = [string]::Empty
            [string] $local:NewProfileString = [string]::Empty

            $local:NewProfileHex = (Get-ItemProperty -path "REGISTRY::$local:Key" | Select-Object -ExpandProperty $local:SubKey) -join ","
            $local:NewProfileString = ConvertFrom-Hexa -HexString  $NewProfileHex

            if(-not ([string]::IsNullOrEmpty($local:FoundAddress))) {
                # $local:NewProfileString | Out-File (Get-UserLogFilePath -Name "OutlookPSTs") -Append
            }
            $local:PSTLocations += $NewProfileString
        }
        if($local:SubKey -eq "001f6610") {

            # Local variables
            [string] $local:OSTHex = [string]::Empty
            [string] $local:OSTHexString = [string]::Empty
        
            $local:OSTHex = (Get-ItemProperty -path "REGISTRY::$local:Key" | Select-Object -ExpandProperty $local:SubKey) -join ","
            $local:OSTHexString = ConvertFrom-Hexa -HexString  $OSTHex
            [array] $local:OSTLocations += $OSTHexString
        }
    }
}

$local:OSTLocations

$local:numPSTsFound = ($local:PSTLocations | Measure-Object).count

# Remove old OSTs if in scope
# Local variables
[array] $local:OSTRootFolders = @()
[array] $local:OSTsToRemove = @()
[array] $local:AllOSTRoots = @()

# Capture the root location for OSTs associated with the users configured Outlook profiles
foreach($local:OSTFolder in $local:OSTLocations) {
    [string] $local:OSTName = [string]::Empty
    [string] $local:OSTRoot = [string]::Empty

    [string] $local:OSTName = ($local:OSTFolder -split "\\")[-1]
    [string] $local:OSTRoot = $local:OSTFolder.Trim($local:OSTName)
    [array] $local:AllOSTRoots += [string] $local:OSTRoot
}

$local:AllOSTRoots

# Locate only the unique OST locations
[array] $local:OSTRootFolders = [array] $local:AllOSTRoots | Select-Object -Unique

$local:OSTRootFolders

# Seach those root locations for all OST and NST files
foreach($local:OSTRootFolder in $local:OSTRootFolders) {
    [array] $local:OSTsToRemove += Get-ChildItem $local:OSTRootFolder -Recurse | Where-Object {$_.Name -like "*.ost" -or $_.Name -like "*.nst"}
}

$local:OSTsToRemove

# Delete the old OST
foreach($OSTToRemove in $local:OSTsToRemove) {
    ($OSTToRemove.FullName)
    Remove-Object -Path $OSTToRemove.FullName
}
