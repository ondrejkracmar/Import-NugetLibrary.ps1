#requires -Module PackageManagement
param(
    [Parameter(Mandatory)]
    [String]$Name,

    [String]$Destination = "$(Split-Path $Profile)\Libraries"
)
$ErrorActionPreference = "Stop"
if (!(Test-Path $Destination -Type Container)) {
    throw "The destination path ($Destination) must point to an existing folder, NuGet will install packages in subdirectories of this folder."
}

# Normalize: nuget requires destination NOT end with a slash
$Destination = (Convert-Path $Destination).TrimEnd("\")
Write-Host "Calling PackageManagement\Install-Package '$Name' (this may take a while)..."
$Package = Install-Package -Name $Name -Destination $Destination -ProviderName NuGet -Source 'https://api.nuget.org/v3/index.json' -ForceBootstrap -Verbose:$False # Trust me, you don't want to Install-Package Verbose with the nuget provider

$PackagePath = if (Test-Path "$Destination\$Name") {
    Convert-Path "$Destination\$Name"
} elseif (Test-Path "$Destination\$Name*") {
    Convert-Path "$Destination\$Name*"
} else {
    $Package | Out-Default
    Write-Error "Could not find package after install of Package!"
}

Write-Verbose "Installed $Name to $PackagePath"

# Nuget packages hide their assemblies in folders with version numbers...
if ($PSVersionTable.PSVersion -ge "6.0") {
    $Versions = "netstandard2.1", "netstandard2.0",
        "netstandard1.6", "netstandard1.5", "netstandard1.4", "netstandard1.3", "netstandard1.2", "netstandard1.1", "netstandard1.0",
        "net472", "net471", "net47", "net463", "net462", "nt461", "net46", "net452", "net451", "net45", "net40", "net35", "net20"
} elseif ($PSVersionTable.PSVersion -ge "4.0") {
    $Versions = "net472","net471","net47","net463","net462","nt461","net46", "net452", "net451", "net45", "net40", "net35", "net20"
} elseif ($PSVersionTable.PSVersion -ge "3.0") {
    $Versions = "net40", "net35", "net20", "net45", "net451", "net452", "net46"
} else {
    $Versions = "net35", "net20"
}

# build full path with \ on the end
$LibraryPath = ($Versions -replace "^", "$PackagePath*\lib\") + "$PackagePath*\lib\" + "$PackagePath*" |
    # find the first one that exists
    Convert-Path -ErrorAction SilentlyContinue | Select-Object -First 1

$Number = $LibraryPath -replace '.*?([\d\.]+)$', '$1' -replace '(?<=\d)(?=\d)$', '.'

if (Test-Path "$PackagePath\$Name.nuspec") {
    $References = ([xml](Get-Content "$PackagePath\$Name.nuspec")).package.metadata.references
} elseif (Test-Path "$PackagePath\$Name.nupkg") {
    try {
        $Package = [System.IO.Packaging.Package]::Open( "$PackagePath\$Name.nupkg", "Open", "Read" )
        $Part = $Package.GetPart( $Package.GetRelationshipsByType( "http://schemas.microsoft.com/packaging/2010/07/manifest" ).TargetUri )

        try {
            $Stream = $Part.GetStream()
            $Reader = [System.IO.StreamReader]$Stream
            $References = ([xml]$Reader.ReadToEnd()).package.metadata.references
        } catch [Exception] {
            $PSCmdlet.WriteError( [System.Management.Automation.ErrorRecord]::new($_.Exception, "Unexpected Exception", "InvalidResult", $_) )
        } finally {
            if ($Reader) {
                $Reader.Close()
                $Reader.Dispose()
            }
            if ($Stream) {
                $Stream.Close()
                $Stream.Dispose()
            }
        }
        if ($Package) {
            $Package.Close()
            $Package.Dispose()
        }
    } catch [Exception] {
        $PSCmdlet.WriteError( [System.Management.Automation.ErrorRecord]::new($_.Exception, "Cannot Open Path", "OpenError", $Path) )
    }
}

# If there's no references node, this is an old package, just reference everything we found
if (!$References) {
    $Assemblies = Get-ChildItem $LibraryPath -Filter *.dll
} else {
    $group = $references.Group | Where-Object { $_.targetFramework.EndsWith($number) }
    if ($group) {
        $Files = $group.reference.File
    } else {
        # If we can't figure out the right group, just get all the references:
        $Files = @($references.SelectNodes("//*[name(.) = 'reference']").File | Select -Unique)
    }
    $Assemblies = Get-Item ($Files -replace "^", "$LibraryPath\")
}

# Just for the purpose of the verbose output
Push-Location $Destination
# since we don't know the order, we'll just loop a few times
for ($e = 0; $e -lt $Assemblies.Count; $e++) {
    $success = $true
    foreach ($assm in $Assemblies) {
        Write-Verbose "Import Library $(Resolve-Path $Assm.FullName -Relative)"
        Add-Type -Path $Assm.FullName -ErrorAction SilentlyContinue -ErrorVariable failure
        if ($failure) {
            $success = $false
        } else {
            Write-Host "LOADING: " -Fore Cyan -NoNewline
            Write-Host $LibraryPath\ -Fore Yellow -NoNewline
            Write-Host $Assm.Name -Fore White
        }
    }
    # if we loaded everything ok, we're done
    if ($success) {
        break
    }
}
Pop-Location
if (!$success) {
    throw $failure
}
return

