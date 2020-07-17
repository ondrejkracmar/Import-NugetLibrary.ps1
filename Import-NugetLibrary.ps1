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
# For the purpose of (verbose) output
Push-Location $Destination

########################
# The PowerShell Team is breaking my heart with this, but this command is not at all reliable
# As a result, we require nuget.exe 4 or 5 even though the output from that is completely un-parseable
######
# $Package = Install-Package -Name $Name -Destination $Destination -ProviderName NuGet -Source https://api.nuget.org/v3/index.json -ForceBootstrap -Verbose:$False # Trust me, you don't want to Install-Package Verbose with the nuget provider
########################
if ( -not ((Get-Command Nuget).Version.Major -ge 4)) {
    throw "In order to use this script you must have nuget.exe with a version greater than 4.0 in your path"
}
Write-Progress -Activity "Nuget Install" -Status ">> nuget install $Name -OutputDirectory '$Destination' -PackageSaveMode nuspec -ForceEnglishOutput" -Id 1
nuget install $Name -OutputDirectory $Destination -PackageSaveMode nuspec | ForEach-Object {
    if ($_) {
        Write-Information $_
        if ($InformationPreference -ne "Continue") {
            Write-Progress -Activity "Information" -Status $_ -ParentId 1
        }
    }
}
Write-Progress -Activity "Information" -Completed
Write-Progress -Activity "Nuget Install" -Completed

# Nuget packages hide their assemblies in folders with version numbers...
$Frameworks = ".NETFramework4.5", ".NETStandard2.1", ".NETStandard2.0"
if ($PSVersionTable.PSVersion -ge "6.0") {
    $Versions = "netstandard2.1", "netstandard2.0",
    "netstandard1.6", "netstandard1.5", "netstandard1.4", "netstandard1.3", "netstandard1.2", "netstandard1.1", "netstandard1.0",
    "net472", "net471", "net47", "net463", "net462", "nt461", "net46", "net452", "net451", "net45", "net40", "net35", "net20"
    $Frameworks = ".NETStandard2.1", ".NETStandard2.0", ".NETFramework4.5"
} elseif ($PSVersionTable.PSVersion -ge "4.0") {
    $Versions = "net472", "net471", "net47", "net463", "net462", "nt461", "net46", "net452", "net451", "net45", "net40", "net35", "net20"
} elseif ($PSVersionTable.PSVersion -ge "3.0") {
    $Versions = "net40", "net35", "net20", "net45", "net451", "net452", "net46"
} else {
    $Versions = "net35", "net20"
}

filter Import-Dependency {
    [CmdletBinding()]
    param(
        $RootFolder = $Destination,

        [Parameter(ValueFromPipelineByPropertyName, Mandatory)]
        [Alias("Name")]
        $Id,

        [Parameter(ValueFromPipelineByPropertyName)]
        $Version,

        [string[]]$Frameworks = ".NETStandard2.1"
    )
    $Package = if (Test-Path "$RootFolder\$Id*") {
        (Get-ChildItem "$RootFolder\$Id*") | ForEach-Object {
            $PackageName, $PackageVersion = $_.Name -split "(?<=\D)\.(?=\d+)"
            if (-not $Version -or $PackageVersion -eq $Version) {
                [PSCustomObject]@{
                    Name    = $PackageName
                    Version = $PackageVersion
                    Path    = $_.PSPath
                }
            }
        } | Sort-Object Version -Descending | Select-Object -First 1
    } else {
        Write-Error "Could not find package '$Id' in '$RootFolder' after install"
    }

    Write-Verbose "Found $Id in $($Package.Path)"

    # build full path with \ on the end
    $LibraryPath = ($Versions -replace "^", "$($Package.Path)*\lib\") + "$($Package.Path)*\lib\" + "$($Package.Path)*" |
        # find the first one that exists
        Convert-Path -ErrorAction SilentlyContinue | Select-Object -First 1

    $Number = $LibraryPath -replace '.*?([\d\.]+)$', '$1' -replace '(?<=\d)(?=\d)$', '.'

    if (Test-Path "$($Package.Path)\$Id.nuspec") {
        Write-Verbose "Reading `$NugetPackageData from '$($Package.Path)\$($Id).nuspec')"
        $global:Metadata = ($global:NugetPackageData = ([xml](Get-Content "$($Package.Path)\$Id.nuspec"))).package.metadata
    } elseif (Test-Path "$($Package.Path)\$Id.nupkg") {
        try {
            Write-Verbose "Reading `$NugetPackageData from '$($Package.Path)\$($Id).nupkg')"
            $Package = [System.IO.Packaging.Package]::Open( "$($Package.Path)\$Id.nupkg", "Open", "Read" )
            $Part = $Package.GetPart( $Package.GetRelationshipsByType( "http://schemas.microsoft.com/packaging/2010/07/manifest" ).TargetUri )

            try {
                $Stream = $Part.GetStream()
                $Reader = [System.IO.StreamReader]$Stream
                $global:Metadata = ($global:NugetPackageData = ([xml]$Reader.ReadToEnd())).package.metadata
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
    } else {
        Write-Warning "Package not found in '$($Package.Path)\$Id.nu*'"
    }

    if ($Metadata.dependencies) {
        if ($Framework) {
            if ($Dependencies = @($Metadata.dependencies.group.where( { $_.targetFramework -eq $Framework }, "last", 1).dependency)) {
                $Dependencies | Import-Dependency -RootFolder $RootFolder -Framework $Framework
                break
            }
        } else {
            foreach ($Framework in $Frameworks) {
                if ($Dependencies = @($Metadata.dependencies.group.where( { $_.targetFramework -eq $Framework }, "last", 1).dependency)) {
                    $Dependencies | Import-Dependency -RootFolder $RootFolder -Framework $Framework
                    break
                }
            }
        }
    }
    if (!$Dependencies) {
        Write-Verbose "No dependencies for $Id"
    }


    # If there's no references node, this is an old package, reference everything inside it, and hope for luck
    if (!$Metadata.References) {
        Write-Verbose "No references found. See `$Metadata"
        $Assemblies = Get-ChildItem $LibraryPath -Filter *.dll
    } else {
        Write-Verbose "Need to reference $($Metadata.References.Group.reference.File -join ', ')"
        $group = $Metadata.References.Group | Where-Object { $_.targetFramework.EndsWith($number) }
        if ($group) {
            $Files = $group.reference.File
        } else {
            # If we can't figure out the right group, just get all the references:
            $Files = @($Metadata.References.SelectNodes("//*[name(.) = 'reference']").File | Select-Object -Unique)
        }
        $Assemblies = Get-Item ($Files -replace "^", "$LibraryPath\")
    }

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
                Write-Host "$(Resolve-Path $LibraryPath -Relative)\" -Fore Yellow -NoNewline
                Write-Host $Assm.Name -Fore White
            }
        }
        # if we loaded everything ok, we're done
        if ($success) {
            break
        }
    }
    if (!$success) {
        throw $failure
    }
}

Import-Dependency -RootFolder $Destination -Id $Name

Pop-Location

return

