<#
This script file needs to be encoded as UTF-8 (UTF8) when saved
#>
param (
    [Parameter(Mandatory = $true)]
    [string]
    $server,
    [Parameter(Mandatory = $true)]
    [string]
    $database,
    [Parameter(Mandatory = $true)]
    [string]
    $binFile,
    [Parameter(Mandatory = $true)]
    [string]
    $version,
    [Parameter(Mandatory = $true)]
    [string]
    $baseName
)

# :: if required files don't exist or are an old version then create (overwrite) the files

$pbitNotExist = -not (Test-Path -Path "$env:TEMP\$baseName.pbit" -PathType Leaf)
$xlsxNotExist = -not (Test-Path -Path "$env:TEMP\$baseName.xlsx" -PathType Leaf)
$txtNotExist = -not (Test-Path -Path "$env:TEMP\$baseName.txt" -PathType Leaf)
$versionOld = $version -gt (Get-Content -Encoding UTF8 -Path "$env:TEMP\$baseName.txt" -ErrorAction SilentlyContinue)

if ( $pbitNotExist -or $xlsxNotExist -or $txtNotExist -or $versionOld ) {

    Add-Type -Assembly 'System.IO.Compression.FileSystem'

    # :: expand files from archive

    $zip = [System.IO.Compression.ZipFile]::Open($binFile, 'read')
    $entry = $zip.GetEntry("$baseName.pbit"); [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, "$env:TEMP\$($entry.Name)", $true)
    $entry = $zip.GetEntry("$baseName.xlsx"); [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, "$env:TEMP\$($entry.Name)", $true)
    $zip.Dispose()

    # :: open pbit and load data model to memory

    $zip = [System.IO.Compression.ZipFile]::Open("$env:TEMP\$baseName.pbit", 'read')
    $bytes = [byte[]]::new($zip.GetEntry("DataModelSchema").Length); $zip.GetEntry("DataModelSchema").Open().Read($bytes, 0, $bytes.Length)
    $string = [System.Text.Encoding]::Unicode.GetString($bytes)
    $zip.Dispose()

    # :: replace default source location with the corrected one

    $string = $string.Replace("C:\\Users\\Public\\$baseName.xlsx", ("$env:TEMP\$baseName.xlsx").Replace('\', '\\'))
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($string)

    # :: save data model in memory to pbit

    $zip = [System.IO.Compression.ZipFile]::Open("$env:TEMP\$baseName.pbit", 'update')
    $zip.GetEntry("DataModelSchema").Delete()
    $zip.CreateEntry("DataModelSchema")
    $zip.GetEntry("DataModelSchema").Open().Write($bytes, 0, $bytes.Length)
    $zip.Dispose()

    # :: save version file

    Set-Content -Encoding UTF8 -Path "$env:TEMP\$baseName.txt" -Value $version

}    

Start-Process -FilePath "explorer.exe" -ArgumentList @("$env:TEMP\$baseName.pbit")
