$baseName = "practicedataset"

# Set bulk of json values
$version = "3.0.0"
$name = "Practice Dataset"
$description = "Practice Dataset"
$path = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe"

# Get json value for arguments
$ps1 = {
    param (
        [Parameter(Mandatory = $true)]
        [string]$server,
        [Parameter(Mandatory = $true)]
        [string]$database
    )
    
    # tmp variables
    # $baseName = "practicedataset"
    
    # Check current environment
    $ver = Get-Content "$env:TEMP\$baseName.ver" -ErrorAction SilentlyContinue
    $pbit = -not (Test-Path "$env:TEMP\$baseName.pbit" -PathType Leaf)
    $xlsx = -not (Test-Path "$env:TEMP\$baseName.xlsx" -PathType Leaf)
    
    # Get latest release details https://docs.github.com/en/free-pro-team@latest/rest/reference/repos#releases
    $repo = "samaguire/PBIdesktop-PracticeDataset"
    # $releases = Invoke-WebRequest "https://api.github.com/repos/$repo/releases" | ConvertFrom-Json
    $releases = Invoke-WebRequest "https://api.github.com/repos/$repo/releases/latest" | ConvertFrom-Json # is the most recent non-prerelease, non-draft release
    $tag = $releases[0].tag_name
    $zipurl = $releases[0].zipball_url
    
    # Download latest version if newer than the current version or required files are missing and update the pbit
    if ($tag -gt $ver -or $pbit -or $xlsx) {
    
        # Clear existing files
        Get-ChildItem $env:TEMP -Filter "$baseName.*" | Remove-Item -Recurse
    
        # Download latest version
        $file = "$env:TEMP\$baseName.zip"
        If (Test-Path $file) { Remove-Item $file }
        Invoke-WebRequest $zipurl -Out $file
    
        # Extract latest version
        Expand-Archive $file -DestinationPath "$env:TEMP\$baseName"
        Get-ChildItem "$env:TEMP\$baseName" -Recurse -Filter "$baseName.xlsx" | Move-Item -Destination "$env:TEMP"
        Get-ChildItem "$env:TEMP\$baseName" -Recurse -Filter "$baseName.pbit" | Move-Item -Destination "$env:TEMP"
    
        # Open pbit and load data model to memory
        Add-Type -Assembly 'System.IO.Compression.FileSystem'
        $zip = [System.IO.Compression.ZipFile]::Open("$env:TEMP\$baseName.pbit", 'read')
        $bytes = [byte[]]::new($zip.GetEntry("DataModelSchema").Length); $zip.GetEntry("DataModelSchema").Open().Read($bytes, 0, $bytes.Length) | Out-Null
        $string = [System.Text.Encoding]::Unicode.GetString($bytes)
        $zip.Dispose()
    
        # Replace default source location with the corrected one
        $string = $string.Replace("C:\\Users\\Public\\$baseName.xlsx", ("$env:TEMP\$baseName.xlsx").Replace('\', '\\'))
        $bytes = [System.Text.Encoding]::Unicode.GetBytes($string)
    
        # Save data model in memory to pbit
        $zip = [System.IO.Compression.ZipFile]::Open("$env:TEMP\$baseName.pbit", 'update')
        $zip.GetEntry("DataModelSchema").Delete()
        $zip.CreateEntry("DataModelSchema") | Out-Null
        $zip.GetEntry("DataModelSchema").Open().Write($bytes, 0, $bytes.Length)
        $zip.Dispose()
    
        # Save version file
        Set-Content "$env:TEMP\$baseName.ver" -Value $tag
    
    }
    
    # Launch
    Start-Process -FilePath "explorer.exe" -ArgumentList @("$env:TEMP\$baseName.pbit")
}
$ps1String = $ps1.ToString().Replace('$baseName', $baseName).Replace('$version', $version)
$ps1Base64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($ps1String))
$command = { Invoke-Command -ScriptBlock ([Scriptblock]::Create([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($ps1Base64)))) -ArgumentList @('%server%', '%database%') }
$commandString = $command.ToString().Replace('$ps1Base64', "'$ps1Base64'")
# $arguments = "-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -Command &{$commandString}"
$arguments = "-NoProfile -ExecutionPolicy Bypass -Command &{$commandString}"

# Get json value for iconData
$imageType = "png"
$imageBase64 = [System.Convert]::ToBase64String((Get-Content -Raw -Encoding Byte "$PSScriptRoot\resources\$baseName.$imageType"))
$iconData = "data:image/$imageType;base64,$imageBase64"

# Create json file
$json = @"
{
  "version": "$version",
  "name": "$name",
  "description": "$description",
  "path": "$path",
  "arguments": "$arguments",
  "iconData": "$iconData"
}
"@
Set-Content "$PSScriptRoot\$baseName.pbitool.json" -Value $json

# Test command that gets launched by powershell
&$command
