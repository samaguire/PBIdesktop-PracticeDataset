$baseName = "practicedataset"
$basePath = Split-Path -parent $PSCommandPath

# :: set bulk of json values

$version = "2.3.4"
$name = "Practice Dataset"
$description = "Practice Dataset"
$path = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe"

# :: get json value for arguments

$ps1 = {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $server,
        [Parameter(Mandatory = $true)]
        [string]
        $database
    )
    $binFileMain = "${env:CommonProgramFiles(x86)}\Microsoft Shared\Power BI Desktop\External Tools\$baseName.pbitool.bin"
    $binFileAlt = "$env:CommonProgramFiles\Microsoft Shared\Power BI Desktop\External Tools\$baseName.pbitool.bin"
    if (Test-Path -Path $binFileMain -PathType Leaf) {
        $binFile = $binFileMain
    }
    elseif (Test-Path -Path $binFileAlt -PathType Leaf) {
        $binFile = $binFileAlt
    }
    else {
        exit
    }
    Add-Type -Assembly 'System.IO.Compression.FileSystem'
    $zip = [System.IO.Compression.ZipFile]::Open($binFile, 'read')
    $bytes = [byte[]]::new($zip.GetEntry("$baseName.ps1").Length); $zip.GetEntry("$baseName.ps1").Open().Read($bytes, 0, $bytes.Length) | Out-Null
    $scriptBlock = [Scriptblock]::Create([System.Text.Encoding]::UTF8.GetString($bytes))
    $zip.Dispose()
    Invoke-Command -ScriptBlock $scriptBlock -ArgumentList @($server, $database, $binFile, '$version', '$baseName')
}
$ps1String = $ps1.ToString().Replace('$baseName', $baseName).Replace('$version', $version)
$ps1Base64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($ps1String))
$command = { Invoke-Command -ScriptBlock ([Scriptblock]::Create([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($ps1Base64)))) -ArgumentList @('%server%', '%database%') }
$commandString = $command.ToString().Replace('$ps1Base64', "'$ps1Base64'")
$arguments = "-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -Command &{$commandString}"

# :: get json value for iconData

$imageType = "png"
$imageBase64 = [System.Convert]::ToBase64String((Get-Content -Raw -Encoding Byte -Path "$basePath\resources\$baseName.$imageType"))
$iconData = "data:image/$imageType;base64,$imageBase64"

# :: create json file

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

Set-Content -Encoding UTF8 -Path "$basePath\$baseName.pbitool.json" -Value $json

# :: test command that gets launched by powershell
&$command
