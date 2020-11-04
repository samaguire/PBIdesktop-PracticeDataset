$baseName = "practicedataset"
$basePath = Split-Path -parent $PSCommandPath

Add-Type -Assembly 'System.IO.Compression.FileSystem'
$compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal

# create zip

Remove-Item -Path "$basePath\$baseName.pbitool.bin" -ErrorAction SilentlyContinue
$zip = [System.IO.Compression.ZipFile]::Open("$basePath\$baseName.pbitool.bin", 'create')
$zip.Dispose()

# update zip

$zip = [System.IO.Compression.ZipFile]::Open("$basePath\$baseName.pbitool.bin", 'update')
$ext = "ps1"; [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, "$basePath\resources\$baseName.$ext", "$baseName.$ext", $compressionLevel)
$ext = "pbit"; [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, "$basePath\resources\$baseName.$ext", "$baseName.$ext", $compressionLevel)
$ext = "xlsx"; [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, "$basePath\resources\$baseName.$ext", "$baseName.$ext", $compressionLevel)
$zip.Dispose()
