## MAX CHARACTERS: 253
$message = "******************************
Alert Message
******************************"

$outputDir = "\\Server\MessageAlert\Messages"
$outputFile = Join-Path $outputDir "Alert_$($env:COMPUTERNAME)_$(Get-Date -Format "yyyy-MM-dd_HHmmss")"

$message | Out-File $outputFile -Force