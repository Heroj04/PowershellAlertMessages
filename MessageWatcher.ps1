### Settings
$computers = @()
$groups = @()
$ous = @()

## File Watcher
$Path = "C:\MessageAlert\Messages"
$HistoryPath = "C:\MessageAlert\MessageHistory"
$Timeout = 1000

### Expand Groups and OUs
foreach ($groupName in $groups) {
    $groupMembers = Get-ADGroupMember $groupName -Recursive | ForEach-Object {$_.Name}
    foreach ($member in $groupMembers) {
        if ($computers -notcontains $member) {
            $computers += $member
        }
    }
}
foreach ($ouString in $ous) {
    $ouMembers = Get-ADComputer -Filter * -SearchBase $ouString | ForEach-Object {$_.Name}
    foreach ($member in $ouMembers) {
        if ($computers -notcontains $member) {
            $computers += $member
        }
    }
}

### DEFINE ACTIONS AFTER AN EVENT IS DETECTED
function action {
    param
    (
        [Parameter(Mandatory)]
        [String]
        $changedFilePath
    )
    $content = Get-Content -Path $changedFilePath -Raw

    ## Send Message to each computer
    foreach ($computerName in $computers) {
        if (test-connection $computerName -quiet -Count 1) {
            msg.exe * /SERVER:$computerName "$content"
            #Start-Process "msg" -ArgumentList "* /SERVER:$computerName `"$content`""
            Write-Host "Message to $computerName sent" -BackgroundColor Black -ForegroundColor White
        } else {
            Write-Host "$computerName OFFLINE" -BackgroundColor Black -ForegroundColor Red
        }
    }

    ## Delete the message file
    #Remove-Item $changedFilePath
    Move-Item $changedFilePath $HistoryPath -Force
}

### SET Watcher
try {
    Write-Warning "Watcher is monitoring $Path"

    # start monitoring manually in a loop:
    do {

        Start-Sleep $Timeout

        $files = Get-ChildItem -Path $Path -Filter $FileFilter -File -Recurse

        if ($files.Length -gt 0) {
            foreach ($file in $files) {
                action -changedFilePath $file.FullName
            }
        }

    } while ($true)
} finally {
    Write-Warning 'Watcher removed.'
}