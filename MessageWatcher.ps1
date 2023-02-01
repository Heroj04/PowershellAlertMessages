### Settings
$computers = @()
$groups = @()
$ous = @()
$Path = "C:\MessageAlert\Messages"
$HistoryPath = "C:\MessageAlert\MessageHistory"
$Timeout = 1000
$refreshTimeout = 900000

$lastRefresh = Get-Date
$allComputers = @()

### Expand Groups and OUs
function refreshAllComputers {
    $return = @()
    $return += $computers
    foreach ($groupName in $groups) {
        $groupMembers = Get-ADGroupMember $groupName -Recursive | ForEach-Object {$_.Name}
        foreach ($member in $groupMembers) {
            if ($return -notcontains $member) {
                $return += $member
            }
        }
    }
    foreach ($ouString in $ous) {
        $ouMembers = Get-ADComputer -Filter * -SearchBase $ouString | ForEach-Object {$_.Name}
        foreach ($member in $ouMembers) {
            if ($return -notcontains $member) {
                $return += $member
            }
        }
    }
    Write-Host "Refreshed Computer List" -BackgroundColor Black -ForegroundColor White
    return $return
}
$allComputers = refreshAllComputers

### DEFINE ACTIONS AFTER AN EVENT IS DETECTED
function action {
    param
    (
        [Parameter(Mandatory)]
        [String]
        $changedFilePath
    )
    $content = Get-Content -Path $changedFilePath -Raw
    Write-Output "Processing Message ($changedFilePath):`n$content"

    ## Send Message to each computer
    foreach ($computerName in $allComputers) {
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

        Start-Sleep -Milliseconds $Timeout

        if ($lastRefresh.AddMilliseconds($refreshTimeout) -lt (Get-Date)) {
            $allComputers = refreshAllComputers
            $lastRefresh = Get-Date
        }

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