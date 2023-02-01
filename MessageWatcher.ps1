### Settings
$computers = @()
$groups = @()
$ous = @()

## File Watcher
$Path = "C:\MessageAlert\Messages"
$HistoryPath = "C:\MessageAlert\MessageHistory"
$IncludeSubfolders = $true
$AttributeFilter = [IO.NotifyFilters]::FileName
$ChangeTypes = [System.IO.WatcherChangeTypes]::Created
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
        [System.IO.WaitForChangedResult]
        $ChangeInformation
    )
    Write-Warning 'Change detected:'
    $ChangeInformation | Out-String | Write-Host -ForegroundColor DarkYellow

    ## Read the message file
    $changedFilePath = Join-Path $Path $ChangeInformation.Name
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
    Write-Warning "FileSystemWatcher is monitoring $Path"
    
    # create a filesystemwatcher object
    $watcher = New-Object -TypeName IO.FileSystemWatcher -ArgumentList $Path, $FileFilter -Property @{
        IncludeSubdirectories = $IncludeSubfolders
        NotifyFilter = $AttributeFilter
    }

    # start monitoring manually in a loop:
    do {
        # wait for changes for the specified timeout
        # IMPORTANT: while the watcher is active, PowerShell cannot be stopped
        # so it is recommended to use a timeout of 1000ms and repeat the
        # monitoring in a loop. This way, you have the chance to abort the
        # script every second.
        $result = $watcher.WaitForChanged($ChangeTypes, $Timeout)
        # if there was a timeout, continue monitoring:
        if ($result.TimedOut) { continue }
        
        action -Change $result
        # the loop runs forever until you hit CTRL+C    
    } while ($true)
} finally {
    # release the watcher and free its memory:
    $watcher.Dispose()
    Write-Warning 'FileSystemWatcher removed.'
}