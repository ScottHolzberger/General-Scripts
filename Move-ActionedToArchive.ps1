<#
Move-ActionedToArchive.ps1
Moves messages from one folder to another in a mailbox using Microsoft Graph PowerShell.
Logs all activity to CSV and throttles requests with retry/backoff.

Requires: Microsoft.Graph (Mail)
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$MailboxUPN = "sales@stls.com.au",

    [Parameter(Mandatory=$true)]
    [string]$SourceFolderName = "ACTIONED",

    [Parameter(Mandatory=$true)]
    [string]$TargetFolderName = "Archive",

    [int]$BatchSize = 25,                 # how many messages per batch
    [int]$MaxBatches = 200,               # safety cap: prevents infinite loops
    [int]$DelayMsBetweenMoves = 250,      # base throttle per move
    [int]$DelayMsBetweenBatches = 1000,   # throttle per batch

    [string]$LogPath = ".\ActionedToArchive_MoveLog.csv",

    [switch]$WhatIf
)

# ----------------------------
# Helpers
# ----------------------------

function Write-LogRow {
    param(
        [string]$Status,
        [string]$MessageId,
        [string]$InternetMessageId,
        [string]$Subject,
        [string]$ReceivedDateTime,
        [string]$Error
    )
    $row = [pscustomobject]@{
        Timestamp          = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        MailboxUPN         = $MailboxUPN
        SourceFolder       = $SourceFolderName
        TargetFolder       = $TargetFolderName
        Status             = $Status
        MessageId          = $MessageId
        InternetMessageId  = $InternetMessageId
        Subject            = $Subject
        ReceivedDateTime   = $ReceivedDateTime
        Error              = $Error
    }
    $row | Export-Csv -Path $LogPath -Append -NoTypeInformation -Encoding UTF8
}

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxAttempts = 8,
        [int]$BaseDelaySeconds = 2
    )

    $attempt = 0
    while ($true) {
        try {
            $attempt++
            return & $ScriptBlock
        }
        catch {
            $msg = $_.Exception.Message

            # Heuristic detect throttling/transient errors (429/503/504 show up in message text frequently)
            $isThrottle = ($msg -match "429" -or $msg -match "Too Many Requests" -or $msg -match "throttl" )
            $isTransient = ($msg -match "503" -or $msg -match "504" -or $msg -match "temporar" -or $msg -match "timeout")

            if ($attempt -ge $MaxAttempts -or (-not ($isThrottle -or $isTransient))) {
                throw
            }

            # Exponential backoff with jitter
            $delay = [Math]::Min(90, ($BaseDelaySeconds * [Math]::Pow(2, $attempt - 1)))
            $jitter = Get-Random -Minimum 0 -Maximum 3
            Start-Sleep -Seconds ($delay + $jitter)
        }
    }
}

function Find-FolderIdByDisplayName {
    param(
        [string]$UserId,
        [string]$FolderDisplayName
    )

    function Find-InChildren {
        param(
            [string]$UserId,
            [string]$ParentFolderId,
            [string]$FolderDisplayName
        )

        # Get child folders of a given folder
        $children = Invoke-WithRetry {
            Get-MgUserMailFolderChildFolder -UserId $UserId -MailFolderId $ParentFolderId -All
        }

        foreach ($c in $children) {
            if ($c.DisplayName -eq $FolderDisplayName) { return $c.Id }

            $found = Find-InChildren -UserId $UserId -ParentFolderId $c.Id -FolderDisplayName $FolderDisplayName
            if ($found) { return $found }
        }

        return $null
    }

    # First scan top-level folders
    $top = Invoke-WithRetry { Get-MgUserMailFolder -UserId $UserId -All }

    foreach ($f in $top) {
        if ($f.DisplayName -eq $FolderDisplayName) { return $f.Id }

        $found = Find-InChildren -UserId $UserId -ParentFolderId $f.Id -FolderDisplayName $FolderDisplayName
        if ($found) { return $found }
    }

    throw "Folder '$FolderDisplayName' not found in mailbox '$UserId'."
}


# ----------------------------
# Connect to Graph
# ----------------------------

Import-Module Microsoft.Graph.Mail -ErrorAction Stop

# Use Mail.ReadWrite for your own mailbox; use Mail.ReadWrite.Shared when acting on shared mail you have access to. [4](https://learn.microsoft.com/en-us/graph/outlook-share-messages-folders)[5](https://graphpermissions.merill.net/permission/Mail.ReadWrite.Shared)
$scopes = @("Mail.ReadWrite.Shared")
# If this is a shared mailbox you access via delegation, use the shared scope:
# $scopes = @("Mail.ReadWrite.Shared")

Connect-MgGraph -Scopes $scopes | Out-Null

Write-Host "Connected to Microsoft Graph."

# ----------------------------
# Resolve folders
# ----------------------------

$sourceFolderId = Find-FolderIdByDisplayName -UserId $MailboxUPN -FolderDisplayName $SourceFolderName
$targetFolderId = Find-FolderIdByDisplayName -UserId $MailboxUPN -FolderDisplayName $TargetFolderName

Write-Host "Source folder '$SourceFolderName' Id: $sourceFolderId"
Write-Host "Target folder '$TargetFolderName' Id: $targetFolderId"

# Ensure log file has headers (create if missing)
if (-not (Test-Path $LogPath)) {
    "" | Out-File $LogPath
    # Write a header row by exporting an empty object once
    [pscustomobject]@{
        Timestamp=""; MailboxUPN=""; SourceFolder=""; TargetFolder=""; Status=""; MessageId=""; InternetMessageId=""; Subject=""; ReceivedDateTime=""; Error=""
    } | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
    # Remove the blank row created by header bootstrap
    (Get-Content $LogPath | Select-Object -Skip 2) | Set-Content $LogPath
}

# ----------------------------
# Batch drain loop
# ----------------------------

for ($batch = 1; $batch -le $MaxBatches; $batch++) {

    # Get next batch of messages from the source folder
    # Get-MgUserMailFolderMessage supports -Top and returns messages in the folder. [2](https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.mail/get-mgusermailfoldermessage?view=graph-powershell-1.0)
    $messages = Invoke-WithRetry {
        Get-MgUserMailFolderMessage `
            -UserId $MailboxUPN `
            -MailFolderId $sourceFolderId `
            -Top $BatchSize `
            -Property @("id","subject","receivedDateTime","internetMessageId")
    }

    if (-not $messages -or $messages.Count -eq 0) {
        Write-Host "No more messages found in '$SourceFolderName'. Done."
        break
    }

    Write-Host "Batch $($batch): moving $($messages.Count) message(s)..."

    foreach ($m in $messages) {
        $msgId = $m.Id
        $subj  = $m.Subject
        $rcv   = $m.ReceivedDateTime
        $imid  = $m.InternetMessageId

        try {
            if ($WhatIf) {
                Write-Host "[WhatIf] Would move: $subj"
                Write-LogRow -Status "WhatIf" -MessageId $msgId -InternetMessageId $imid -Subject $subj -ReceivedDateTime $rcv -Error ""
            }
            else {
                # Move-MgUserMessage moves message to destination folder id. [3](https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.mail/move-mgusermessage?view=graph-powershell-1.0)
                Invoke-WithRetry {
                    $params = @{ destinationId = $targetFolderId }
                    Move-MgUserMessage -UserId $MailboxUPN -MessageId $msgId -BodyParameter $params | Out-Null
                }

                Write-LogRow -Status "Moved" -MessageId $msgId -InternetMessageId $imid -Subject $subj -ReceivedDateTime $rcv -Error ""
            }
        }
        catch {
            $err = $_.Exception.Message
            Write-Host "FAILED move: $subj | $err" -ForegroundColor Yellow
            Write-LogRow -Status "Failed" -MessageId $msgId -InternetMessageId $imid -Subject $subj -ReceivedDateTime $rcv -Error $err
        }

        Start-Sleep -Milliseconds $DelayMsBetweenMoves
    }

    Start-Sleep -Milliseconds $DelayMsBetweenBatches
}

Disconnect-MgGraph | Out-Null
Write-Host "Done. Log: $LogPath"