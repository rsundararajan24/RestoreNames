<#
.SYNOPSIS
  Restores original file names in a folder by stripping a random
  OneDrive/SharePoint-style suffix token that a bulk export process added.

  Author: Raghuraman Sundararajan

  Example:
    2nd_year_report_01UKI5AIVINYUC4U33JBCJWDCKCODYEDYZ.pdf
    -> 2nd_year_report.pdf

.WORKFLOW
  1) Build the manifest (dry run, touches nothing):
       .\RestoreNames.ps1 -Mode Manifest -FolderPath "C:\Path\To\Your\Folder"
     Optionally control the batch-size schedule (progressive sizing), e.g.
     4 batches of 25, then 50, then 100s for the rest:
       .\RestoreNames.ps1 -Mode Manifest -FolderPath "C:\Path\To\Your\Folder" -BatchSizes 25,25,25,25,50,100

  2) Open FolderPath\_rename_manifest.csv in Excel. Filter Status = "OK" and
     Batch = 1. Spot-check a handful of rows against the real files.
     Also check the Note column for any "duplicate group" entries - these
     are collisions that got numbered suffixes (_1, _2, ...), rename them
     to something more meaningful by hand in the CSV if you know which is
     which, before that batch is Applied.

  3) Preview batch 1 (still touches nothing):
       .\RestoreNames.ps1 -Mode Preview -FolderPath "C:\Path\To\Your\Folder" -Batch 1

  4) Apply batch 1 for real:
       .\RestoreNames.ps1 -Mode Apply -FolderPath "C:\Path\To\Your\Folder" -Batch 1

  5) Verify a few renamed files in Explorer, then repeat steps 3-4 for
     Batch 2, 3, ... through the last batch shown in the manifest summary.

.SAFETY NOTES
  - This script only ever RENAMES files - it never opens, edits, or deletes
    file content. Nothing is destroyed even by a wrong rename.
  - Before step 1, back up the whole folder, e.g.:
      robocopy "C:\Path\To\Your\Folder" "C:\Path\To\Backup_2026-01-01" /E
  - Files whose name doesn't match the expected suffix pattern are marked
    NOMATCH and are never renamed.
  - Files that would collide with another file's proposed name (two
    different originals ending up with the same target name) are NOT
    skipped - every member of the collision group gets a numbered suffix
    (_1, _2, ...), ordered oldest-modified first, and is flagged in the
    Note column so you can rename it properly by hand if you want.
  - Every Apply run appends to _rename_log.csv (timestamp, old name, new
    name, result) so you always have a full audit trail.
  - Check a couple of your real filenames before running Manifest: the
    script assumes the suffix is an underscore followed by 20-40 uppercase
    letters/digits immediately before the extension. Adjust $SuffixPattern
    below if yours differ.
#>

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("Manifest","Preview","Apply")]
    [string]$Mode,

    [Parameter(Mandatory=$true)]
    [string]$FolderPath,

    # Batch-size schedule used only in -Mode Manifest. Batch 1 gets the first
    # number of files, batch 2 the next, etc. Once the list is exhausted, the
    # LAST number repeats for all remaining batches. Default: four batches of
    # 25, then 50, then 100s for everything after.
    [int[]]$BatchSizes = @(25,25,25,25,50,100),

    [int]$Batch = 0
)

$ManifestPath = Join-Path $FolderPath "_rename_manifest.csv"
$LogPath      = Join-Path $FolderPath "_rename_log.csv"

# Underscore + 20-40 uppercase letters/digits, right before the extension.
$SuffixPattern = '^(?<base>.+)_(?<token>[A-Z0-9]{20,40})$'

function Get-BatchNumber {
    param([int]$Index, [int[]]$Sizes)
    $cum = 0
    for ($b = 0; $b -lt $Sizes.Count; $b++) {
        $cum += $Sizes[$b]
        if ($Index -le $cum) { return $b + 1 }
    }
    $lastSize = $Sizes[$Sizes.Count - 1]
    $extra = $Index - $cum
    $extraBatches = [Math]::Ceiling($extra / $lastSize)
    return $Sizes.Count + $extraBatches
}

function Build-Manifest {
    $files = Get-ChildItem -Path $FolderPath -File |
             Where-Object { $_.Name -notlike "_rename_*" }

    if ($files.Count -eq 0) {
        Write-Warning "No files found in $FolderPath"
        return
    }

    $rows = foreach ($f in $files) {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        $ext  = $f.Extension
        if ($stem -match $SuffixPattern) {
            [PSCustomObject]@{
                OriginalName    = $f.Name
                ProposedName    = "$($Matches['base'])$ext"
                Base            = $Matches['base']
                Ext             = $ext
                LastWriteTime   = $f.LastWriteTime
                PatternStatus   = "OK"
                Note            = ""
                CollisionGroup  = ""
                Batch           = 0
            }
        } else {
            [PSCustomObject]@{
                OriginalName    = $f.Name
                ProposedName    = $f.Name
                Base            = ""
                Ext             = $ext
                LastWriteTime   = $f.LastWriteTime
                PatternStatus   = "NOMATCH"
                Note            = ""
                CollisionGroup  = ""
                Batch           = 0
            }
        }
    }

    # Resolve collisions: any ProposedName shared by 2+ OK rows gets numbered
    # suffixes for ALL members, oldest-modified first. Record the shared
    # pre-collision name in CollisionGroup so members can be found and kept
    # together later, regardless of batch.
    $groups = $rows | Where-Object PatternStatus -eq "OK" | Group-Object ProposedName
    foreach ($g in $groups) {
        if ($g.Count -gt 1) {
            $ordered = $g.Group | Sort-Object LastWriteTime
            $n = 0
            foreach ($item in $ordered) {
                $n++
                $item.CollisionGroup = $g.Name
                $item.ProposedName = "$($item.Base)_$n$($item.Ext)"
                $item.Note = "duplicate group of $($g.Count) (was: $($g.Name))"
            }
        }
    }

    # First pass: assign batches sequentially per the size schedule.
    $i = 0
    $ordered = $rows | Sort-Object OriginalName
    foreach ($r in $ordered) {
        if ($r.PatternStatus -eq "OK") {
            $i++
            $r.Batch = Get-BatchNumber -Index $i -Sizes $BatchSizes
        }
    }

    # Second pass: pull every member of a collision group into the earliest
    # batch any of them landed in, so duplicates are always reviewed together
    # instead of being split across separate sign-offs. This can make that
    # batch a little larger than the schedule called for - safety over a
    # perfectly even count.
    $collisionGroups = $ordered | Where-Object { $_.CollisionGroup -ne "" } |
                        Group-Object CollisionGroup
    foreach ($cg in $collisionGroups) {
        $minBatch = ($cg.Group | Measure-Object -Property Batch -Minimum).Minimum
        foreach ($item in $cg.Group) { $item.Batch = $minBatch }
    }

    $manifest = foreach ($r in $ordered) {
        [PSCustomObject]@{
            OriginalName   = $r.OriginalName
            ProposedName   = $r.ProposedName
            Status         = $r.PatternStatus
            Batch          = $r.Batch
            CollisionGroup = $r.CollisionGroup
            Note           = $r.Note
        }
    }

    $manifest | Export-Csv -Path $ManifestPath -NoTypeInformation -Encoding UTF8

    $okCount  = ($manifest | Where-Object Status -eq "OK").Count
    $noMatch  = ($manifest | Where-Object Status -eq "NOMATCH").Count
    $dupCount = ($manifest | Where-Object { $_.Note -like "duplicate group*" }).Count
    $groupCnt = ($manifest | Where-Object { $_.CollisionGroup -ne "" } |
                 Select-Object -ExpandProperty CollisionGroup -Unique).Count
    $batchCnt = ($manifest | Where-Object Status -eq "OK" |
                 Select-Object -ExpandProperty Batch -Unique).Count

    Write-Host "Manifest written to $ManifestPath"
    Write-Host "  OK (will rename):     $okCount  across $batchCnt batch(es), schedule: $($BatchSizes -join ',')"
    Write-Host "  NOMATCH (left alone): $noMatch"
    Write-Host "  Duplicate groups:     $groupCnt ($dupCount files total, kept together within one batch each)"
}

function Process-Batch {
    param([switch]$ApplyChanges)

    if (-not (Test-Path $ManifestPath)) {
        Write-Error "No manifest found at $ManifestPath. Run -Mode Manifest first."
        return
    }
    $manifest = Import-Csv $ManifestPath
    $rows = $manifest | Where-Object { $_.Status -eq "OK" -and [int]$_.Batch -eq $Batch }

    if ($rows.Count -eq 0) {
        Write-Warning "No OK rows found for batch $Batch. Check the manifest's Batch column."
        return
    }

    $results = foreach ($r in $rows) {
        $src = Join-Path $FolderPath $r.OriginalName
        $dst = Join-Path $FolderPath $r.ProposedName
        $result = [PSCustomObject]@{
            Timestamp    = Get-Date -Format s
            Batch        = $Batch
            OriginalName = $r.OriginalName
            ProposedName = $r.ProposedName
            Action       = if ($ApplyChanges) { "RENAME" } else { "WHATIF" }
            Result       = ""
        }

        if (-not (Test-Path $src)) {
            $result.Result = "SKIP - source missing"
        } elseif (Test-Path $dst) {
            $result.Result = "SKIP - target already exists"
        } elseif ($ApplyChanges) {
            try {
                Rename-Item -Path $src -NewName $r.ProposedName -ErrorAction Stop
                $result.Result = "SUCCESS"
            } catch {
                $result.Result = "ERROR - $($_.Exception.Message)"
            }
        } else {
            $result.Result = "WOULD RENAME"
        }
        $result
    }

    $results | Format-Table Timestamp, OriginalName, ProposedName, Result -AutoSize

    if ($ApplyChanges) {
        # Retry the log write a few times in case the CSV is open elsewhere
        # (e.g. in Excel for review) - the renames above have already
        # happened regardless, so this only protects the audit trail.
        $maxAttempts = 5
        $attempt = 0
        $written = $false
        while (-not $written -and $attempt -lt $maxAttempts) {
            $attempt++
            try {
                $results | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -Append
                $written = $true
            } catch {
                if ($attempt -ge $maxAttempts) {
                    Write-Warning "Could not write to $LogPath after $maxAttempts attempts - it may be open in another program (e.g. Excel). The renames above still succeeded; close the file and re-run this same command to backfill the log, or append manually."
                } else {
                    Start-Sleep -Seconds 2
                }
            }
        }
        if ($written) {
            Write-Host "Batch $Batch applied. Log appended to $LogPath"
        }
    } else {
        # Also write a full-width text file so long filenames are never
        # truncated (PowerShell's console table view clips wide columns).
        $previewPath = Join-Path $FolderPath "_preview_batch$Batch.txt"
        $results | Format-Table Timestamp, OriginalName, ProposedName, Result -Wrap |
            Out-String -Width 300 | Out-File -FilePath $previewPath -Encoding UTF8
        Write-Host "Preview only - nothing was renamed."
        Write-Host "Full untruncated preview written to $previewPath"
        Write-Host "Re-run with -Mode Apply -Batch $Batch when ready."
    }
}

switch ($Mode) {
    "Manifest" { Build-Manifest }
    "Preview"  { Process-Batch }
    "Apply"    { Process-Batch -ApplyChanges }
}
