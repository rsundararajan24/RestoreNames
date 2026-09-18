# RestoreNames — Batch-Safe Filename Restoration

A PowerShell tool for stripping random suffix tokens that cloud storage
platforms (OneDrive, SharePoint, etc.) sometimes append to files during
bulk export or conflict resolution — restoring the original filenames in
small, verifiable batches instead of one irreversible pass.

## Problem Statement

When files are exported in bulk from certain cloud storage / collaboration
platforms (commonly after a conflict, a dedup pass, or a bulk-download
export), each filename can get a random alphanumeric token appended before
the extension, e.g.:

```
2nd_year_report.pdf
  → 2nd_year_report_01UKI5AIVINYUC4U33JBCJWDCKCODYEDYZ.pdf
```

Across a large folder (hundreds of files), restoring the original names by
hand is slow and error-prone. A naive scripted approach ("strip everything
after the last underscore") is risky at scale because:

1. **Pattern mismatches** — not every filename necessarily fits the
   expected suffix shape, and a bad match could mangle a legitimate name.
2. **Collisions** — if two files originally shared the same name (e.g. two
   uploads of the same document, or genuinely different documents that
   happen to share a name), stripping the suffix on both produces a name
   clash, and one file can silently overwrite the other.
3. **No undo** — a bulk rename with a bug has no built-in way to know what
   happened or roll it back.
4. **Scale + stakes** — when the files are personal records (financial,
   medical, identity documents), a scripted "just run it" approach isn't
   acceptable; verification has to be built into the workflow, not bolted
   on after the fact.

## Approach

The script splits the job into two clearly separated phases, run
independently and never mixed:

**1. Planning phase (`-Mode Manifest`) — touches nothing.**
Scans every file, computes what its original name would be by stripping
the suffix, and writes a full CSV manifest before any file is renamed.
During this pass it also:
- Flags any filename that doesn't match the expected suffix pattern
  (`NOMATCH`) so it's never touched.
- Detects **collisions** — files that would end up sharing a name — and
  disambiguates every member of the group with a numbered suffix
  (`_1`, `_2`, ...), ordered oldest-modified first, so nothing silently
  overwrites anything else.
- Groups colliding files into the **same batch**, even if it makes that
  batch a little larger than the target size, so duplicates are always
  reviewed together rather than scattered across separate sign-offs.
- Assigns files to numbered batches using a configurable, **progressive
  size schedule** (e.g. start with small batches of 25 to build
  confidence, then increase to 50, then 100 once the pattern is proven
  correct).

**2. Execution phase (`-Mode Preview` / `-Mode Apply`) — one batch at a
time.**
- `Preview` shows exactly what a batch would do, without renaming
  anything — for manual review before committing.
- `Apply` renames only that batch's files, and appends a timestamped
  record (old name, new name, result) to a persistent log — giving a full
  audit trail and a way to reconstruct what happened if anything needs to
  be reversed.

This turns a single risky bulk operation into a sequence of small,
independently verifiable, resumable steps — each one safe to stop after,
resume later, or review with a domain expert (in this case, the file
owner) before proceeding.

## Usage

```powershell
# 1. Back up the folder first (outside the folder itself)
robocopy "C:\Path\To\Your\Folder" "C:\Path\To\Backup_2026-01-01" /E

# 2. Build the manifest — dry run, plans everything, renames nothing
.\RestoreNames.ps1 -Mode Manifest -FolderPath "C:\Path\To\Your\Folder" -BatchSizes 25,25,25,25,50,100

# 3. Review the manifest CSV (_rename_manifest.csv) — check a batch's
#    OriginalName -> ProposedName pairs, and check the CollisionGroup /
#    Note columns for any duplicate sets

# 4. Preview a batch — still renames nothing
.\RestoreNames.ps1 -Mode Preview -FolderPath "C:\Path\To\Your\Folder" -Batch 1

# 5. Apply the batch — renames for real, logs the result
.\RestoreNames.ps1 -Mode Apply -FolderPath "C:\Path\To\Your\Folder" -Batch 1

# 6. Repeat steps 3-5 for each remaining batch
```

If script execution is blocked by Windows' default policy:

```powershell
Unblock-File .\RestoreNames.ps1
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
```
(`-Scope Process` only affects the current PowerShell window — your
system's default policy is untouched.)

## Safety Properties

- **Nothing is renamed during planning.** The manifest step is entirely
  read-only.
- **Content is never touched.** Only filenames change — file bytes are
  never opened, read, or modified, so a bad rename is always reversible
  from the log.
- **Collisions can't silently overwrite.** Every colliding target name is
  disambiguated before any renaming starts, and the rename step itself
  also checks `Test-Path` on the destination as a second guard.
- **Full audit trail.** `_rename_log.csv` records every rename attempt
  with a timestamp and result (`SUCCESS` / `SKIP` / `ERROR`), across every
  batch, for the life of the run.
- **Resumable.** Batches are processed and applied independently — you
  can stop after any batch and resume later without re-planning.

## Known Limitations / Possible Improvements

Resolved during the first real run of this script:

- **Log writes can fail if the CSV is open elsewhere** (e.g. in Excel for
  review). The Apply path now retries the log write with a short backoff
  (up to 5 attempts) before giving up — the renames themselves aren't
  affected either way, since they happen before the log write is
  attempted.
- **Console output truncation.** Long filenames were getting clipped in
  PowerShell's default table view. Preview mode now also writes a full,
  untruncated text file (`_preview_batch<N>.txt`) alongside the console
  output, so nothing needs manual `Out-File -Width` piping to review.

Still open:

- **Collision groups still need human judgment.** The script can tell you
  two files would produce the same name, but not which one is the "real"
  document vs. an amendment/duplicate — that final call is left to the
  file owner (see the `Note` and `CollisionGroup` columns in the
  manifest).
- **True duplicate detection.** A content hash (`Get-FileHash`) comparison
  within each collision group could distinguish byte-identical duplicates
  (safe to auto-resolve) from same-named-but-different documents (still
  needs manual review) — not currently implemented.

## License

MIT — see [LICENSE](LICENSE).

