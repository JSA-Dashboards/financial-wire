<#
    fetch_fin.ps1 - Pull the Morning Financial Wire pieces out of Outlook.

    Reads sources.json, takes the newest matching message per source, and parses
    the three fixed-layout bits straight out of the body:
      snapshot  - Bloomberg "Market Snapshot" table
      reading   - WSJ "What I'm Reading" bullets
      history   - WSJ "Today in Markets History"
      macro     - PJ Quaid's "first look" paragraphs (his text, verbatim)
    Writes fin_input.json for build_fin_wire.py.

    Usage:
      powershell -ExecutionPolicy Bypass -File .\fetch_fin.ps1
      powershell -ExecutionPolicy Bypass -File .\fetch_fin.ps1 -MsgDir C:\Users\me\Downloads   # test on saved .msg files
#>
param(
    [int]$Hours = 0,
    [string]$MsgDir = "",
    [string]$MsgNameRx = "first look|Markets Daily|Markets A\.M\.",   # test mode: only open these saved files
    [string]$OutFile = "$PSScriptRoot\fin_input.json"
)

$ErrorActionPreference = "Stop"
$cfg = Get-Content "$PSScriptRoot\sources.json" -Raw | ConvertFrom-Json
$outlook = New-Object -ComObject Outlook.Application

function Clean-Body([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    $text = $text -replace '<https?://[^>]*>', ''
    $text = $text -replace '<mailto:[^>]*>', ''
    $text = $text -replace '(?m)^\s*(External|Internal) \([^)]*\)\s*$', ''
    $text = $text -replace '(?m)^\s*(Safe|Graymail|Spam|Phish|More\.\.\.)(\s+(Safe|Graymail|Spam|Phish|More\.\.\.))*\s*$', ''
    # zero-width / non-breaking padding Bloomberg and WSJ use for preview text
    $text = $text -replace "[$([char]0x200B)-$([char]0x200F)$([char]0x034F)$([char]0xFEFF)$([char]0x00AD)]", ''
    $text = $text -replace '[ \t]+', ' '
    $text = $text -replace '(\r?\n\s*){2,}', "`n"
    return $text.Trim()
}

function Clean-Quaid([string]$text) {
    $tail = @('Trading commodity futures, options on futures', 'This email is not intended for residents',
              'No part of this material may be copied', 'Full Disclaimer')
    foreach ($mk in $tail) {
        $m = [regex]::Match($text, [regex]::Escape($mk))
        if ($m.Success) { $text = $text.Substring(0, $m.Index) }
    }
    # signature block: everything from the sender's name line on
    $m = [regex]::Match($text, '(?m)^\s*(PJ Quaid|P\.J\. Quaid)\s*$')
    if ($m.Success) { $text = $text.Substring(0, $m.Index) }
    return $text.Trim()
}

function Parse-Snapshot([string]$body) {
    $rows = @()
    $m = [regex]::Match($body, '(?s)Market Snapshot\s*\n(.*?)\nMarket data as of ([^\n.]*)')
    if (-not $m.Success) { return @{ rows = @(); asOf = "" } }
    foreach ($ln in ($m.Groups[1].Value -split "`n")) {
        $r = [regex]::Match($ln.Trim(), '^(.*?)\s+(\$?[\d,]+(?:\.\d+)?%?)\s+([+\-\u2212][\d,]+(?:\.\d+)?%?)$')
        if ($r.Success) { $rows += [pscustomobject]@{ name = $r.Groups[1].Value; level = $r.Groups[2].Value; change = ($r.Groups[3].Value -replace [char]0x2212, '-') } }
    }
    return @{ rows = $rows; asOf = $m.Groups[2].Value.Trim() }
}

function Parse-Reading([string]$body) {
    $m = [regex]::Match($body, "(?s)What I.{1,2}m Reading\s*\n(.*?)\nToday in Markets History")
    if (-not $m.Success) { return @() }
    $items = @()
    foreach ($ln in ($m.Groups[1].Value -split "`n")) {
        $t = ($ln -replace '^[\s\u2022\*\-]+', '').Trim()
        if ($t) { $items += $t -replace '\s+\)', ')' }
    }
    return $items
}

function Parse-History([string]$body) {
    $m = [regex]::Match($body, '(?s)Today in Markets History\s*\n(.*?)\n(Beyond the Newsroom|About Me|\Z)')
    if (-not $m.Success) { return "" }
    return ($m.Groups[1].Value -replace '^[^\w"\u201C]+', '').Trim()
}

# ---- gather candidate messages ----
$msgs = @()
if ($MsgDir) {
    # Open a temp copy so a lingering Outlook handle on the original can't block the read.
    $tmpDir = Join-Path $env:TEMP ("fin_msgs_" + [guid]::NewGuid().ToString("N").Substring(0, 8))
    New-Item -ItemType Directory $tmpDir -Force | Out-Null
    $n = 0
    foreach ($f in Get-ChildItem $MsgDir -Filter *.msg | Where-Object { $_.Name -match $MsgNameRx }) {
        try { $n++; $c = Join-Path $tmpDir "$n.msg"; Copy-Item $f.FullName $c; $msgs += $outlook.Session.OpenSharedItem($c) } catch {}
    }
} else {
    if ($Hours -le 0) { $Hours = if ((Get-Date).DayOfWeek -eq "Monday") { 72 } else { 24 } }
    $cutoff = (Get-Date).AddHours(-$Hours)
    $all = $outlook.GetNamespace("MAPI").GetDefaultFolder(6).Items
    $all.Sort("[ReceivedTime]", $true)
    $items = $all.Restrict("[ReceivedTime] >= '" + $cutoff.ToString("MM/dd/yyyy hh:mm tt") + "'")
    foreach ($i in $items) { $msgs += $i }
}

$found = @{}
foreach ($msg in $msgs) {
    try {
        $from = "$($msg.SenderName) $($msg.SenderEmailAddress)"
        $subj = [string]$msg.Subject
        foreach ($s in $cfg.sources) {
            if ($from -notlike "*$($s.match)*") { continue }
            $ok = $false
            foreach ($p in $s.subjectAny) { if ($subj -match $p) { $ok = $true; break } }
            if (-not $ok) { continue }
            $prev = $found[$s.key]
            if ($null -eq $prev -or $msg.ReceivedTime -gt $prev.ReceivedTime) { $found[$s.key] = $msg }
        }
    } catch {}
}

# ---- parse ----
$out = [ordered]@{
    generatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm")
    recipient   = $cfg.recipient
    subject     = "$($cfg.subjectPrefix) - $((Get-Date).ToString('ddd MMM d'))"
    snapshot    = @(); snapshotAsOf = ""; reading = @(); history = ""; macro = @()
    sources     = @(); gaps = @()
}

foreach ($s in $cfg.sources) {
    $msg = $found[$s.key]
    if ($null -eq $msg) { $out.gaps += "$($s.label): no email found in the window."; continue }
    $received = $msg.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
    $out.sources += "$($s.label) ($received)"
    $body = Clean-Body $msg.Body

    switch ($s.key) {
        "bloomberg" {
            $snap = Parse-Snapshot $body
            $out.snapshot = @($snap.rows); $out.snapshotAsOf = $snap.asOf
            if ($snap.rows.Count -eq 0) { $out.gaps += "Bloomberg: Market Snapshot table not found - layout may have changed." }
        }
        "wsj" {
            $out.reading = @(Parse-Reading $body)
            $out.history = Parse-History $body
            if ($out.reading.Count -eq 0) { $out.gaps += "WSJ: 'What I'm Reading' not found." }
            if (-not $out.history)        { $out.gaps += "WSJ: 'Today in Markets History' not found." }
        }
        "pj" {
            $q = Clean-Quaid $body
            # Macro commentary is the long prose paragraphs. The same email also carries
            # option flow, open-interest and vol lines - short, and not part of this wire.
            $paras = @($q -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -gt 150 })
            # The first line is often the preview text, repeated in full right after.
            if ($paras.Count -gt 1 -and $paras[1].StartsWith($paras[0].TrimEnd('.', ' ', [char]0x2026))) { $paras = $paras[1..($paras.Count - 1)] }
            $out.macro = $paras
            if ($out.macro.Count -eq 0) { $out.gaps += "PJ Quaid: no paragraphs extracted." }
        }
    }
}

($out | ConvertTo-Json -Depth 6) | Out-File $OutFile -Encoding utf8

"Snapshot rows : $($out.snapshot.Count)  ($($out.snapshotAsOf))"
"Reading items : $($out.reading.Count)"
"History       : $(if ($out.history) { 'yes' } else { 'MISSING' })"
"PJ paragraphs : $($out.macro.Count)"
foreach ($g in $out.gaps) { "GAP: $g" }
"Written to: $OutFile"
