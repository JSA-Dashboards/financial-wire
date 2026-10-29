<#
    fetch_pj.ps1 - Every PJ Quaid email in the window, reduced to its prose paragraphs.
    Feeds the Headline Candidates page (headlines.py); nothing goes to the wire from here.

      powershell -ExecutionPolicy Bypass -File .\fetch_pj.ps1 [-Hours 24]
#>
param(
    [int]$Hours = 0,
    [string]$Match = "Quaid",
    [string]$OutFile = "$PSScriptRoot\pj_candidates.json"
)
$ErrorActionPreference = "Stop"
if ($Hours -le 0) { $Hours = if ((Get-Date).DayOfWeek -eq "Monday") { 72 } else { 30 } }

$tail = @('Trading commodity futures, options on futures', 'This email is not intended for residents',
          'No part of this material may be copied', 'Full Disclaimer')
function Clean([string]$t) {
    $t = $t -replace '<https?://[^>]*>', '' -replace '<mailto:[^>]*>', ''
    $t = $t -replace '[^\x20-\x7E\r\n‘’“”–—…]', ' '
    foreach ($mk in $tail) { $m = [regex]::Match($t, [regex]::Escape($mk)); if ($m.Success) { $t = $t.Substring(0, $m.Index) } }
    return $t
}

$o = New-Object -ComObject Outlook.Application
. "$PSScriptRoot\outlook_mail.ps1"
$cut = (Get-Date).AddHours(-$Hours)
$rows = @()
foreach ($m in (Get-RecentMail $o $cut)) {
    try {
        if ("$($m.SenderName) $($m.SenderEmailAddress)" -notlike "*$Match*") { continue }
        foreach ($ln in ((Clean $m.Body) -split "`n")) {
            $p = ($ln -replace '\s+', ' ').Trim()
            # long prose only; skips option-flow, vol and signature lines
            if ($p.Length -le 150) { continue }
            if ($p -match 'StoneX Group|all rights reserved|Senior Vice President|Options O-\d') { continue }
            # trade tickets ("Corn S 250 hz -5 put 22") are mostly numbers; prose is mostly words
            $w = @($p -split ' ')
            if (@($w | Where-Object { $_ -match "^[A-Za-z'.,;:()\-]+$" }).Count / $w.Count -lt 0.7) { continue }
            $k = $p.Substring(0, 60)
            # the same paragraph often repeats as preview text or across sends
            if ($rows | Where-Object { $_.text.StartsWith($k) -or $p.StartsWith($_.text.Substring(0, 60)) }) { continue }
            $rows += [pscustomobject]@{ subject = [string]$m.Subject; received = $m.ReceivedTime.ToString("ddd HH:mm"); text = $p }
        }
    } catch {}
}
($rows | ConvertTo-Json -Depth 3) | Out-File $OutFile -Encoding utf8
"PJ Quaid paragraphs: $($rows.Count) from $(@($rows | Select-Object -ExpandProperty subject -Unique).Count) email(s) -> $OutFile"
