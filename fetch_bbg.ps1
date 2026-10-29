<#
    fetch_bbg.ps1 - Bloomberg "Markets Daily": the Five Things and On the move bullets.
    Feeds the Headline Candidates page (headlines.py); nothing goes to the wire from here.

      powershell -ExecutionPolicy Bypass -File .\fetch_bbg.ps1 [-Hours 30]
#>
param(
    [int]$Hours = 0,
    [string]$OutFile = "$PSScriptRoot\bbg_candidates.json"
)
$ErrorActionPreference = "Stop"
if ($Hours -le 0) { $Hours = if ((Get-Date).DayOfWeek -eq "Monday") { 72 } else { 30 } }

$o = New-Object -ComObject Outlook.Application
. "$PSScriptRoot\outlook_mail.ps1"
$cut = (Get-Date).AddHours(-$Hours)
$msg = $null
foreach ($m in (Get-RecentMail $o $cut)) {
    try {
        if ("$($m.SenderEmailAddress)" -like "*news.bloomberg.com*" -and $m.Subject -match 'Markets Daily') { $msg = $m; break }
    } catch {}
}

$rows = @()
if ($msg) {
    $t = $msg.Body -replace '<https?://[^>]*>', '' -replace '<mailto:[^>]*>', ''
    $t = $t -replace "[$([char]0x200B)-$([char]0x200F)$([char]0x034F)$([char]0xFEFF)$([char]0x00AD)]", ''
    $t = $t -replace '[ \t]+', ' '
    $when = $msg.ReceivedTime.ToString("ddd HH:mm")
    foreach ($sec in @(@{ name = "Five things"; rx = '(?s)Five things you need to know\s*\n(.*?)\n(?=[^*\n])' },
                       @{ name = "On the move"; rx = '(?s)\nOn the move\s*\n(.*?)\n(?=[^*\n])' })) {
        $m2 = [regex]::Match($t, $sec.rx)
        if (-not $m2.Success) { continue }
        foreach ($ln in ($m2.Groups[1].Value -split "`n")) {
            $b = ($ln -replace '^[\s\*•]+', '' -replace '\s+', ' ').Trim()
            if ($b.Length -gt 25) { $rows += [pscustomobject]@{ section = $sec.name; received = $when; text = $b } }
        }
    }
}
($rows | ConvertTo-Json -Depth 3) | Out-File $OutFile -Encoding utf8
"Bloomberg bullets: $($rows.Count) -> $OutFile"
