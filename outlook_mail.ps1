<#
    Shared by fetch_fin.ps1, fetch_pj.ps1 and fetch_bbg.ps1 (dot-sourced).

    Recent mail from the Inbox AND the Archive folder. A mailbox rule files some of these
    newsletters (WSJ Markets A.M. on 10-29) straight into Archive, so an Inbox-only scan
    reads "no email found" for a message that is plainly in the mailbox.
#>
function Get-RecentMail($outlook, [datetime]$since) {
    $ns = $outlook.GetNamespace("MAPI")
    $folders = @($ns.GetDefaultFolder(6))
    try {
        $root = $ns.GetDefaultFolder(6).Parent
        foreach ($f in $root.Folders) { if ($f.Name -eq "Archive") { $folders += $f } }
    } catch {}
    $filter = "[ReceivedTime] >= '" + $since.ToString("MM/dd/yyyy hh:mm tt") + "'"
    $out = @()
    foreach ($f in $folders) {
        try {
            $items = $f.Items
            $items.Sort("[ReceivedTime]", $true)
            foreach ($i in $items.Restrict($filter)) { $out += $i }
        } catch {}
    }
    return @($out | Sort-Object { try { $_.ReceivedTime } catch { [datetime]::MinValue } } -Descending)
}
