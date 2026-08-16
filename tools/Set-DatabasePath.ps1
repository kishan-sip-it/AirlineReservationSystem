<#
.SYNOPSIS
    Repoints AirReservation.mdb's connection string across the whole project
    if you place the project folder somewhere other than C:\AirlineApp\.

.USAGE
    Open PowerShell in the project's ROOT folder (the one containing AirReservation.vbp)
    and run:

        .\tools\Set-DatabasePath.ps1 -NewPath "D:\Projects\AirlineReservationSystem\AirReservation.mdb"

    Then reopen AirReservation.vbp in VB6 (close it first if it's already open).
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$NewPath
)

$oldConn = 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\AirlineApp\AirReservation.mdb;Persist Security Info=False'
$newConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$NewPath;Persist Security Info=False"

$targets = Get-ChildItem -Path .. -Include *.frm, *.Dsr -Recurse
$changed = 0

foreach ($file in $targets) {
    $content = Get-Content -Raw -Encoding Default $file.FullName
    if ($content -match [regex]::Escape($oldConn)) {
        $content = $content -replace [regex]::Escape($oldConn), $newConn
        Set-Content -Encoding Default -NoNewline $file.FullName $content
        Write-Host "Updated: $($file.Name)"
        $changed++
    }
}

Write-Host "`nDone. $changed file(s) repointed to: $NewPath"
Write-Host "Make sure the .mdb file itself is actually at that path before running the app."
