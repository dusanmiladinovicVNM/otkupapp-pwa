# check_merge.ps1 - Read-only provera: da li je main ispred grane i ima li merge konflikata
# Upotreba: python-free, cist git; poziva se iz root-a repoa:
#   powershell -File tools\check_merge.ps1
#   powershell -File tools\check_merge.ps1 -BaseBranch origin/main

param(
    [string]$BaseBranch = "origin/main"
)

$ErrorActionPreference = "Stop"

Write-Output "=== $BaseBranch ispred grane? ==="
git fetch origin --quiet
$ahead = git log --oneline "HEAD..$BaseBranch"
if ($ahead) {
    Write-Output $ahead
} else {
    Write-Output "(prazno = grana sadrzi ceo $BaseBranch)"
}

Write-Output ""
Write-Output "=== konflikti ==="
$mb = git merge-base HEAD $BaseBranch
$conflicts = git merge-tree $mb $BaseBranch HEAD |
    Select-String -Pattern '^\+<<<<<<<|changed in both' -SimpleMatch:$false

if ($conflicts) {
    Write-Output $conflicts
    Write-Output ""
    Write-Output "REZULTAT: KONFLIKT - merge/rebase ce traziti rucnu intervenciju"
    exit 1
} else {
    Write-Output "(prazno = nema konflikta)"
    Write-Output ""
    Write-Output "REZULTAT: CISTO"
    exit 0
}
