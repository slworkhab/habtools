Clear-Host
do {
    Write-Host "`n=== MENU ==="
    Write-Host "1. Nettoyer (clean.py)"
    Write-Host "2. Copier les fichiers XLSX (copywin.py)"
    Write-Host "3. Extraire les requêtes (extractwin.py)"
    Write-Host "4. Quitter"
    $choice = Read-Host "Choisissez une option (1-4)"

    switch ($choice) {
        1 { python "clean.py" }
        2 { python "copywin.py" }
        3 { python "extractwin.py" }
        4 { Write-Host "Au revoir !" }
        default { Write-Host "Option invalide." }
    }
} while ($choice -ne 4)
