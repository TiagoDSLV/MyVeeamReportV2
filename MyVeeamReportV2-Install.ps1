# Définir le dossier de destination
$destinationFolder = "C:\Tools\MyVeeamReport"

# Créer le dossier s'il n'existe pas
if (-not (Test-Path -Path $destinationFolder)) {
    New-Item -Path $destinationFolder -ItemType Directory | Out-Null
}

# URLs des fichiers (raw GitHub)
$fileUrls = @(
    "https://raw.githubusercontent.com/TiagoDSLV/MyVeeamReportV2/refs/heads/main/MyVeeamReportV2-Config.ps1",
    "https://raw.githubusercontent.com/TiagoDSLV/MyVeeamReportV2/refs/heads/main/MyVeeamReportV2-Script.ps1"
)

# Télécharger les fichiers
foreach ($url in $fileUrls) {
    $fileName = Split-Path -Path $url -Leaf
    $destinationPath = Join-Path -Path $destinationFolder -ChildPath $fileName

    try {
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction Stop
        $response.Content | Out-File -FilePath $destinationPath -Encoding utf8
    } catch {
        Write-Error "Erreur - Échec du téléchargement des fichiers merci de les copier manuellement depuis le repository Github suivant : https://github.com/TiagoDSLV/MyVeeamReportV2"
    }
}

Write-Host "OK - Fichiers téléchargés dans : $destinationFolder"

# --------------------------------------
# Création de la tâche planifiée SYSTEM
# --------------------------------------

$taskName = "ATHEO - MyVeeamReportV2 - Daily"
$scriptFile = Join-Path $destinationFolder "MyVeeamReportV2-Script.ps1"
$configFile = "MyVeeamReportV2-Config.ps1"

# Définir l'action
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptFile`" -ConfigFileName `"$configFile`""

# Déclencheur : tous les jours à 8h30
$trigger = New-ScheduledTaskTrigger -Daily -At 8:30AM

# Compte SYSTEM avec élévation
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

# Paramètres de la tâche
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

# Tâche complète
$task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings

# Enregistrement
Register-ScheduledTask -TaskName $taskName -InputObject $task -Force | Out-Null

Write-Host "OK - Tâche planifiée '$taskName' créée pour 8h30 tous les jours (compte SYSTEM)"

# --------------------------------------
# Ouvrir automatiquement le fichier de configuration
# --------------------------------------

Write-Host "OK - Veuillez maintenant modifier le fichier de configuration qui s'ouvre dans powershell"
$configPath = Join-Path $destinationFolder $configFile
Start-Process powershell_ise.exe $configPath
