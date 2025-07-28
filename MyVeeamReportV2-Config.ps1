<#====================================================================
Author        : Tiago DA SILVA - ATHEO INGENIERIE
Version       : 1.0.3
Creation Date : 2025-07-01
Last Update   : 2025-07-01
GitHub Repo   : https://github.com/TiagoDSLV/MyVeeamReportV2/
====================================================================

DESCRIPTION:
My Veeam Report is a flexible reporting script for Veeam Backup and
Replication. This report can be customized to report on Backup, Replication,
Backup Copy, Tape Backup, SureBackup and Agent Backup jobs as well as
infrastructure details like repositories, proxies and license status. 

====================================================================#>

# Configuration du Rapport
    $SDBackup = $true
    $SDCopy = $true
    $SDReplication = $true
    $SDTape = $true
    $SDAgent = $true
    $SDSure =$true

# Nom du client
    $Client = ""
  # Mail Collecteur GLPI
    $MailGLPI = ""
# Chemin du dossier où stocker les rapports
    $path = "C:\Tools\MyVeeamReport\Reports"
# Rétention des fichiers de rapports JSON en jours
    $JPurge = 60
	
# Report mode (RPO) - valid modes: any number of hours, Weekly or Monthly
# 24, 48, "Weekly", "Monthly"
    $reportMode = 24
	
# Configuration de l'envoi d'email
    $emailHost = ""
    $emailPort = 25
    $emailEnableSSL = $false
    $emailUser = ""
    $emailPass = ""
    $emailFrom = ""
    $emailTo = ""

# Exclusions
    # Exclure les VMs des sections des sauvegardes manquantes et réussies
    $excludevms = @("*_replica")
    # Exclure les VMs des sections des sauvegardes manquantes et réussies dans les dossiers suivants
    $excludeFolder = @("")
    # Exclure les VMs des sections des sauvegardes manquantes et réussies dans les datacenter suivants
    $excludeDC = @("")
    # Exclure les VMs des sections des sauvegardes manquantes et réussies dans les clusters suivants
    $excludeCluster = @("")
    # Exclure les VMs des sections des sauvegardes manquantes et réussies dans les tags suivants
    $excludeTags = @("")
    # Exclure les modèles des sections des sauvegardes manquantes et réussies
    $excludeTemp = $true
    # Exclure les repositories
    $excludedRepositories = @("")

# Seuils d'alerte
    # Espace libre sur les Repository en %
    $repoCritical = 10
    $repoWarn = 20
    # Espace libre sur les cibles de réplication en %
    $replicaCritical = 10
    $replicaWarn = 20
    # Jours restant sur la licence
    $licenseCritical = 15
    $licenseWarn = 30

# Location of Veeam Core dll  
    $VeeamCorePath = "C:\Program Files\Veeam\Backup and Replication\Backup\Veeam.Backup.Core.dll"
