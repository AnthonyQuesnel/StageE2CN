<#
Titre : Mise à jour d'un ordinateur
Date de création : 31/05/2022
Auteur : Anthony Quesnel
Version : 1.1
Date de modification : 02/06/22
Modifications effectuées :
    - Ecriture d'un fichier log
Résumé du fonctionnement du script :
    - L'ordinateur recherche les mises à jour non-installées
    - Si des mises à jour sont trouvées:
        - L'ordinateur télécharge les mises à jour
        - L'ordinateur installe les mises à jour
    - Si aucune mise à jour n'est trouvée le script s'arrête
#>

### Déclaration des variables ###

# Chemin d'accès au fichier log
$CheminDossierLogs = "C:\Logs\"
$CheminFichierLog = $CheminDossierLogs + "MaJ-" + (Get-Date -Format "ddMMyyyy") + ".txt"

# Définition d'un critère de recherche pour les mises à jour non installées
$CritereMaJ = "IsInstalled=0"

# Création des objets permettant le téléchargement et l'installation de mises à jour
$DownloadMaJ = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateDownloader()
$InstallMaJ = New-Object -ComObject Microsoft.Update.Installer

### Lancement des commandes ###

# Création d'un dossier "Logs" où stocker les logs
if (-not (Test-Path "$CheminDossierLogs")) {
    New-Item "$CheminDossierLogs" -itemType Directory
}

# Recherche de mises à jour correspondant au critère de recherche
try {
	$RechercheMaJ = (New-Object -ComObject Microsoft.Update.Searcher).Search($CritereMaJ).Updates
}
catch {
    Add-Content -Path $CheminFichierLog -Value "Une erreur est survenue dans la recherche de mises à jour"
    exit
}

# Fin du script si aucune mise à jour trouvée
if ($RechercheMaJ.Count -eq 0) {
    Add-Content -Path $CheminFichierLog -Value "L'ordinateur est déjà à jour"
    exit
}

# Téléchargement des mises à jour
$DownloadMaJ.Updates = $RechercheMaJ

try {
    $DownloadMaJ.Download()
}
catch {
    Add-Content -Path $CheminFichierLog -Value "Une erreur est survenue dans le téléchargement des mises à jour"
    exit
}

# Installation des mises à jour
$InstallMaJ.Updates = $RechercheMaJ

try {
    $InstallMaJ.Install()
}
catch {
    Add-Content -Path $CheminFichierLog -Value "Une erreur est survenue dans l'installation des mises à jour"
    exit
}

# Ecriture dans le fichier log des mises à jour installées
Add-Content -Path $CheminFichierLog -Value ("Nombre de mises à jour installées: " + $RechercheMaJ.Count)

foreach ($MiseAJour in $RechercheMaJ) {
    Write-Host ("- " + $MiseAJour.Title)
}