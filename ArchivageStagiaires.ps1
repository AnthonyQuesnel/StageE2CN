<#
Titre : Archivage de comptes stagiaires
Date de création : 19/05/2022
Auteur : Anthony Quesnel
Version : 1.1
Date de modification : 24/05/2022
Modifications effectuées :
    - Ajout d'un log
    - Les comptes non utilisés ne sont pas archivés et supprimés
Résumé du fonctionnement du script :
    - Détection des comptes stagiaires qui sont inactifs depuis un nombre de jours donnés
    - Désactivation des comptes inactifs et archivage des données utilisateur associées
    - Détection des comptes inactifs qui ne sont plus utilisé depuis un autre nombre de jours donnés
	- Suppression de ces comptes ainsi que de des archives associées
#>

### Chargement de librairies ###
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Chemin de l'OU Stagiaires dans l'Active Directory
$NomOUUsers = "Users"
$NomOUParent = "E2CN"
$NomDomaineAD = "aqtest"
$ExtensionDomaineAD = "loc"
$PathOUUsers = "OU=$NomOUUsers,OU=$NomOUParent,DC=$NomDomaineAD,DC=$ExtensionDomaineAD"

$NomOUStagiaires = "Stagiaires"
$PathOUStagiaires = "OU=$NomOUStagiaires,$PathOUUsers"

$ModeleNomOUPromo = "Promo "
$ModeleCheminDossierStagiaire = "C:\DossiersReseauStagiaires\"

$NomOUArchive = "Archives"
$PathOUArchive = "OU=$NomOUArchive,OU=$NomOUUsers,OU=$NomOUParent,DC=$NomDomaineAD,DC=$ExtensionDomaineAD"
$CheminDossierArchive = "C:\$NomOUArchive\"

$CheminDossierLogs = "C:\Logs\"
$CheminFichierLog = $CheminDossierLogs + "Archivage-" + (Get-Date -Format "ddMMyyyy") + ".txt"

# Limite de jours de dernière connexion
$LimiteArchivage = 90
$LimiteSuppression = 1825

### Lancement des commandes ###

# Création d'une OU "Archives" où stocker les stagiaires à archiver
if (-not (Get-ADOrganizationalUnit -Filter "distinguishedName -eq '$PathOUArchive'")) {
    # Création de l'OU si elle n'existe pas
    New-ADOrganizationalUnit -Name $NomOUArchive -Path $PathOUUsers
}

# Création d'un dossier "Archives" où stocker les documents zippés des stagiaires
if (-not (Test-Path "$CheminDossierArchive")) {
    New-Item "$CheminDossierArchive" -itemType Directory
}

# Création d'un dossier "Logs" où stocker les logs
if (-not (Test-Path "$CheminDossierLogs")) {
    New-Item "$CheminDossierLogs" -itemType Directory
}

## Archivage ##

# Création d'une première ligne dans le log pour l'archivage
Add-Content -Path $CheminFichierLog -Value "--- Archivage de comptes utilisateurs ---"

# Sélection de la liste des stagiaires qui ne se sont pas connectés depuis le nombre de jours limite d'archivage
$ListeStagiairesArchivage = Get-ADUser -Filter * -SearchBase $PathOUStagiaires -Properties LastLogon | `
    Select GivenName, Surname, SamAccountName, @{Name="OU";Expression={($_.DistinguishedName -split ",OU=",3)[1]}}, @{Name="LastLogon";Expression={[datetime]::FromFileTime($_.LastLogon)}} | `
    Where-Object {($_.LastLogon -lt ((Get-Date).AddDays(-$LimiteArchivage))) -and ($_.LastLogon -ne ([datetime]::FromFileTime(0)))}

# Désactive les comptes stagiaires inactifs depuis la date limite et les déplace dans l'OU Archive. Archive également les données
foreach ($Stagiaire in $ListeStagiairesArchivage) {

    # Chemin du dossier des données du stagiaire
    $CheminDossierStagiaire = $ModeleCheminDossierStagiaire + $Stagiaire.OU + "\" + $Stagiaire.Surname.ToUpper() + "_" + $Stagiaire.GivenName + "_(" + $Stagiaire.SamAccountName + ")"
    $CheminArchiveStagiaire = $CheminDossierArchive + $Stagiaire.Surname.ToUpper() + "_" + $Stagiaire.GivenName + "_(" + $Stagiaire.SamAccountName + ").zip"

    # Désactive le compte du stagiaire et le déplace dans l'OU Archives
    Set-ADUser -Identity $Stagiaire.SamAccountName -Enabled $false
    Get-ADUser -Identity $Stagiaire.SamAccountName | Move-ADObject -TargetPath $PathOUArchive

    # Compression du contenu du dossier personnel du stagiaire et enregistrement dans l'archive
    [IO.Compression.ZipFile]::CreateFromDirectory("$CheminDossierStagiaire", "$CheminArchiveStagiaire")

    # Suppression du dossier personnel du stagiaire
    Remove-Item "$CheminDossierStagiaire" -Recurse -Force

    # Ajout dans le log du compte archivé
    Add-Content -Path $CheminFichierLog -Value ("Le compte de l'utilisateur " + $Stagiaire.GivenName + " " + $Stagiaire.Surname + " a été archivé")
}

## Suppression ##

# Création d'un espacement et d'une première ligne dans le log pour la suppression
Add-Content -Path $CheminFichierLog -Value " "
Add-Content -Path $CheminFichierLog -Value "--- Suppression de comptes utilisateurs ---"

# Sélection de la liste des stagiaires archivés qui ne se sont pas connectés depuis le nombre de jours limite de suppression
$ListeStagiairesSuppression = Get-ADUser -Filter * -SearchBase $PathOUArchive -Properties LastLogon | `
    Select GivenName, Surname, SamAccountName, @{Name="LastLogon";Expression={[datetime]::FromFileTime($_.LastLogon)}} | `
    Where-Object {(-not $_.Enabled) -and ($_.LastLogon -lt ((Get-Date).AddDays(-$LimiteSuppression))) -and ($_.LastLogon -ne ([datetime]::FromFileTime(0)))}

# Suppression des comptes et des archives stagiaires inactifs depuis la date limite
foreach ($Stagiaire in $ListeStagiairesSuppression) {
    
    # Chemin de l'archive du stagiaire
    $CheminArchiveStagiaire = $CheminDossierArchive + $Stagiaire.Surname.ToUpper() + "_" + $Stagiaire.GivenName + "_(" + $Stagiaire.SamAccountName + ").zip"

    # Suppression du compte et de l'archive du stagiaire
    Remove-ADUser -Identity $Stagiaire.SamAccountName -Confirm:$false
    Remove-Item "$CheminArchiveStagiaire" -Force

    # Ajout dans le log du compte supprimé
    Add-Content -Path $CheminFichierLog -Value ("Le compte de l'utilisateur " + $Stagiaire.GivenName + " " + $Stagiaire.Surname + " a été supprimé")
}

# Ouverture du fichier de log à la fin de l'archivage
Invoke-Item $CheminFichierLog