<#
Titre : Création d'utilisateurs stagiaires dans l'Active Directory
Date de création : 12/05/2022
Auteur : Anthony Quesnel
Version : 1.0
Date de modification : -
Modifications effectuées : -
Résumé du fonctionnement du script :
    - L'utilisateur choisit le numéro de la promotion et le fichier CSV à importer
    - Les utilisateurs sont importés via le fichier CSV
    - Création d'un utilisateur dans l'Active Directory enregistré dans son OU de destination
    - Création du dossier partagé et association au compte AD du stagiaire
#>

### Chargement de librairies ###
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

### Création de fonctions ###

# Fonction permettant d'ajouter une lettre non accentuée à un identifiant
Function Add-LettreIdentifiant {
    param (
        $Lettre,
        $Identifiant
    )

    # Vérifie si la lettre est accentuée ou non. Si la lettre est accentuée elle est remplacée par le caractère non-accentué correspondant
    switch -Regex ($Lettre) {
        "[a-z]" {
            $Identifiant += $Lettre
        }
        "[áàâäãå]" {
            $Identifiant += "a"
        }
        "ç" {
            $Identifiant += "c"
        }
        "[éèêë]" {
            $Identifiant += "e"
        }
        "[íìîï]" {
            $Identifiant += "i"
        }
        "ñ" {
            $Identifiant += "n"
        }
        "[óòôöõ]" {
            $Identifiant += "o"
        }
        "[úùûü]" {
            $Identifiant += "u"
        }
        "[ýÿ]" {
            $Identifiant += "y"
        }
        "æ" {
            $Identifiant += "ae"
        }
        "œ" {
            $Identifiant += "oe"
        }
    }

    # Envoi de l'identifiant modifié
    return $Identifiant
}

### Déclaration des variables ###

# Chemin de l'OU Stagiaires dans l'Active Directory
$NomOUStagiaires = "Stagiaires"
$NomOUUsers = "Users"
$NomOUParent = "E2CN"
$NomDomaineAD = "aqtest"
$ExtensionDomaineAD = "loc"
$PathOUStagiaires = "OU=$NomOUStagiaires,OU=$NomOUUsers,OU=$NomOUParent,DC=$NomDomaineAD,DC=$ExtensionDomaineAD"

$ModeleNomOUPromo = "Promo "

# Mot de passe générique
$MotDePasse = ConvertTo-SecureString -String ("E2CN@" + (Get-Date -Format "yyyy")) -AsPlainText -Force

# Variables du dossier partagé du stagiaire
$ModeleCheminDossierStagiaire = "C:\DossiersReseauStagiaires\"
$ModeleCheminPartageDossierStagiaire = "\\SRV-AQTEST\"

# Fenêtre de choix du fichier CSV à importer
$ChoixFichierCSV = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    Title = "Choix du fichier CSV à importer"
    InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    Filter = "Fichiers CSV (*.csv)|*.csv"
}

### Lancement des commandes ###

# Choix de la promotion par l'utilisateur
do {
    $NumeroPromo = [Microsoft.VisualBasic.Interaction]::InputBox("Entrez le numéro de promotion des stagiaires", "Création de comptes stagiaires")
} until ($NumeroPromo -match "^\d*$" -or $NumeroPromo -match "^$")

# Sortie du script si aucune entrée ou annulation
if ($NumeroPromo -match "^$") {
    Write-Host "Programme arrêté par l'utilisateur"
    exit
}

# Vérification de l'existence de la promotion dans l'AD
if (Get-ADOrganizationalUnit -Filter "distinguishedName -eq 'OU=$ModeleNomOUPromo$NumeroPromo,$PathOUStagiaires'") {
    [System.Windows.Forms.MessageBox]::Show("L'OU $ModeleNomOUPromo$NumeroPromo existe déjà","",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
}

# Choix du fichier CSV à importer
$ChoixFichierCSV.ShowDialog()

# Sortie du script si aucune entrée ou annulation
if ($ChoixFichierCSV.FileName -match "^$") {
    Write-Host "Programme arrêté par l'utilisateur"
    exit
}

# Création de l'OU et du groupe des stagiaires si inexistants
if (-not (Get-ADOrganizationalUnit -Filter "distinguishedName -eq 'OU=$ModeleNomOUPromo$NumeroPromo,$PathOUStagiaires'")) {
    # Création de l'OU si elle n'existe pas
    New-ADOrganizationalUnit -Name "$ModeleNomOUPromo$NumeroPromo" -Path "$PathOUStagiaires"
}

if (-not (Get-ADGroup -Filter "Name -eq '$ModeleNomOUPromo$NumeroPromo'")) {
    # Création de l'OU si elle n'existe pas
    New-ADGroup -Name "$ModeleNomOUPromo$NumeroPromo" -GroupScope Global -Path "OU=$ModeleNomOUPromo$NumeroPromo,$PathOUStagiaires"
}

# Importation de la liste des stagiaires
try {
	$ListeStagiaires = Import-CSV -Delimiter ';' -Encoding Default -Path $ChoixFichierCSV.FileName | Select-Object Prénom, Nom
}
catch {
	[System.Windows.Forms.MessageBox]::Show("Erreur dans l'importation de la liste des utilisateurs","",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Ajout des utilisateurs dans l'Active Directory
foreach ($Stagiaire in $ListeStagiaires) {

    ## Identifiant Stagiaire ##

    # Initialisation de l'identifiant du stagiaire
    $IDStagiaire = ""

    # Ajout de la première lettre du prénom à l'identifiant
    $IDStagiaire = Add-LettreIdentifiant $Stagiaire.Prénom.Substring(0,1).ToLower() $IDStagiaire

    # Ajout des lettres du nom de famille à l'identifiant
    for ($i = 0; $i -lt $Stagiaire.Nom.Length; $i++) {
        
        # fin de la génération de l'identifiant si le premier mot d'un nom de famille est fini (hors particule) ou si l'identifiant a une longueur de 8 caractères
        if ((($Stagiaire.Nom.Substring($i,1).ToLower() -match "[ \-]") -and ($IDStagiaire.Length -gt 3)) -or ($IDStagiaire.Length -ge 8)) {
            break
        }
        else {
            $IDStagiaire = Add-LettreIdentifiant $Stagiaire.Nom.Substring($i,1).ToLower() $IDStagiaire
        }
    }

    # Vérification que l'identifiant n'est pas présent dans l'Active Directory
    if ((Get-ADUser -Filter "SamAccountName -eq '$IDStagiaire'")) {
        # Initialisation d'un numéro à incrémenter
        $Numero = 0

        do {
            # Incrémentation du numéro jusqu'à trouver un identifiant n'existant pas dans l'Active Directory
            $Numero++

            # Création de l'identifiant à tester
            if (("$IDStagiaire" + "$Numero").Length -gt 8) {
                $IDStagiaireTest = $IDStagiaire.Substring(0,(8 - $Numero.Length)) + $Numero
            }
            else {
                $IDStagiaireTest = $IDStagiaire + $Numero
            }            
        } while (Get-ADUser -Filter "SamAccountName -eq '$IDStagiaireTest'")

        # Envoi de l'identifiant final
        $IDStagiaire = $IDStagiaireTest
    }

    # Vérification que le nom d'utilisateur n'est pas présent dans l'Active Directory
    $NomCompletStagiaire = $Stagiaire.Prénom + " " + $Stagiaire.Nom

    if ((Get-ADUser -Filter "Name -eq '$NomCompletStagiaire'")) {
        $NomCompletStagiaire = $NomCompletStagiaire + $Numero
    }

    ## Ajout d'un utilisateur dans l'AD ##

    # Initialisation des chemins de dossier du stagiaire
    $NomPartageDossierStagiaire = $IDStagiaire + "$"
    $CheminDossierStagiaire = "$ModeleCheminDossierStagiaire$ModeleNomOUPromo$NumeroPromo\" + $Stagiaire.Nom.ToUpper() + "_" + $Stagiaire.Prénom + "_($IDStagiaire)"

	#Création du compte utilisateur du stagiaire dans l'Active Directory et placement dans son Unité d'Organisation
    try {
		New-ADUser `
            -Name $NomCompletStagiaire `
            -GivenName $Stagiaire.Prénom `
            -Surname $Stagiaire.Nom `
            -DisplayName ($Stagiaire.Prénom + " " + $Stagiaire.Nom) `
            -SamAccountName "$IDStagiaire" `
            -UserPrincipalName "$IDStagiaire@$NomDomaineAD.$ExtensionDomaineAD" `
            -Accountpassword $MotDePasse `
            -ChangePasswordAtLogon $true `
            –HomeDrive 'P:' `
            –HomeDirectory "$ModeleCheminPartageDossierStagiaire$NomPartageDossierStagiaire" `
            -Path "OU=$ModeleNomOUPromo$NumeroPromo,$PathOUStagiaires" `
            -Enabled $true
	}
	catch {
		[System.Windows.Forms.MessageBox]::Show(("Erreur dans la création de l'utilisateur " + $Stagiaire.Prénom + " " + $Stagiaire.Nom),"",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
		exit		
	}

    #Ajout de l'utilisateur à son groupe de sécurité
    try {
		Add-AdGroupMember -Identity "$ModeleNomOUPromo$NumeroPromo" -Members $IDStagiaire
    }
    catch {
		[System.Windows.Forms.MessageBox]::Show(("Erreur dans l'ajout de l'utilisateur " + $Stagiaire.Prénom + " " + $Stagiaire.Nom + " dans le groupe $ModeleNomOUPromo$NumeroPromo"),"",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
		exit		
    }

    # Création du dossier utilisateur
    try {
		New-Item "$CheminDossierStagiaire" -itemType Directory
	}
	catch {
		[System.Windows.Forms.MessageBox]::Show(("Erreur dans la création du dossier personnel du stagiaire " + $Stagiaire.Prénom + " " + $Stagiaire.Nom),"",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
		exit
	}

	# Partage du dossier utilisateur
    try {
		New-SmbShare -Name $NomPartageDossierStagiaire -Path $CheminDossierStagiaire -ChangeAccess $IDStagiaire
    }
    catch {
		[System.Windows.Forms.MessageBox]::Show(("Erreur dans le partage du dossier du stagiaire " + $Stagiaire.Prénom + " " + $Stagiaire.Nom),"",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
		exit		
    }
}

# Fin de l'import
[System.Windows.Forms.MessageBox]::Show("La création des comptes stagiaires de la $ModeleNomOUPromo$NumeroPromo a été effectuée avec succès.")