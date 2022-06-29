### Chargement de librairies ###
Add-Type -AssemblyName System.Windows.Forms

### Création des fenêtres ###
Function New-BoiteEntreeInfosStagiaire {

    # Création de la fenêtre
    $BoiteEntreeInfosStagiaire = New-Object System.Windows.Forms.Form 
    $BoiteEntreeInfosStagiaire.Text = "Création de comptes stagiaires"
    $BoiteEntreeInfosStagiaire.Size = New-Object System.Drawing.Size(400,170)
    $BoiteEntreeInfosStagiaire.StartPosition = 'CenterScreen'

    # Création d'un label pour le champ d'entrée du prénom
    $LabelEntreePrenom = New-Object System.Windows.Forms.Label
    $LabelEntreePrenom.Location = New-Object System.Drawing.Point(5,7)
    $LabelEntreePrenom.Size = New-Object System.Drawing.Size(245,20)
    $LabelEntreePrenom.Text = "Prénom du stagiaire:"

    # Création d'un champ d'entrée texte du prénom
    $ChampEntreePrenom = New-Object System.Windows.Forms.TextBox
    $ChampEntreePrenom.Location = New-Object System.Drawing.Point(250,5)
    $ChampEntreePrenom.Size = New-Object System.Drawing.Size(125,20)

    # Création d'un label pour le champ d'entrée du nom de famille
    $LabelEntreeNom = New-Object System.Windows.Forms.Label
    $LabelEntreeNom.Location = New-Object System.Drawing.Point(5,37)
    $LabelEntreeNom.Size = New-Object System.Drawing.Size(245,20)
    $LabelEntreeNom.Text = "Nom de famille du stagiaire:"

    # Création d'un champ d'entrée texte du nom de famille
    $ChampEntreeNom = New-Object System.Windows.Forms.TextBox
    $ChampEntreeNom.Location = New-Object System.Drawing.Point(250,35)
    $ChampEntreeNom.Size = New-Object System.Drawing.Size(125,20)

    # Création d'un label pour le champ d'entrée de la promo
    $LabelEntreePromo = New-Object System.Windows.Forms.Label
    $LabelEntreePromo.Location = New-Object System.Drawing.Point(5,67)
    $LabelEntreePromo.Size = New-Object System.Drawing.Size(245,20)
    $LabelEntreePromo.Text = "Numéro de la promotion du stagiaire:"

    # Création d'un champ d'entrée texte de la promo
    $ChampEntreePromo = New-Object System.Windows.Forms.TextBox
    $ChampEntreePromo.Location = New-Object System.Drawing.Point(250,65)
    $ChampEntreePromo.Size = New-Object System.Drawing.Size(125,20)


    # Création d'un bouton OK
    $BoutonOK = New-Object System.Windows.Forms.Button
    $BoutonOK.Location = New-Object System.Drawing.Point(220,95)
    $BoutonOK.Size = New-Object System.Drawing.Size(75,23)
    $BoutonOK.Text = 'OK'
    $BoutonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    # Création d'un bouton Annuler
    $BoutonAnnuler = New-Object System.Windows.Forms.Button
    $BoutonAnnuler.Location = New-Object System.Drawing.Point(300,95)
    $BoutonAnnuler.Size = New-Object System.Drawing.Size(75,23)
    $BoutonAnnuler.Text = 'Annuler'
    $BoutonAnnuler.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    # Ajout des éléments créés à la fenêtre
    $BoiteEntreeInfosStagiaire.Controls.Add($LabelEntreePrenom)
    $BoiteEntreeInfosStagiaire.Controls.Add($ChampEntreePrenom)

    $BoiteEntreeInfosStagiaire.Controls.Add($LabelEntreeNom)
    $BoiteEntreeInfosStagiaire.Controls.Add($ChampEntreeNom)

    $BoiteEntreeInfosStagiaire.Controls.Add($LabelEntreePromo)
    $BoiteEntreeInfosStagiaire.Controls.Add($ChampEntreePromo)

    $BoiteEntreeInfosStagiaire.AcceptButton = $BoutonOK
    $BoiteEntreeInfosStagiaire.Controls.Add($BoutonOK)

    $BoiteEntreeInfosStagiaire.CancelButton = $BoutonAnnuler
    $BoiteEntreeInfosStagiaire.Controls.Add($BoutonAnnuler)

    # Met la boite de dialogue au premier plan à son ouverture
    $BoiteEntreeInfosStagiaire.Topmost = $true

    # Affiche la boite de dialogue
    $result = $BoiteEntreeInfosStagiaire.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        # Envoi des données entrée par l'utilisateur sous forme d'un tableau s'il appuye sur le bouton OK
        #[0] = Prénom, [1] = Nom, [2] = Promo
        return $ChampEntreePrenom.Text, $ChampEntreeNom.Text, $ChampEntreePromo.Text
    }
    else
    {
        exit
    }
}

### Déclaration des variables ###

# Chemin de l'OU Stagiaires dans l'Active Directory
$NomOUStagiares = "Elèves"
$NomOUUsers = "Utilisateurs"
$NomOUParent = "KoXoAdm"
$NomDomaineAD = "aqtest"
$ExtensionDomaineAD = "loc"
$PathOUStagiares = "OU=$NomOUStagiares,OU=$NomOUUsers,OU=$NomOUParent,DC=$NomDomaineAD,DC=$ExtensionDomaineAD"

$ModeleNomOUPromo = "Promo "

# Mot de passe générique
$MotDePasse = ConvertTo-SecureString -String ("E2CN@" + (Get-Date -Format "yyyy")) -AsPlainText -Force

### Lancement des commandes ###

# Entrée des infos du stagiaire par l'utilisateur
$InfosStagiaire = New-BoiteEntreeInfosStagiaire

# Vérification de l'existence de l'utilisateur dans l'AD
try {
    $Stagiaire = Get-ADUser -Filter ("(GivenName -eq '" + $InfosStagiaire[0] + "') -and (Surname -eq '" + $InfosStagiaire[1] + "')") -SearchBase ("OU=$ModeleNomOUPromo" + $InfosStagiaire[2] + ",$PathOUStagiares") | Select GivenName, Surname, SamAccountName

    # Changement du mot de passe et demande de réinitialisation à la première connexion
    Set-ADAccountPassword -Identity $Stagiaire.SamAccountName -NewPassword $MotDePasse
    Set-ADUser -Identity $Stagiaire.SamAccountName -ChangePasswordAtLogon $true -PasswordNeverExpires $false

    [System.Windows.Forms.MessageBox]::Show("Mot de passe de l'utilisateur " + $Stagiaire.GivenName + " " + $Stagiaire.Surname + " réinitialisé.")
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Aucun utilisateur trouvé","",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}