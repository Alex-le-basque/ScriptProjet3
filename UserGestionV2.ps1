# Importer le module Active Directory
Import-Module ActiveDirectory

# Chemin vers le fichier CSV
$csvPath = "C:\\Users\\Administrator\\Desktop\\CSV_Master.csv"

# Lire le fichier CSV
$users = Import-Csv -Path $csvPath -Delimiter ','

# Fonction pour générer un DN unique
function Get-UniqueDN {
    param (
        [string]$baseDn
    )
    $dn = $baseDn
    $counter = 1
    while (Get-ADUser -LDAPFilter "(distinguishedName=$dn)") {
        $dn = "$baseDn $counter"
        $counter++
    }
    return $dn
}

# Fonction pour générer un SamAccountName unique
function Get-UniqueSamAccountName {
    param (
        [string]$samAccountName
    )
    $baseSamAccountName = $samAccountName
    $counter = 1
    while (Get-ADUser -Filter { SamAccountName -eq $samAccountName }) {
        $samAccountName = "$baseSamAccountName$counter"
        $counter++
    }
    return $samAccountName
}

# Fonction pour ajouter ou mettre à jour un utilisateur
function AddOrUpdateUser {
    param (
        [PSCustomObject]$user
    )

    # Validation des champs obligatoires
    if (-not $user.Prenom -or -not $user.Nom) {
        Write-Output "L'utilisateur avec des informations manquantes (Prenom ou Nom) ne peut pas être ajouté. Ligne: $($user | Out-String)"
        return
    }

    $prenom = $user.Prenom.Trim()
    $nom = $user.Nom.Trim()

    if (-not $prenom -or -not $nom) {
        Write-Output "L'utilisateur avec des informations manquantes après nettoyage (Prenom ou Nom) ne peut pas être ajouté. Ligne: $($user | Out-String)"
        return
    }

    $societe = $user.Societe
    $departement = $user.OU
    $service = $user.SousOU
    $groupDep = $user.Groupe_Departement
    $groupServ = $user.Groupe_Service
    $telephoneFixe = $user."Telephone fixe"
    $telephonePortable = $user."Telephone portable"

    # Définir l'OU
    $ouPath = if ($service -eq "NA") {
        "OU=$departement,OU=BillU-Users,DC=BILLU,DC=LAN"
    } else {
        "OU=$service,OU=$departement,OU=BillU-Users,DC=BILLU,DC=LAN"
    }

    # Construire le Distinguished Name (DN)
    $baseDn = "CN=$prenom $nom,$ouPath"
    $dn = Get-UniqueDN -baseDn $baseDn

    # Générer des noms d'utilisateur uniques
    $samAccountName = Get-UniqueSamAccountName -samAccountName ($prenom + "." + $nom).ToLower()
    $userPrincipalName = "$samAccountName@billu.lan"

    # Paramètres de l'utilisateur
    $userParams = @{
        SamAccountName    = $samAccountName
        UserPrincipalName = $userPrincipalName
        Name              = "$prenom $nom"
        GivenName         = $prenom
        Surname           = $nom
        DisplayName       = "$prenom $nom"
        Path              = $ouPath
        Enabled           = $true
        ChangePasswordAtLogon = $true
        PasswordNeverExpires  = $false
        Title             = $user.Fonction
        OfficePhone       = $telephoneFixe
        MobilePhone       = $telephonePortable
    }

    # Vérifier si l'utilisateur existe dans l'Active Directory
    $existingUser = Get-ADUser -Filter { SamAccountName -eq $samAccountName } -ErrorAction SilentlyContinue

    try {
        if ($existingUser) {
            # Mettre à jour l'utilisateur existant
            Set-ADUser @userParams
            Write-Output "Utilisateur $prenom $nom mis à jour avec succès à $ouPath"
        } else {
            # Ajouter l'utilisateur et ajouter aux groupes du département et du service
            New-ADUser @userParams
            Add-ADGroupMember -Identity $groupDep -Members $samAccountName
            Add-ADGroupMember -Identity $groupServ -Members $samAccountName
            Write-Output "Utilisateur $prenom $nom ajouté avec succès à $ouPath"
        }
    } catch {
        Write-Output "Erreur lors de la gestion de $prenom $nom : $_"
    }
}

# Boucle sur chaque ligne du CSV pour ajouter ou mettre à jour les utilisateurs
foreach ($user in $users) {
    AddOrUpdateUser -user $user
}

# Déplacer les utilisateurs qui ne sont pas présents dans le fichier CSV mais présents dans l'Active Directory
$adUsers = Get-ADUser -Filter * -SearchBase "OU=BillU-Users,DC=BILLU,DC=LAN" -Properties SamAccountName

foreach ($adUser in $adUsers) {
    $samAccountName = $adUser.SamAccountName

    # Vérifier si l'utilisateur est présent dans le fichier CSV
    $csvUser = $users | Where-Object { $_.Prenom.Trim() -eq $adUser.GivenName -and $_.Nom.Trim() -eq $adUser.Surname }

    if (-not $csvUser) {
        # Déplacer l'utilisateur vers l'OU spécifiée
        try {
            Move-ADObject -Identity $adUser.DistinguishedName -TargetPath "OU=CorbeilleOU=User-BillU,DC=BillU,DC=lan"
            Write-Output "Utilisateur $samAccountName déplacé vers l'OU Corbeille"
        } catch {
            Write-Output "Erreur lors du déplacement de $samAccountName : $_"
        }
    }
}
