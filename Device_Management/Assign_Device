# Fonction pour vérifier et installer les modules manquants
function Install-MissingModules {
    param (
        [string[]]$ModuleNames
    )

    foreach ($moduleName in $ModuleNames) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            Write-Host "Installation du module $moduleName"
            Install-Module -Name $moduleName -Scope AllUsers -Force -AllowClobber -ErrorAction Stop
        } else {
            Write-Host "Le module $moduleName est déjà installé"
        }
    }
}

# Vérifier et installer les modules manquants
try {
    Install-MissingModules -ModuleNames "Microsoft.Graph", "PnP.PowerShell", "PSWritePDF"
} catch {
    Write-Error "Erreur lors de l'installation des modules : $_"
    exit
}

# Importer les modules nécessaires
try {
    Import-Module Microsoft.Graph -ErrorAction Stop
    Import-Module PnP.PowerShell -ErrorAction Stop
    Import-Module PSWritePDF -ErrorAction Stop
} catch {
    Write-Error "Erreur lors de l'importation des modules : $_"
    exit
}

# Connexion à Microsoft Graph
try {
    Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All", "User.Read.All" -ErrorAction Stop
} catch {
    Write-Error "Erreur lors de la connexion à Microsoft Graph : $_"
    exit
}

# Définir les accessoires
$accessoire = ", "
$choix_Chargeur = Read-Host "Combien fournissez-vous de : Chargeur (0 pour aucun)"
if ($choix_Chargeur -ge 1) {
    $accessoire = $accessoire + " Chargeur $choix_Chargeur, "
}
$choix_adaptateur = Read-Host "Combien fournissez-vous de : adaptateur (0 pour aucun)"
if ($choix_adaptateur -ge 1) {
    $accessoire = $accessoire + " adaptateur $choix_adaptateur, "
}
$choix_casqueUSB = Read-Host "Combien fournissez-vous de : casque USB (0 pour aucun)"
if ($choix_casqueUSB -ge 1) {
    $accessoire = $accessoire + " Casque USB $choix_casqueUSB, "
}
$choix_casque_Sans_fils = Read-Host "Combien fournissez-vous de : casque Sans fils (0 pour aucun)"
if ($choix_casque_Sans_fils -ge 1) {
    $accessoire = $accessoire + " Casque Sans fils $choix_casque_Sans_fils, "
}
$choix_casque_Avaya = Read-Host "Combien fournissez-vous de : casque Téléphone Fixe Avaya (0 pour aucun)"
if ($choix_casque_Avaya -ge 1) {
    $accessoire = $accessoire + " Casque Avaya $choix_casque_Avaya,"
}
$choix_souris = Read-Host "Combien fournissez-vous de souris (0 pour aucun)"
if ($choix_souris -ge 1) {
    $accessoire = $accessoire + " Souris $choix_souris, "
}
$choix_tapis_souris = Read-Host "Combien fournissez-vous de : tapis de souris (0 pour aucun)"
if ($choix_tapis_souris -ge 1) {
    $accessoire = $accessoire + " Tapis de Souris $choix_tapis_souris, "
}
$choix_Clavier = Read-Host "Combien fournissez-vous de Clavier (0 pour aucun)"
if ($choix_Clavier -ge 1) {
    $accessoire = $accessoire + " Clavier $choix_Clavier, "
}
$choix_Sacoche = Read-Host "Combien fournissez-vous de Sacoche (0 pour aucun)"
if ($choix_Sacoche -ge 1) {
    $accessoire = $accessoire + " Sacoche $choix_Sacoche, "
}
$choix_Docking = Read-Host "Combien fournissez-vous de Docking (0 pour aucun)"
if ($choix_Docking -ge 1) {
    $accessoire = $accessoire + " Docking $choix_Docking, "
}
$choix_Ecran = Read-Host "Combien fournissez-vous de Écran (0 pour aucun)"
if ($choix_Ecran -ge 1) {
    $accessoire = $accessoire + " Écran $choix_Ecran"
}

Write-Host "Voici la liste des accessoires $accessoire"

# Saisie des informations utilisateur et ordinateur
$Email_salarie = Read-Host "Veuillez saisir l'adresse mail de l'utilisateur"
$computer_name = Read-Host "Veuillez saisir le nom de l'ordinateur (entier)"

# Récupérer les propriétés de l'ordinateur et de l'utilisateur
try {
    $computer_properties = Get-ADComputer -Filter { (samaccountname -like "*") } -Properties * | Where-Object { $_.Name -eq "$computer_name" } -ErrorAction Stop
    $userazure = Get-MgUser -Filter "proxyAddresses/any(c:c eq 'SMTP:$Email_salarie')" -ErrorAction Stop
} catch {
    Write-Error "Erreur lors de la récupération des propriétés de l'ordinateur ou de l'utilisateur : $_"
    exit
}

# Vérifier les informations de l'utilisateur et de l'ordinateur
if ($null -eq $userazure) {
    Write-Host "Mauvaise adresse $Email_salarie !!! Fin du script"
    exit
}

if ($null -eq $computer_properties) {
    Write-Host "Mauvais nom de poste $computer_name !!! Fin du script"
    $mauvaisNom = Read-Host "Est-ce un mac ? (n/autre)"
    if ($mauvaisNom -eq 'n') {
        exit
    }
}

try {
    $userproperties = Get-ADUser -Filter { (samaccountname -like "*") } -Properties * | Where-Object { $_.mail -eq "$Email_salarie" } -ErrorAction Stop
} catch {
    Write-Error "Erreur lors de la récupération des propriétés de l'utilisateur : $_"
    exit
}

# Détection de l'OS
$os = $computer_properties.OperatingSystem
$test_os = $os.Contains('Windows')
if ($test_os -eq $true) {
    $os = "Windows"
    Write-Host "Le poste est un Windows, il sera rangé dans l'OU Windows"
    $type = "Portable"
    $type2 = "Laptop"
} else {
    $os = "Mac"
    Write-Host "Le poste est un Mac, il sera rangé dans l'OU MAC"
}

$choix_portable = Read-Host "Est-ce un fixe ou un portable ? (f ou autre)"
if ($choix_portable -eq "f") {
    $type = "Fixe"
    $type2 = "Desktop"
} else {
    $type = "Portable"
    $type2 = "Laptop"
}

Write-Host "Le type défini est $type et l'OU sera donc $type2"

# Déplacer l'ordinateur dans l'OU appropriée
try {
    $computer_DistinguishedName = $computer_properties.DistinguishedName
    $DistinguishedName = $userproperties.DistinguishedName
    $index = $DistinguishedName.IndexOf("Account")
    $site = $DistinguishedName.Substring($index + 11, 9)
    Write-Host "Découpe : $site"
    $new_computer_DistinguishedName = "OU=$type2,OU=$os,OU=Computer,OU=$site,OU=ENTIS,DC=cetremut,DC=pri"
    Write-Host "DistinguishedName : $computer_DistinguishedName et new : $new_computer_DistinguishedName"
    Move-ADObject $computer_DistinguishedName -TargetPath $new_computer_DistinguishedName -ErrorAction Stop
} catch {
    Write-Error "Erreur lors du déplacement de l'ordinateur dans l'OU : $_"
    exit
}

# Récupérer les informations de l'appareil dans Intune
try {
    $deviceintunes = Get-MgDeviceManagementManagedDevice -Select devicename, id, userDisplayName, lastSyncDateTime | Sort-Object DeviceName -ErrorAction Stop
    foreach ($deviceintune in $deviceintunes) {
        $pc = $deviceintune.deviceName
        if ($pc -eq $computer_name) {
            $deviceId = $deviceintune.id
        }
    }
} catch {
    Write-Error "Erreur lors de la récupération des informations de l'appareil dans Intune : $_"
    exit
}

# Obtenir les informations de l'appareil
try {
    $device = Get-MgDeviceManagementManagedDevice -ManagedDeviceId $deviceId -ErrorAction Stop
    $fabricant = $device.Manufacturer
    $model = $device.Model
    $serialnumber = $device.SerialNumber
} catch {
    Write-Error "Erreur lors de la récupération des informations de l'appareil : $_"
    exit
}

# Générer le nom du fichier PDF
$date = Get-Date -Format "yyyyMMdd"
$pdfFileName = "${date}_${Email_salarie}_${computer_name}.pdf"
$pdfOutputPath = "C:\temp\$pdfFileName"

# Créer un fichier PDF avec PSWritePDF
try {
    New-PDF {
        # Ajouter du texte au document
        New-PDFText -Text "Fiche d'attribution de matériel"
        New-PDFText -Text "Utilisateur: $($userazure.UserPrincipalName)"
        New-PDFText -Text "Ordinateur: $computer_name"
        New-PDFText -Text "Fabricant: $fabricant"
        New-PDFText -Text "Modèle: $model"
        New-PDFText -Text "Numéro de série: $serialnumber"
        New-PDFText -Text "Accessoires: $accessoire"
    } -FilePath $pdfOutputPath -ErrorAction Stop
} catch {
    Write-Error "Erreur lors de la création du fichier PDF : $_"
    exit
}

Write-Host "PDF généré avec succès à l'emplacement : $pdfOutputPath"

# Connexion à SharePoint et téléchargement du fichier
try {
    $siteURL = "https://entis.sharepoint.com/sites/glpi"
    $clientId = "eee6a5ff-706f-4ad3-bd4f-00bc0674b1af"
    Connect-PnPOnline -Url $siteURL -ClientId $clientId -ErrorAction Stop
    $folder = "Documents%20partages"
    $file = Add-PnPFile -Path $pdfOutputPath -Folder $folder -ErrorAction Stop
    Write-Host "Fichier PDF téléchargé sur SharePoint avec succès."
} catch {
    Write-Error "Erreur lors du téléchargement du fichier sur SharePoint : $_"
    exit
}

# Attribution Intune
try {
    Add-ADGroupMember GC_MDM_Intune_Compliance_Pilot -Members $computer -ErrorAction Stop
} catch {
    Write-Error "Erreur lors de l'attribution Intune : $_"
    exit
}
