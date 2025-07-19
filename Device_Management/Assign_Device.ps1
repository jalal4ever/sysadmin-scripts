# Vérification et installation des modules manquants
function Install-MissingModules {
    param ([string[]]$ModuleNames)
    foreach ($moduleName in $ModuleNames) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            Write-Host "Installation du module $moduleName"
            Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber
        } else {
            Write-Host "Le module $moduleName est déjà installé"
        }
    }
}

Install-MissingModules -ModuleNames "Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.DeviceManagement", "PnP.PowerShell", "PSWritePDF"

# Importation des modules
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.DeviceManagement
Import-Module PnP.PowerShell
Import-Module PSWritePDF
Import-Module ActiveDirectory


# Authentification Microsoft Graph
Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All", "User.Read.All"

# Authentification SharePoint
$siteURL = "https://entis.sharepoint.com/sites/glpi"
$folder = "Documents%20partages"
$clientId = "eee6a5ff-706f-4ad3-bd4f-00bc0674b1af"
Connect-PnPOnline -Url $siteURL -ClientId $clientId -Interactive

# Fonction pour la saisie des accessoires
function Get-Accessoires {
    $liste = @(
        @{ nom = "Chargeur" },
        @{ nom = "Adaptateur" },
        @{ nom = "Casque USB" },
        @{ nom = "Casque Sans fils" },
        @{ nom = "Casque Avaya" },
        @{ nom = "Souris" },
        @{ nom = "Tapis de souris" },
        @{ nom = "Clavier" },
        @{ nom = "Sacoche" },
        @{ nom = "Docking" },
        @{ nom = "Écran" }
    )
    $accessoire = ""
    foreach ($item in $liste) {
        $count = Read-Host "Combien fournissez-vous de : $($item.nom) (0 pour aucun)"
        if ([int]$count -ge 1) {
            $accessoire += " $($item.nom) $count,"
        }
    }
    return $accessoire.Trim(',')
}

# Saisie des informations utilisateur et ordinateur
$email = Read-Host "Adresse mail de l'utilisateur"
$pcName = Read-Host "Nom complet du poste"

# Récupération des informations utilisateur et ordinateur
$userAzure = Get-MgUser -Filter "proxyAddresses/any(c:c eq 'SMTP:$email')"
if (-not $userAzure) {
    Write-Error "Utilisateur Azure inexistant."
    exit 1
}

$userAD = Get-ADUser -Filter "mail -eq '$email'" -Properties *
if (-not $userAD) {
    Write-Error "Utilisateur AD inexistant."
    exit 1
}

$computerAD = Get-ADComputer -Identity $pcName -Properties *
if (-not $computerAD) {
    Write-Error "Ordinateur inexistant."
    exit 1
}

# Détection de l'OS et du type
$os = if ($computerAD.OperatingSystem -match "Windows") { "Windows" } else { "Mac" }
$type = if ($os -eq "Mac") { "Mac" } else { "Windows" }
$type2 = Read-Host "Est-ce un fixe ou un portable ? (f ou p)" | ForEach-Object {
    if ($_ -eq "f") { "Desktop" } else { "Laptop" }
}

# Déplacement de l'ordinateur dans la bonne OU
$dnUser = $userAD.DistinguishedName
if ($dnUser -match "OU=([^,]+),OU=ENTIS") {
    $site = $matches[1]
    $newOU = "OU=$type2,OU=$type,OU=Computer,OU=$site,OU=ENTIS,DC=cetremut,DC=pri"
    Move-ADObject -Identity $computerAD.DistinguishedName -TargetPath $newOU
    Write-Host "Poste déplacé dans : $newOU"
} else {
    Write-Error "Impossible de déterminer l'OU du site"
    exit 1
}

# Récupération des informations de l'appareil dans Intune
$device = Get-MgDeviceManagementManagedDevice -Filter "deviceName eq '$pcName'" -Select "id,deviceName,manufacturer,model,serialNumber" | Select-Object -First 1
if (-not $device) {
    Write-Error "Poste introuvable dans Intune."
    exit 1
}

# Génération du PDF
$date = Get-Date -Format "yyyyMMdd"
$pdfFile = "C:\temp\${date}_${pcName}.pdf"
$accessoires = Get-Accessoires
New-PDF {
    New-PDFText -Text "Fiche d'attribution de matériel"
    New-PDFText -Text "Utilisateur : $($userAzure.UserPrincipalName)"
    New-PDFText -Text "Ordinateur : $pcName"
    New-PDFText -Text "Fabricant : $($device.Manufacturer)"
    New-PDFText -Text "Modèle : $($device.Model)"
    New-PDFText -Text "Numéro de série : $($device.SerialNumber)"
    New-PDFText -Text "Accessoires : $accessoires"
} -FilePath $pdfFile

Write-Host "PDF généré : $pdfFile"

# Upload du PDF sur SharePoint
Add-PnPFile -Path $pdfFile -Folder $folder
Write-Host "Upload SharePoint terminé."

# Ajout au groupe de conformité Intune
Add-ADGroupMember -Identity "GC_MDM_Intune_Compliance_Pilot" -Members $pcName