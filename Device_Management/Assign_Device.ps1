#region [Fonction] Installation des modules manquants
function Install-MissingModules {
    param ([string[]]$ModuleNames)
    foreach ($moduleName in $ModuleNames) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            Write-Host "Installation du module $moduleName"
            Install-Module -Name $moduleName -Scope AllUsers -Force -AllowClobber
        } else {
            Write-Host "Le module $moduleName est déjà installé"
        }
    }
}
Install-MissingModules -ModuleNames "Microsoft.Graph", "PnP.PowerShell", "PSWritePDF"
#endregion

#region [Import] Modules et authentification
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.DeviceManagement
Import-Module PnP.PowerShell
Import-Module PSWritePDF

# Authentification Microsoft Graph
Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All", "User.Read.All"

# Authentification Sharepoit
$siteURL = "https://entis.sharepoint.com/sites/glpi"
$folder ="Documents%20partages"
$clientId = "eee6a5ff-706f-4ad3-bd4f-00bc0674b1af"
Connect-PnPOnline -Url $siteURL -ClientId $clientId -Interactive

#endregion

#region [Fonction] Saisie des accessoires
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
#endregion

#region [Entrées] Utilisateur + ordinateur
$email = Read-Host "Adresse mail de l'utilisateur"
$pcName = Read-Host "Nom complet du poste"
#endregion

#region [Récup] Utilisateur & ordinateur AD / Graph
try {
    # Récupérer l'utilisateur Azure
    $userAzure = Get-MgUser -Filter "proxyAddresses/any(c:c eq 'SMTP:$email')"
    if (-not $userAzure) {
        throw "Utilisateur Azure inexistant."
    }

    # Récupérer l'utilisateur Active Directory
    $userAD = Get-ADUser -Filter "mail -eq '$email'" -Properties *
    if (-not $userAD) {
        throw "Utilisateur AD inexistant."
    }

    # Récupérer les propriétés de l'ordinateur
    $computerAD = Get-ADComputer -Identity $pcName -Properties *
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
#endregion

#region [Détection OS et type]
$type2 = "Laptop"
$type = "Portable"
$os = ($computerAD.OperatingSystem -match "Windows") ? "Windows" : "Mac"

if ($os -eq "Mac") {
    Write-Host "Poste Mac. Rangé dans OU MAC."
} else {
    Write-Host "Poste Windows."
}

$choix = Read-Host "Est-ce un fixe ou portable ? (f ou p)"
if ($choix -eq "f") {
    $type = "Fixe"
    $type2 = "Desktop"
}
Write-Host "Type: $type / OU cible: $type2"
#endregion

#region [Déplacement de l'ordinateur dans la bonne OU]
$dnUser = $userAD.DistinguishedName
if ($dnUser -match "OU=([^,]+),OU=ENTIS") {
    $site = $matches[1]
    $newOU = "OU=$type2,OU=$os,OU=Computer,OU=$site,OU=ENTIS,DC=cetremut,DC=pri"
    Move-ADObject -Identity $computerAD.DistinguishedName -TargetPath $newOU
    Write-Host "Poste déplacé dans : $newOU"
} else {
    Write-Error "Impossible de déterminer l'OU du site"
}
#endregion

#region [Intune] Récupération info poste
$device = Get-MgDeviceManagementManagedDevice -Filter "deviceName eq '$pcName'" `
    -Select "id,deviceName,manufacturer,model,serialNumber" | Select-Object -First 1

if (-not $device) {
    Write-Error "Poste introuvable dans Intune."
    exit 1
}
#endregion

#region [PDF] Génération
$date = Get-Date -Format "yyyyMMdd"
$pdfFile = "C:\\temp\\${date}_${pcName}.pdf"
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
#endregion

#region [SharePoint] Upload du PDF
Add-PnPFile -Path $pdfFile -Folder $folder
Write-Host "Upload SharePoint terminé."
#endregion

#region [Intune] Ajout au groupe de conformité
Add-ADGroupMember -Identity "GC_MDM_Intune_Compliance_Pilot" -Members $pcName
#endregion