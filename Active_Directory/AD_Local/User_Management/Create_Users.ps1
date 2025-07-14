# encoding: utf-8

<#
.SYNOPSIS
    Script for creating Active Directory users from a CSV file.

.DESCRIPTION
    This script automates the creation of user accounts in Active Directory based on data provided in a CSV file.
    It includes functionality to clean and format user data, validate essential fields, and log the creation process.
    Version: 1.1
    Created: 2023-10-13
    Author: JEL for Entis Mutuelles

.PREREQUISITES
    - PowerShell 5.1 or later.
    - Active Directory module for PowerShell must be installed and available.
    - The script must be run with sufficient permissions to create users in Active Directory.
    - CSV files must be prepared and placed in the correct directories:
        - A CSV file named "Nouvel arrivant*.csv" containing user data must be present in the script's root directory.
        - A CSV file named "Adresse_Agences.csv" containing agency addresses must be present in the "Data" subdirectory.
        - A CSV file named "Groupe_service.csv" containing service groups must be present in the "Data" subdirectory.

.INPUTS
    - "Nouvel arrivant*.csv": Contains user information such as first name, last name, location, etc.
    - "Adresse_Agences.csv": Contains agency details like OU, phone, address, etc.
    - "Groupe_service.csv": Contains service group details.

.OUTPUTS
    - Logs: The script generates log files in a "Logs" directory located in the root directory where the script is executed.
        - Log files are named with the current date in the format "yyyyMMdd_execution_script_result.txt".
        - These logs contain details about the script execution, including attempts to create users, successes, and errors.
    - Processed User File: A file named with the current date in the format "yyyyMMdd_processed_created_users.txt" is generated in the "Logs" directory.
        - This file contains the processed user data.
    - Support Ticket Message: A formatted message is generated for each user created, which can be copied and pasted directly into a support ticket.
        - This message includes the user's full name, username (SamAccountName), email address, and password.

.NOTES
    Future Improvements:
    - Generate a ready-to-share message with user credentials for easy communication.
    - The default password is now randomized to 16 characters for each user.
#>

Import-Module ActiveDirectory

# Function to generate a random password
function Generate-RandomPassword {
    $chars = 'ABCDEFGHJKLMNOPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz0123456789-!?@_'
    $passwordLength = 16
    $random = New-Object System.Random
    $password = -join (1..$passwordLength | ForEach-Object { $chars[$random.Next($chars.Length)] })
    return $password
}

# Function to clean strings by removing special characters
function Clean-String {
    param (
        [Parameter(Mandatory)]
        [string]$String,
        [Parameter(Mandatory)]
        [bool]$ReplaceSpacesWithUnderscore
    )

    if ([string]::IsNullOrEmpty($String)) {
        return $String
    }

    $cleanString = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
    $cleanString = $cleanString -replace "'", ""

    if ($ReplaceSpacesWithUnderscore) {
        $cleanString = $cleanString -replace " ", "_"
    }

    return $cleanString
}

# Function to remove accents from strings
function Remove-Accent {
    param (
        [Parameter(Mandatory)]
        [string]$String
    )

    if ([string]::IsNullOrEmpty($String)) {
        return $String
    }

    $objD = $String.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object Text.StringBuilder
    for ($i = 0; $i -lt $objD.Length; $i++) {
        $c = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($objD[$i])
        if ($c -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($objD[$i])
        }
    }
    return "$sb".Normalize([Text.NormalizationForm]::FormC)
}

# Function to validate file path
function Test-FilePath {
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path)) {
        Write-Error "The file path '$Path' does not exist."
        return $false
    }

    return $true
}

# Get the current directory where the script is executed
$currentDirectory = Get-Location

# Define paths relative to the current directory
$logDirectory = Join-Path -Path $currentDirectory -ChildPath "Logs"
$dataDirectory = Join-Path -Path $currentDirectory -ChildPath "Data"

# Ensure Logs directory exists
if (-not (Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory | Out-Null
}

$path_new_users_exchange = Join-Path -Path $currentDirectory -ChildPath "new_users_exchange.csv"
$path = Join-Path -Path $currentDirectory -ChildPath "new_users.csv"

# Clear content of output files if user confirms
$choix_rename = Read-Host "Voulez-vous écraser le fichier <new_users.csv> (o/autre)"
if ($choix_rename -eq 'o') {
    Clear-Content -Path $path
}

$choix_rename = Read-Host "Voulez-vous écraser le fichier <new_users_exchange.csv> (o/autre)"
if ($choix_rename -eq 'o') {
    Clear-Content -Path $path_new_users_exchange
}

# Rename the input file
$inputFile = Get-ChildItem -Path "$currentDirectory\Nouvel arrivant*.csv" | Select-Object -First 1
if ($inputFile) {
    $newInputFileName = "$(Get-Date -Format 'yyyyMMdd')_processed_created_users.txt"
    $newInputFilePath = Join-Path -Path $logDirectory -ChildPath $newInputFileName
    Move-Item -Path $inputFile.FullName -Destination $newInputFilePath | Out-Null
    $path_Export_sn = $newInputFilePath
} else {
    Write-Error "No input file found matching 'Nouvel arrivant*.csv'."
    exit
}

# Import CSV files
$users = Import-CSV $path_Export_sn -Header COL1, COL2, COL3, COL4, COL5, COL6, COL7, COL8, COL9, COL10, COL11, COL12, COL13, COL14, COL15, COL16, COL17, COL18, COL19, COL20, COL21, COL22, COL23, COL24, COL25, COL26, COL27, COL28, COL29, COL30, COL31, COL32, COL33, COL34, COL35, COL36, COL37, COL38, COL39, COL40, COL41, COL42, COL43, COL44, COL45, COL46 -Delimiter ";" -Encoding UTF8
$lieux_csv = Import-CSV (Join-Path -Path $dataDirectory -ChildPath "Adresse_Agences.csv") -Delimiter ","
$Services_csv = Import-CSV (Join-Path -Path $dataDirectory -ChildPath "Groupe_service.csv") -Delimiter ";" -Encoding UTF8

$logFilePath = Join-Path -Path $logDirectory -ChildPath "$(Get-Date -Format 'yyyyMMdd')_execution_script_result.txt"
$processedFilePath = Join-Path -Path $logDirectory -ChildPath "$(Get-Date -Format 'yyyyMMdd')_processed_created_users.txt"

# Log the start of the script
$logMessage = "Script started on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
$logMessage | Out-File -FilePath $logFilePath -Append

foreach ($user in $users) {
    if ($user.COL7 -eq 'Nom2') {
        Write-Host "Skipping header row."
        continue
    }

    $prenom = Clean-String -String $user.COL8 -ReplaceSpacesWithUnderscore $false
    $nom = Clean-String -String $user.COL7 -ReplaceSpacesWithUnderscore $false
    $ville = Clean-String -String $user.COL14 -ReplaceSpacesWithUnderscore $true
    $Structure = '@' + $user.COL16
    $Lieu = $user.COL14 -replace " ", "_"
    $Responsable = $user.COL4
    $Fonction = $user.COL13
    $Service = $user.COL15 -replace " ", "_"
    $date = if ([string]::IsNullOrEmpty($user.COL11)) { "0" } else { $user.COL11 }
    $contrat = $user.COL10

    # Find OU and other details from lieux_csv
    $OU = $lieux_csv | Where-Object { $_.lieu -eq $ville } | Select-Object -ExpandProperty OU
    $tel = $lieux_csv | Where-Object { $_.lieu -eq $ville } | Select-Object -ExpandProperty tel
    $adresse = $lieux_csv | Where-Object { $_.lieu -eq $ville } | Select-Object -ExpandProperty Adresse
    $Codepostal = $lieux_csv | Where-Object { $_.lieu -eq $ville } | Select-Object -ExpandProperty Codepostal

    # Log each user creation attempt
    $logMessage = "Attempting to create user: $prenom $nom`n"
    $logMessage | Out-File -FilePath $logFilePath -Append

    # Check for essential values
    if ([string]::IsNullOrEmpty($prenom)) {
        $logMessage = "Error: First name is required for user creation. Skipping user.`n"
        $logMessage | Out-File -FilePath $logFilePath -Append
        continue
    }

    if ([string]::IsNullOrEmpty($nom)) {
        $logMessage = "Error: Last name is required for user creation. Skipping user.`n"
        $logMessage | Out-File -FilePath $logFilePath -Append
        continue
    }

    if ([string]::IsNullOrEmpty($OU)) {
        $logMessage = "Error: OU is required for user creation. Skipping user.`n"
        $logMessage | Out-File -FilePath $logFilePath -Append
        continue
    }

    $pass = Generate-RandomPassword
    $SamAccountName = "$($nom.Substring(0, [Math]::Min(15, $nom.Length)))$($prenom.Substring(0, 1))".ToLower()
    $displayName = "$prenom $nom"
    $mail = "$($prenom.Substring(0, 1)).$nom@$($Structure.TrimStart('@'))".ToLower()

    # Create the user in Active Directory
    try {
        New-ADUser -DisplayName $displayName -Name $displayName -UserPrincipalName $mail -GivenName $prenom -Surname $nom -AccountPassword (ConvertTo-SecureString $pass -AsPlainText -Force) -Enabled $true -Path "OU=User,OU=Account,OU=$OU,OU=ENTIS,DC=cetremut,DC=pri" -EmailAddress $mail -OfficePhone $tel -Description $Service -Office $Lieu -StreetAddress $adresse -PostalCode $Codepostal -City $ville -SamAccountName $SamAccountName -Manager $Responsable -Title $Fonction -Department $Service

        $logMessage = "Successfully created user: $displayName in OU: $OU`n"
        $logMessage | Out-File -FilePath $logFilePath -Append

        # Generate a message for the support ticket
        $ticketMessage = @"
Bonjour,

Voici les informations pour le nouvel utilisateur créé :

Nom complet : $displayName
Nom d'utilisateur (SamAccountName) : $SamAccountName
Adresse e-mail : $mail
Mot de passe : $pass

Cordialement,
"@

        Write-Host "Message pour le ticket de support :"
        Write-Host $ticketMessage
    }
    catch {
        $logMessage = "Error creating user $displayName: $($_.Exception.Message)`n"
        $logMessage | Out-File -FilePath $logFilePath -Append
    }
}