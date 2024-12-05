# Importeer de Active Directory-module
Import-Module ActiveDirectory

# CSV bestand locatie
$csvPath = "D:\Scripts\users.csv"

# Basis OU pad
$baseOuPath = "OU=DTA,DC=DataTransferAnalog,DC=com"
$domainPath = "DC=DataTransferAnalog,DC=com"

$domein = "DataTransferAnalog"

# Controleer of de hoofd OU 'DTA' bestaat, en maak deze indien nodig aan
if (-not (Get-ADOrganizationalUnit -Filter { Name -eq 'DTA' })) {
    New-ADOrganizationalUnit -Name 'DTA' -Path $domainPath
    Write-Host "Hoofd-OU 'DTA' aangemaakt."
} else {
    Write-Host "Hoofd-OU 'DTA' bestaat al."
}

# Importeer CSV bestand
$users = Import-Csv -Path $csvPath

# Loop door elke gebruiker in het CSV bestand
foreach ($user in $users) {
    # Parameters voor de nieuwe gebruiker
    $firstName = $user.Voornaam
    $lastName = $user.Achternaam
    $department = $user.Afdeling
    $title = $user.Functie
    $phoneNumber = $user.Telefoonnummer
    $address = $user.Adres
    $postalCode = $user.Postcode
    $city = $user.Plaatsnaam

    # Maak een gebruikersnaam (bijv. voornaam.achternaam)
    $username = "$firstName.$lastName"
    $email = "$username@$domein.com"

    # Bepaal de OU voor de gebruiker op basis van de afdeling
    $ouPath = "OU=$department,$baseOuPath"

    # Controleer of de OU voor de afdeling bestaat, maak anders aan
    if (-not (Get-ADOrganizationalUnit -Filter { DistinguishedName -eq $ouPath })) {
        New-ADOrganizationalUnit -Name $department -Path $baseOuPath
        Write-Host "OU '$department' aangemaakt in 'DTA'."
    } else {
        Write-Host "OU '$department' bestaat al."
    }

    # Controleer of de security group bestaat voor de afdeling, maak deze aan als deze niet bestaat
    $groupName = "$department"
    $group = Get-ADGroup -Filter { Name -eq $groupName } -ErrorAction SilentlyContinue

    if (-not $group) {
        # Maak de security group aan als deze niet bestaat
        New-ADGroup -Name $groupName `
                    -GroupScope Global `
                    -GroupCategory Security `
                    -Path $ouPath
        Write-Host "Security group '$groupName' aangemaakt in OU '$department'."
    } else {
        Write-Host "Security group '$groupName' bestaat al."
    }

    # Creëer de nieuwe gebruiker
    try {
        New-ADUser -Name "$firstName $lastName" `
                   -GivenName $firstName `
                   -Surname $lastName `
                   -UserPrincipalName "$username@$domein.local" `
                   -SamAccountName $username `
                   -EmailAddress $email `
                   -Department $department `
                   -Title $title `
                   -OfficePhone $phoneNumber `
                   -StreetAddress $address `
                   -PostalCode $postalCode `
                   -City $city `
                   -Path $ouPath `
                   -AccountPassword (ConvertTo-SecureString "Wachtwoord1" -AsPlainText -Force) `
                   -ChangePasswordAtLogon $true `
                   -Enabled $true

        Write-Host "Gebruiker $username succesvol aangemaakt."

        # Voeg de gebruiker toe aan de security group
        Add-ADGroupMember -Identity $groupName -Members $username
        Write-Host "Gebruiker $username toegevoegd aan groep $groupName."

    } catch {
        Write-Host "Fout bij het aanmaken van gebruiker $username of toevoegen aan groep. Fout: $($_.Exception.Message)"
        # Extra foutafhandeling of logging kan hier worden toegevoegd
    }
}