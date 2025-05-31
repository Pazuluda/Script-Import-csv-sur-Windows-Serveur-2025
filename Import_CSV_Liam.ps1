
Clear-Host

# CHARGEMENT DU MODULE ActiveDirectory
if (-not (Get-Module -Name ActiveDirectory)) {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        Write-Host "ERREUR : MODULE ActiveDirectory NON DISPONIBLE" -ForegroundColor Red
        exit
    }
}

$domainDN = (Get-ADDomain).DistinguishedName
$domainName = ($domainDN -replace 'DC=', '') -replace ',', '.'

# CREATION DE L'UO PRINCIPALE
$uoPrincipale = "SIEGE"
$uoPrincipaleDN = "OU=$uoPrincipale,$domainDN"
if (-not (Get-ADOrganizationalUnit -LDAPFilter "(distinguishedName=$uoPrincipaleDN)" -ErrorAction SilentlyContinue)) {
    New-ADOrganizationalUnit -Name $uoPrincipale -Path $domainDN -ProtectedFromAccidentalDeletion $true -Verbose
}

# CREATION DES UO SECONDAIRES
$uoSecondaires = @("COMPTA", "IT", "DIRECTION", "DL")
foreach ($uo in $uoSecondaires) {
    $path = "OU=$uo,OU=$uoPrincipale,$domainDN"
    if (-not (Get-ADOrganizationalUnit -LDAPFilter "(distinguishedName=$path)" -ErrorAction SilentlyContinue)) {
        New-ADOrganizationalUnit -Name $uo -Path $uoPrincipaleDN -ProtectedFromAccidentalDeletion $true -Verbose
    }
}

# CHEMIN DU CSV
$csvPath = Read-Host "ENTREZ LE CHEMIN COMPLET DU FICHIER CSV"
if (-not (Test-Path $csvPath)) {
    Write-Host "FICHIER INTROUVABLE : $csvPath" -ForegroundColor Red
    exit
}

$csvUsers = Import-Csv $csvPath

foreach ($ligne in $csvUsers) {
    $prenom = $ligne.Prenom.Trim()
    $nom = $ligne.Nom.Trim()
    $display = $ligne.DisplayName.Trim()
    $sam = $ligne.SamAccountName.Trim().ToLower()
    $upn = $ligne.UserPrincipalName.Trim().ToLower()
    $uo = $ligne.UO.Trim().ToUpper()
    $groupesUtilisateurs = $ligne.GroupesUtilisateurs
    $relationsGGversDL = $ligne.GGVersDL

    $userOU = "OU=$uo,OU=$uoPrincipale,$domainDN"

    if (-not (Get-ADOrganizationalUnit -LDAPFilter "(distinguishedName=$userOU)" -ErrorAction SilentlyContinue)) {
        New-ADOrganizationalUnit -Name $uo -Path $uoPrincipaleDN -ProtectedFromAccidentalDeletion $true -Verbose
    }

    $userParams = @{
        Name               = $display
        GivenName          = $prenom
        Surname            = $nom
        SamAccountName     = $sam
        UserPrincipalName  = $upn
        Path               = $userOU
        AccountPassword    = (ConvertTo-SecureString "Ertyuiop," -AsPlainText -Force)
        Enabled            = $true
    }

    New-ADUser @userParams -Verbose

    if ($groupesUtilisateurs) {
        $groupes = $groupesUtilisateurs -split '[;,]'
        foreach ($gg in $groupes) {
            $gg = $gg.Trim()
            $ggPath = $userOU

            if (-not (Get-ADGroup -Filter "Name -eq '$gg'" -SearchBase $ggPath -ErrorAction SilentlyContinue)) {
                New-ADGroup -Name $gg -GroupScope Global -GroupCategory Security -Path $ggPath -Verbose
            }

            try {
                Add-ADGroupMember -Identity $gg -Members $sam -Verbose
            } catch {
                Write-Host "⚠️ ERREUR AJOUT $sam DANS $gg" -ForegroundColor Red
            }
        }
    }

    if ($relationsGGversDL) {
        $liaisons = $relationsGGversDL -split '[;,]'
        foreach ($liaison in $liaisons) {
            if ($liaison -match '^(.*?)=>(.*?)$') {
                $gg = $matches[1].Trim()
                $dl = $matches[2].Trim()
                $dlPath = "OU=DL,OU=$uoPrincipale,$domainDN"

                if (-not (Get-ADGroup -Filter "Name -eq '$dl'" -SearchBase $dlPath -ErrorAction SilentlyContinue)) {
                    New-ADGroup -Name $dl -GroupScope DomainLocal -GroupCategory Security -Path $dlPath -Verbose
                }

                try {
                    Add-ADGroupMember -Identity $dl -Members $gg -Verbose
                } catch {
                    Write-Host "⚠️ ERREUR AJOUT $gg DANS $dl" -ForegroundColor Red
                }
            }
        }
    }
}

Write-Host "`n✅ IMPORT TERMINE AVEC SUCCES !" -ForegroundColor Green
