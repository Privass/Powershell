Import-Module ExchangeOnlineManagement
#Connect-ExchangeOnline

# Définition des variables du script
$logPath = "c:\temp\x500_generator.log"
Add-Content -Path $logPath -Encoding UTF8 -Value "UPN,mailNickname,LegacyExchangeDN,etat"
$e = 0 # Initialisation du compteur d'erreur
$m = 0 # Initialisation du compteur d'utilisateur modifié
$a = 0 # Initialisation du compteur de x500 deja existant
$c = 0 # Initialisation du compteur de BAL Cloud
$StartTime = (Get-Date).Millisecond
$LegacyExchangeDN = ""
$test = ""
$etat = ""

# Récupération des comptes utilisateurs actifs avec l'attribut sync et l'attribut LegacyExchangeDN défini.
$Users = Get-ADUser -LDAPFilter '(&(objectCategory=person)(extensionAttribute1=Sync)(objectClass=user)(!userAccountControl:1.2.840.113556.1.4.803:=2)(LegacyExchangeDN=*))' -Properties SamAccountName,UserPrincipalName,proxyAddresses,mailNickname | select SamAccountName,UserPrincipalName,proxyAddresses,mailNickname

$TotalItems=$Users.Count
$CurrentItem = 0
$PercentComplete = 0

#Pour chaque utilisateur
foreach ($User in $Users)
{
    # Barre de progression
    Write-Progress -Activity "En cours " -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete
    
    # Définition des variables de la boucle
    $i = 0
    $UPN = $User.UserPrincipalName
    $mailNickname = $User.mailNickname
    $SamAccountName = $User.SamAccountName
    
    # Nettoyage des variables de la boucle
    Clear-Variable LegacyExchangeDN
    Clear-Variable test
    Clear-Variable etat

    $test = Get-Mailbox -Identity $UPN -ErrorAction SilentlyContinue # Test si la BAL est cloud

    if ($test -eq $null) # Si la BAL n'est pas Cloud
    {
        $LegacyExchangeDN = Get-MailUser -Identity $UPN -ErrorAction SilentlyContinue | select LegacyExchangeDN  # Récupération de la valeur LegacyExchangeDN avec l'UPN
        if ($? -eq "false") { # Si la commande précédente ne fonctionne pas
            #Write-Host "Impossible de récupérer les informations du compte $UPN avec l'UPN"
            #Write-Host "Nouvel essai avec le mailNickname $mailNickname"
            $LegacyExchangeDN = Get-MailUser -Identity $mailNickname -ErrorAction SilentlyContinue | select LegacyExchangeDN # Récupération de la valeur LegacyExchangeDN avec le mailNickname
        }
    
        if ($LegacyExchangeDN -ne $null) # Si nous avons réussi à récuperer le LegacyExcangeDN dans ExchangeOnline
        {
            $LegacyExchangeDN = "x500:$($LegacyExchangeDN.LegacyExchangeDN)"

            foreach ($Value in $($User.proxyAddresses)) # Pour chaque valeur de l'attribut proxyAddresses
            {
                if ($Value -eq $LegacyExchangeDN) # Test des valeurs dans proxyAddresses avec la valeur LegacyExchangeDN dans le cloud
                {
                    $i++ # +1 si l'utilisateur posséde déja le x500
                }
            }

            if ($i -eq 0) # Si l'utilisateur ne posséde pas le x500
            {
                $m++ # Ajout de 1 au compteur des modifications
                #Set-ADUser $SamAccountName -add @{ProxyAddresses=$LegacyExchangeDN} -ErrorAction SilentlyContinue
                #if ($? -eq "false")
                #{
                #    Write-Host "Erreur lors de l'ajout du x500 pour $UPN" -BackgroundColor Red
                #    $etat = "Erreur1"
                #} else {
                #    $etat = "Ajout"
                #}
                #Write-Host " "
                #Write-Host "Ajout du x500 pour $UPN" -BackgroundColor Green
                #Write-Host " "
                #Write-Host "Le LegacyExchangeDN de $($User.UserPrincipalName) est $LegacyExchangeDN" -BackgroundColor Green
                #Write-Host " "
            } else { # Sinon si il a déja le x500
                $a++
                $etat = "Existant"
                #Write-Host "$UPN à déja un x500 valide."
            }        
        } else { # Sinon si nous n'avons pas réussi à récuperrer le LegacyExchangeDN das ExchangeOnline
            Write-Host " "
            Write-Host "La récupération des informations a échouée pour $UPN, Impossible de verifier le x500" -BackgroundColor Red
            Write-Host " "
            $e++ # Ajout de 1 compteur d'erreur
            $etat = "Erreur2"
        }
    } else { # Sinon la BAL est Cloud
        $c++
        $etat = "Cloud"
        #Write-Host " "
        #Write-Host "La BAL de $UPN est cloud, il n'y a rien à faire" -BackgroundColor Yellow
        #Write-Host " "
    }  

    $CurrentItem++
    $PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

    Add-Content -Path $logPath -Encoding UTF8 -Value "$UPN,$mailNickname,$LegacyExchangeDN,$etat"

}

$EndTime = (Get-Date).Millisecond

Write-Host "############################## FIN ############################"
Write-Host " "
Write-Host "$($Users.Count) Utilisateurs scannés / $m Utilisateurs modifiés / $c BAL Cloud / $a avaient déja un x500 / $e Erreurs"
Write-Host "Ce script a pris $($StartTime - $EndTime) Miliseconds à s'executer."
Write-Host " "
Write-Host "###############################################################"