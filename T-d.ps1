######################################################################
# 
######################################################################
		Start-Transcript -path "C:\Scripts\WMT\Log\2-TASK_Traitement_demandes.log" -Append -ErrorAction SilentlyContinue
		write-host "début"
		Import-Module ActiveDirectory
		$loginMP = ""

				Function Get-DomainCreds
		{
			
			[CmdletBinding()]
			Param (
				[Parameter(
						   Mandatory = $true,
						   ParameterSetName = 'Fresh'
						   )]
				[ValidateNotNullOrEmpty()]
				[string[]]$Domain,
				[Parameter(
						   Mandatory = $true,
						   ParameterSetName = 'File'
						   )]
				[Parameter(
						   Mandatory = $true,
						   ParameterSetName = 'Fresh'
						   )]
				[ValidateNotNullOrEmpty()]
				[string]$Path
			)
			
			If ($PSBoundParameters.ContainsKey('Domain'))
			{
				
				$Creds = @{ }
				ForEach ($DomainEach in $Domain)
				{
					$Creds[$DomainEach] = Get-Credential `
														 -Message "Enter credentials for domain $DomainEach" `
														 -UserName "$DomainEach\m.lazaroroot"
				}
				$Creds | Export-Clixml -Path $Path
				
			}
			Else
			{
				
				$Creds = Import-Clixml -Path $Path
				
			}
			
			Return $Creds
		}
Function Remove-StringSpecialCharacters
{
   Param([string]$String)

   $String -replace 'é', 'e' `
           -replace 'è', 'e' `
           -replace 'ç', 'c' `
           -replace 'ë', 'e' `
           -replace 'à', 'a' `
           -replace 'ö', 'o' `
           -replace 'ô', 'o' `
           -replace 'ü', 'u' `
           -replace 'ï', 'i' `
           -replace 'î', 'i' `
           -replace 'â', 'a' `
           -replace 'ê', 'e' `
           -replace 'û', 'u' `
           -replace '-', '' `
           -replace ' ', '' `
           -replace '/', '' `
           -replace '\*', '' `
           -replace "'", "" 
}
		function Remove-StringLatinCharacters
		{
			PARAM (
				[parameter(ValueFromPipeline = $true)]
				[string]$String
			)
			PROCESS
			{
				[Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
			}
		}
		function ConvertLDAPDateToDateTime {
			param (
				[string]$ldapDate
			)
			if ($ldapDate -match "(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2}).0Z") {
				$year = $matches[1]
				$month = $matches[2]
				$day = $matches[3]
				$hour = $matches[4]
				$minute = $matches[5]
				$second = $matches[6]
				return (Get-Date -Year $year -Month $month -Day $day -Hour $hour -Minute $minute -Second $second)
			}
			return $null
		}
		$i = @{ }
		$a_multipass = @{ }
		
		$Path = '\\partages\ECHANGES\INFRAS_MUTUALISEES\WMT\Data\Multipass\Export_Tout_Multipass.csv'
		
		#(Get-Content -Path $Path) | Set-Content -Path $Path -Encoding UTF8
		#Start-Sleep 4
		[hashtable]$global:hashMU = @{ }
		#$a_multipass = import-csv -Path \\partages\ECHANGES\INFRAS_MUTUALISEES\WMT\Data\Multipass\Export_Tout_Multipass.csv -Encoding UTF8 -Delimiter ";"

		#Start-Sleep 5
	#	foreach ($i in $a_multipass)
	#	{
			
		#	$cle_nom = Remove-StringSpecialCharacters $i.Nom 
#$cle_prenom = Remove-StringSpecialCharacters $i."Prénom"
#$cle = $cle_nom + " " + $cle_prenom
#			$global:hashMU[$cle] = $i
			
			
	#	}
		
		$login_de_luser = [Environment]::UserName
		$global:DomCreds = Get-DomainCreds -Path "C:\Users\$login_de_luser\WMT\creds.xml"

		
		#Domaine
		$serveurMBX = "vm-ad-mbx1.mbx.fr"
		$serveurBordeaux_IT = "bordeaux-it.fr"
		$serveurBordeauxIT = "vm-p-0050.bordeaux-it.fr"
		[string]$serveurMBX = "vm-ad-mbx1.mbx.fr"
		[string]$serveurBordeaux_IT = "bordeaux-it.fr"
		$Server_BdxIT = "bordeaux-it.fr" #
		$Server_MBX = "mbx.fr" #
		$ServerCUB = "vm-cub12.cub.local"
		$Server_CUB = "cub.local" #
		$date_demande = ""
		$HOLD = "-"
		#$multipass_name = Remove-StringLatinCharacters $global:hashMU.Keys
		#Partage
		$dir = "C:\Scripts\WMT\DATA\Demandes-UTF8.csv"
		$newpath_variable = "C:\Scripts\WMT\DATA\Demandes_Variables-UTF8.csv"
		
		$results = @()
		$importData = Import-Csv "$dir" -Delimiter "," -Encoding UTF8
		
		#On récup les infos du CSV Demandes_Variables-UTF8
		foreach ($row in $importData)
		{
			$csv_variable = Import-Csv $newpath_variable -Delimiter "," -Encoding UTF8
			$M_user_variable = $csv_variable | ? { $_.number -like $row.number }
			if ($M_user_variable)
			{
				$acces_bal = $M_user_variable.acces_bal
				$date_souhaite = $M_user_variable.date_souhaite
				$nom_compte = $M_user_variable.nom_compte
				$prenom_compte = $M_user_variable.prenom_compte
				$meme_droit_que = $M_user_variable.meme_droit_que
				$type_compte = $M_user_variable.type_compte
				$type_demande_messagerie = $M_user_variable.type_demande_messagerie
				$numero_tel_de_confiance = $M_user_variable.numero_tel_de_confiance
				$x = $row.company
                $date_naissance = $M_user_variable.date_naissance
                $date_depart = $M_user_variable.date_depart
                $manager_name = $M_user_variable.manager_name
                $manager_lastname = $M_user_variable.manager_lastname
                $nom_entreprise = $M_user_variable.nom_entreprise



	<#								
    Switch ($x) {
         "ville-talence.fr" {$TRI_commune = "TAL" ;$likecomm = "*Talence*" }
         "ville-talence.fr" {$TRI_commune = "TAL" ;$likecomm = "*Talence*" }
         default {$TRI_commune = "POK" ;$likecomm = "*PUK*"}
    }
    #>$TRI_commune = ""
				
				if ($x -like "*Talence*")
				{
					$TRI_commune = "TAL"
					$likecomm = "*Talence*"
				}
				if ($x -like "Pessac")
				{
					$TRI_commune = "PES"
					$likecomm = "*Pessac*"
				}
				if ($x -like "*Le Haillan*")
				{
					$TRI_commune = "LEH"
					$likecomm = "*Haillan*"
				}
				if ($x -like "*Floirac*")
				{
					$TRI_commune = "FLO"
					$likecomm = "*Floirac*"
				}
				if ($x -like "*Bordeaux*")
				{
					$TRI_commune = "BOR"
					$likecomm = "Bordeaux"
				}
                if ($x -like "CCAS*Bordeaux*")
				{
					$TRI_commune = "BOR"
					$likecomm = "*Bordeaux*"
				}
				if ($x -like "*Bordeaux*M*")
				{
					$TRI_commune = "MET"
					$likecomm = "*Metropole*"
				}
				if ($x -like "*B*gles*")
				{
					$TRI_commune = "BEG"
					$likecomm = "*Begles*"
				}
				if ($x -like "*M*rignac*")
				{
					$TRI_commune = "MER"
					$likecomm = "*Merignac*"
				}
				if ($x -like "*Bouscat*")
				{
					$TRI_commune = "LEB"
					$likecomm = "*Bouscat*"
				}
				if ($x -like "*Taillan*")
				{
					$TRI_commune = "LET"
					$likecomm = "*Taillan*"
				}
				if ($x -like "*Carbon*Blanc*")
				{
					$TRI_commune = "CAR"
					$likecomm = "*Carbon*Blanc*"
				}
				if ($x -like "*Bruges*")
				{
					$TRI_commune = "BRU"
					$likecomm = "*Bruges*"
				}
                if ($x -like "*Lagrave*")
				{
					$TRI_commune = "AEL"
					$likecomm = "*Lagrave*"
				}
                   if ($x -like "*blanque*")
				{
					$TRI_commune = "BLA"
					$likecomm = "*blanque*"
				}
                   if ($x -like "*aubin*")
				{
					$TRI_commune = "SAM"
					$likecomm = "*aubin*"
				}
				#Vérification si le compte est en retard par rapport a la date d'arrivée
				$u = $date_souhaite
				$u = [datetime]::ParseExact($u, "yyyy-MM-dd HH:mm:ss", $null)
				$u = $u.ToString("dd/MM")
				$Date2 = $u
				
				#On génére la description courte
				#$descrip_courte = $Date2 + " + " + $nom_compte + " " + $prenom_compte + " +"
				if ($nom_compte -eq "-")
				{
					$descrip_courte = "-"
				}
				else
				{
					$descrip_courte = $Date2 + " + " + $nom_compte + " " + $prenom_compte + " +"
				}
				
			}
			
			$New_user = $row.short_description
			if ($New_user -like "Tache de*")
			{
				$New_user = $descrip_courte
				
			}
			if ($New_user -like "*+*+*")
			{
				
				$New_user_delai = $row.short_description
				
				$j = get-date
				$date_souhaite = [Datetime]::ParseExact($date_souhaite, 'yyyy-MM-dd HH:mm:ss', $null)
				$date_depas = $date_souhaite -lt $j
				
				#$sanscaract
				#$date_souhaite
				if ($date_depas -eq "True")
				{
					$date_demande = "-Retard"
				}
				else
				{
					$date_demande = ""
				}
				#$date_depas
				$New_user = (($New_user.Split("+"))[1]).Trim()
				if (($New_user.Split(" ")).count -gt 2)
				{
					$New_user2 = (($New_user.Split(" "))[0]).Trim() + " " + (($New_user.Split(" "))[1]).Trim()
					$New_user_MULTIPASS = (($New_user.Split(" "))[0]).Trim() + "*" + (($New_user.Split(" "))[1]).Trim()
				}
				else
				{
					$New_user2 = (($New_user.Split(" "))[0]).Trim()
				}
			}
			else
			{
				$New_user = "-"
				$New_user2 = "-"
			}
			$row.request_item = "-"
			$CompteAD = ""
			$user_bdx_it = ""
			$accountBIT = ""
			$test_account = ""
			$test_bal_o365 = ""
			
			#On récup uniquement les demandes de création de comtpes
			if ($New_user -ne "-")
			{
				#On traite les villes ATC
				If ($row.company -like "**" -or $row.company -like "Bordeaux*")
				{
					# if attrib4 n'est pas egal compagny'
					
					try
					{
						
						$New_user_without_espace = $New_user.Replace(' ', '*') + "*" + $TRI_commune + "*"
						$accountBIT = get-aduser -f { (Name -like $New_user_without_espace) -And (Name -notlike "*root*") } -server $serveurBordeauxIT -Properties DistinguishedName, Name, SamAccountName, mail, extensionAttribute4, extensionAttribute3, extensionAttribute6, whenCreated | Select-Object Name, SamAccountName, mail, extensionAttribute4, extensionAttribute3, extensionAttribute6, whenCreated, DistinguishedName
						$user_bdx_it = $accountBIT.extensionAttribute4
						$login_user = $accountBIT.SamAccountName
						[datetime]$Date = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						[datetime]$Date = [datetime]$Date.Addhours(-1)
					}
					catch { }
					if (!$accountBIT)
					{
						$New_user_without_espaceEXT = $New_user.Replace(' ', '*') + "*EXT" + "*"
						$likecomm2 = "*" + $likecomm + "*"
						#$accountBIT = get-aduser -f { (Name -like $New_user_without_espaceEXT) -And (Name -notlike "*root*") } -server $serveurBordeauxIT -Properties DistinguishedName, Name, SamAccountName, mail, extensionAttribute4, extensionAttribute3, extensionAttribute6, whenCreated | Select-Object Name, SamAccountName, mail, extensionAttribute4, extensionAttribute3, extensionAttribute6, whenCreated, DistinguishedName
						$accountBIT = get-aduser -f { (Name -like $New_user_without_espaceEXT) -And (Name -notlike "*root*") -And (UserPrincipalName -like $likecomm2) } -server $serveurBordeauxIT -Properties DistinguishedName, Name, SamAccountName, mail, extensionAttribute4, extensionAttribute3, extensionAttribute6, whenCreated, UserPrincipalName | Select-Object Name, SamAccountName, mail, extensionAttribute4, extensionAttribute3, extensionAttribute6, whenCreated, DistinguishedName, UserPrincipalName

						$user_bdx_it = $accountBIT.extensionAttribute4
						$login_user = $accountBIT.SamAccountName
						[datetime]$Date = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						[datetime]$Date = [datetime]$Date.Addhours(-1)
						
					}
					
					#Vérification que le compte est dans l'AD
					if ($accountBIT)
					{
						#write-host $accountBIT
						$row.request_item = "2-A synchro -> INFRACENTRALE"
						$CompteAD += $user_bdx_it
						#OU HOLD 
						$accountBIT.DistinguishedName
						if ($accountBIT.DistinguishedName -like "*OU=HOLD*")
						{
							$HOLD = "OK"
							
						}
											
						if ($accountBIT.extensionAttribute6 -ne $null -and $accountBIT.extensionAttribute3 -ne $null -or $CompteAD -notlike "*@*@*")
						{
							
							if ($accountBIT.extensionAttribute4 -ne $null)
							{
								$row.request_item = "3-En attente synchro M365"
								$PathArray = @()
								
								if ($accountBIT.mail -ne $null)
								{
									$row.request_item = "5-Compte OK - Terminée"
									if($HOLD -like "OK"){
										$row.request_item = "5-Compte OK - HOLD"
									}									
								}
								else
								{
									[datetime]$Date_jour = Get-Date -Format "MM/dd/yyyy"
									$date_du_compte = $accountBIT.whenCreated
									
									#If avec l'heure
									[datetime]$Date = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									[datetime]$Date = [datetime]$Date.Addhours(-1)
									
									
									if ([datetime]$Date_jour -gt $accountBIT.whenCreated)
									{
										$row.request_item = "4-Synchro OK -> BAL à créer"
									}
									else
									{
										
										if ([datetime]$Date.Addminutes(-30) -gt $accountBIT.whenCreated)
										{
											$row.request_item = "4-Synchro OK -> BAL à créer"
										}
										else
										{
											$row.request_item = "4-Synchro KO -> En attente"
										}
									}
									
									#Vérification si l'user a une BAL en cours de création dans sharepoint
									#$newpath_PnP = "C:\Scripts\WMT\DATA\PNP-UTF8.csv"
									#$csv_Pnp = Import-Csv $newpath_PnP -Delimiter ";" -Encoding UTF8
									#$M_user_pnp = $csv_Pnp | ? { $_.Utilisateur -like $user_bdx_it }
									if ($M_user_pnp)
									{
										$row.request_item = "BAL en cours : " + $M_user_pnp.statut
									}
									
								}
								if ($acces_bal -eq "No")
								{
									$row.request_item = "5-Compte OK- Compte sans BAL"
								}
								#$row.request_item = "En attente bonne entitée."				
							}
						} #fin
						else { }
					}
					else
					{
						$row.request_item = "0-Multipass"
					}
				}
				
				
			}
			else
			{
				$row.request_item = "Description à changer dans ASAP."
			}
			if ($row.short_description -like "*#*")
			{
				$row.request_item = "Autre"
			}
			if ($row.short_description -like "*bal de service*" -or $row.short_description -like "Délég de service*" -or $row.short_description -like "Délég user*")
			{
				$row.request_item = "A traiter"
			}
			if ($row.short_description -like "*Citrix*")
			{
				$row.request_item = "A traiter Citrix"
			}
			if ($row.short_description -like "*Demande d'acc*s r*seau*" -or $row.short_description -like "*serveur fichiers*" -or $row.short_description -like "*droits d'accès*"  -or $row.short_description -like "*Modification droit*")
			{
				$row.request_item = "A traiter"
			}
			if ($row.short_description -like "*Demande de droits d'accès réseau*")
			{
				$row.request_item = "Droits réseau"
			}
			if ($row.short_description -like "*@*.fr*")
			{
				$row.request_item = "A traiter - BAL de service"
				$newpath_PnP_shared = "C:\Scripts\WMT\DATA\PNP-SHAREDUTF8.csv"
				#$csv_Pnp_shared = Import-Csv $newpath_PnP_shared -Delimiter ";" -Encoding UTF8

				$test_shared = $row.short_description
				$test_shared2 = $test_shared.split("@")[0]

				$M_SHARED_pnp = $csv_Pnp_shared | ? { $_.data -like "*info.habitat.larousselle*" }
				if ($M_SHARED_pnp)
				{
					$row.request_item = $M_SHARED_pnp.comment
				}
				
				$M_SHARED_pnp2 = $csv_Pnp_shared | ? { $_.PrefixSMTP -like $row.short_description }
				if ($M_SHARED_pnp2)
				{
					$row.request_item = "BAL en cours : " + $M_SHARED_pnp2.statut
				}
			}
			
			#$New_user = $New_user -is [array]
			
			if ($CompteAD -like "*@*@*")
			{
				$row.request_item = "Homonyme, ajouter le matricule après le prenom dans ASAP."
			}
			
			if ($row.short_description -like "*A traiter*")
			{
				$row.request_item = "A traiter"
			}
			
			
			
			if ($row.request_item -like "0-Multipass*")
			{
				if ($row.short_description -like "*stgnr*" -or $row.short_description -like "*presta*" -or $row.short_description -like "*stg*non*" -or $type_compte -like "*stag*non*" -or $type_compte -like "*resta*")
				{
					$row.request_item = "A créer dans Multipass."
				}
				#test name uniquement : 
				#$MotifMulti_name = "*" + $New_user_MULTIPASS + "*"
				#Remove-StringLatinCharacters $MotifMulti_name
				#$voilamulti_name = $global:hashMU.Keys -like $MotifMulti_name
				
				
				#Fin test
				$New_user_without_espace2 = $New_user.Replace(' ', '*')
				$MotifMulti = "*" + $New_user_without_espace2 + "*"
				$sanscaract = Remove-StringLatinCharacters $MotifMulti
				$voilamulti = $global:hashMU.Keys -like $sanscaract.ToLower()


# Remplacer les valeurs statiques par des variables
$displayNameArray = "*" + $prenom_compte.Replace('-', '*').Replace(' ', '*') + "*" + $nom_compte.Replace('-', '*').Replace(' ', '*') + "*"

write-host "recherche de $displayNameArray "
$BMEntiteJuridiqueLibelle = $likecomm


$response = @()
$headers = @{
    "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($loginMP)))
}
# Créer le filtre en utilisant les variables
$fields = "manager_display,BMEntreprise,BMADLogin,BMADExist,BMDateDepartRH,BMGradeLibelle,FZM0301,BMOrganisationLibelle,enatelBeginTime,BMStatutDN,BMPrenomUsuel,givenName,ATCLogin,sn,BMMatriculeOrigine,evdpmSchedulerStatus,evdmaildomain,ATCRootLogin,BMServiceCode,BMDirectionLib,BMStatutLibelle,BMAD,BMEntiteJuridiqueLibelle,BMATCExist,employeeNumberTRI,mail,enatelEndTime,BMUPN,BMDateNaissance,BMPrenoms,whenCreated,whenChanged,displayName,title,employeeNumber"
$filter = "(&(displayName=$displayNameArray)(BMEntiteJuridiqueLibelle=$BMEntiteJuridiqueLibelle))"
$encodedFilter = [System.Web.HttpUtility]::UrlEncode($filter)

# Reste du code pour appeler l'API et faire les conversions
$url = "https://vm-p-0265.bordeaux-it.fr:4245/Portal/api/v1/identities?fields=$fields&filter=$encodedFilter&withEmptyValues=false&searchUsersOption=searchUsers&page=1&pageSize=10"
#$url = "https://vm-p-0265.bordeaux-it.fr:4245/Portal/api/v1/identities?fields=$fields&filter=$encodedFilter&withEmptyValues=false&searchUsersOption=searchUsersInRetention&page=1&pageSize=10"
try {
    $response = Invoke-RestMethod -Uri $url -Method Get -ContentType "application/json" -Headers $headers
	$response | ForEach-Object {
		if ($_.attributes.PSObject.Properties.Name -contains "BMDateNaissance") {
			$dateObj = (ConvertLDAPDateToDateTime $_.attributes.BMDateNaissance).AddDays(1)
			$_.attributes.BMDateNaissance = $dateObj.ToString("dd/MM/yyyy")
		}
		
		if ($_.attributes.PSObject.Properties.Name -contains "enatelEndTime") {
			$dateObj = (Get-Date "1970-01-01 00:00:00").AddSeconds([double]::Parse($_.attributes.enatelEndTime)).AddDays(1)
			$_.attributes.enatelEndTime = $dateObj.ToString("dd/MM/yyyy")
		}
	
		if ($_.attributes.PSObject.Properties.Name -contains "enatelBeginTime") {
			$dateObj = (Get-Date "1970-01-01 00:00:00").AddSeconds([double]::Parse($_.attributes.enatelBeginTime)).AddDays(1)
			$_.attributes.enatelBeginTime = $dateObj.ToString("dd/MM/yyyy")
		}
	}
}
catch {
    #Write-Host "Réponse : $_.Exception.Message"
}

				if ($response)
				{
					#$global:hashMU[$cle].nom -like $MotifMulti.ToLower()
					#write-host "Nom ASAP : "$New_user_without_espace2 " Nom Multipass : "$voilamulti

					#$row.request_item = "En attente bonne entité"
					$row.request_item = "Présent dans MULTIPASS"
					#if ($global:hashMU[$voilamulti].ej -like $likecomm)
					#{
					#	$row.request_item = "Présent dans MULTIPASS"
					#}
				}
				else
				{
					$row.request_item = "0-Multipass-Non trouvé"
					$url = "https://vm-p-0265.bordeaux-it.fr:4245/Portal/api/v1/identities?fields=$fields&filter=$encodedFilter&withEmptyValues=false&searchUsersOption=searchUsersInRetention&page=1&pageSize=10"
					try {
						$response = Invoke-RestMethod -Uri $url -Method Get -ContentType "application/json" -Headers $headers
						$response | ForEach-Object {
							if ($_.attributes.PSObject.Properties.Name -contains "BMDateNaissance") {
								$dateObj = (ConvertLDAPDateToDateTime $_.attributes.BMDateNaissance).AddDays(1)
								$_.attributes.BMDateNaissance = $dateObj.ToString("dd/MM/yyyy")
							}
							
							if ($_.attributes.PSObject.Properties.Name -contains "enatelEndTime") {
								$dateObj = (Get-Date "1970-01-01 00:00:00").AddSeconds([double]::Parse($_.attributes.enatelEndTime)).AddDays(1)
								$_.attributes.enatelEndTime = $dateObj.ToString("dd/MM/yyyy")
							}
						
							if ($_.attributes.PSObject.Properties.Name -contains "enatelBeginTime") {
								$dateObj = (Get-Date "1970-01-01 00:00:00").AddSeconds([double]::Parse($_.attributes.enatelBeginTime)).AddDays(1)
								$_.attributes.enatelBeginTime = $dateObj.ToString("dd/MM/yyyy")
							}
						}
					}
					catch {
						
					}
					if ($response)
					{
						$row.request_item = "0-Multipass-En retention"
					}
				}
				

			}
			if ($row.short_description -like "*- ? -*")
			{
				$row.request_item = "Nom a changer"
			}
			if ($row.short_description_type -eq "Formulaire de demande de création de compte informatique")
			{
				$row.short_description_type = "Compte"
			}
			
			if ($row.short_description_type -eq "Création, modification, délégation ou suppression de boites aux lettres")
			{
				$row.short_description_type = "BAL"
				#$descrip_courte = "Délégation de BAL"
			}
			if ($row.short_description_type -eq "Installation, attribution de droits ou suppression d'une application métier")
			{
				$row.short_description_type = "Appli"
				#$descrip_courte = "Délégation de BAL"
			}
			if ($row.short_description_type -eq "Demander des droits d'accès aux partages réseaux")
			{
				$row.short_description_type = "Partages"
				#$descrip_courte = "Délégation de BAL"
			}
			
			if ($type_demande_messagerie -eq "creation_boite_aux_lettres_nominative")
			{
				$row.short_description_type = "Créer BAL"
				#$descrip_courte = "Délégation de BAL"
			}
			if ($type_demande_messagerie -eq "creation_boite_aux_lettres_service")
			{
				$row.short_description_type = "BAL service"
				#$descrip_courte = "Délégation de BAL"
			}
			if ($type_demande_messagerie -eq "delegation_boite_aux_lettres_nominative")
			{
				$row.short_description_type = "Délég Nominative"
				#$descrip_courte = "Délégation de BAL"
			}
			if ($type_demande_messagerie -eq "delegation_boite_aux_lettres_service")
			{
				$row.short_description_type = "Délég de service"
				#$descrip_courte = "Délégation de BAL"
			}
			if ($type_demande_messagerie -eq "modif_boite_aux_lettres_nominative")
			{
				$row.short_description_type = "Modif BAL Nominative"
				#$descrip_courte = "Délégation de BAL"
			}
			if ($type_demande_messagerie -eq "modif_boite_aux_lettres_service")
			{
				$row.short_description_type = "Modif BAL de service"
				#$descrip_courte = "Délégation de BAL"
			}
			if ($type_demande_messagerie -eq "resiliation_boite_aux_lettres_nominative")
			{
				$row.short_description_type = "Résiliation BAL Nominative"
				#$descrip_courte = "Délégation de BAL"
			}
			if ($type_demande_messagerie -eq "resiliation_boite_aux_lettres_service")
			{
				$row.short_description_type = "Résiliation BAL de service"
				#$descrip_courte = "Délégation de BAL"
			}
			if ($type_demande_messagerie -like "Pour toute autre demande*")
			{
				$row.short_description_type = "Autre"
				#$descrip_courte = "Délégation de BAL"
			}
			if ($type_demande_messagerie -eq "Pour toute autre demande ne figurant pas dans le catalogue des demandes de services")
			{
				$row.short_description_type = "Autre"
				#$descrip_courte = "Délégation de BAL"
			}
			#$type_demande_messagerie = $M_user_variable.type_demande_messagerie
			
			$requestitem = $row.request_item + $date_demande
			$details = @{
				number = $row.number
				sys_created_on = $row.sys_created_on
				state  = $row.state
				assigned_to = $row.assigned_to
				short_description = $row.short_description
				company = $row.company
				request_item = $requestitem
				New_user = $New_user
				Name   = $New_user2
				CompteAD = $CompteAD
				sys_updated_by = $row.sys_updated_by
				short_description_type = $row.short_description_type
				u_nom_prenom_benificiaire = $row.u_nom_prenom_benificiaire
				acces_bal = $acces_bal
				date_souhaite = $date_souhaite
				nom_compte = $nom_compte
				prenom_compte = $prenom_compte
				meme_droit_que = $meme_droit_que
				type_compte = $type_compte
				desc_crt = $descrip_courte
				numero = $numero_tel_de_confiance
				date_demande = $date_demande
                sys_id = $row.sys_id_task
                date_naissance = $date_naissance
                date_depart = $date_depart
                manager_name = $M_user_variable.manager_name
                manager_lastname = $M_user_variable.manager_lastname
                nom_entreprise = $nom_entreprise
                login = $login_user
                compte_a_prolonger_name = $M_user_variable.compte_a_prolonger_name
                compte_a_prolonger_lastname = $M_user_variable.compte_a_prolonger_lastname
				responsable_compte_a_prolonger = $M_user_variable.responsable_compte_a_prolonger_name
                date_fin_compte = $M_user_variable.date_fin_compte
				acces_partage_ajouter = $M_user_variable.acces_partage_ajouter
				informations_complementaires = $M_user_variable.informations_complementaires
				OU_HOLD = $HOLD
				location_u_multipass_id = $row.location_u_multipass_id
				location_u_site = $row.location_u_site
				location_full_name = $row.location_full_name
				location_street = $row.location_street
				location_u_building = $row.location_u_building
				location_u_floor = $row.location_u_floor
				location_u_room = $row.location_u_room
				multipass_ID = $row.multipass_ID
				Multipass_DN = $row.Multipass_DN
				u_direction_generale_code = $row.u_direction_generale_code
				u_direction_generale_description = $row.u_direction_generale_description
				u_direction_generale_id = $row.u_direction_generale_id
				u_direction_generale_name = $row.u_direction_generale_name

				
			}
			
			$results += New-Object PSObject -Property $details
		}
		$results | export-csv -Path C:\Scripts\WMT\DATA\demande_traite.csv -NoTypeInformation -Encoding UTF8 -Delimiter ","
		Start-Sleep 3

		write-host "fin"
Stop-Transcript

