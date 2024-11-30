#Deux parametres, le premier: le csv des serveurs // le deuxieme: le csv des kb a install
# ./checkinstallvb1.ps1 -csvPath nomdulot.csv -csvPath2 kbainstall.csv
param (
    [string]$csvPath,
    [string]$csvPath2
)

# Vérifier si le chemin du fichier CSV est fourni
if (-not $csvPath -or -not $csvPath2) {
    Write-Host "Veuillez fournir le chemin du fichier CSV en argument: ./checkinstallvb1.ps1 -csvPath nomdulot.csv -csvPath kbainstall.csv "
    exit
}

# Vérifier si les fichiers existent
if (-not (Test-Path -Path $csvPath)) {
    Write-Host "Le fichier CSV des serveurs spécifié n'existe pas : $csvPath"
    exit
}

if (-not (Test-Path -Path $csvPath2)) {
    Write-Host "Le fichier CSV des KB spécifié n'existe pas : $csvPath2"
    exit
}

# Extraire le numero de lot 
$NumLot = [System.IO.Path]::GetFileNameWithoutExtension($csvPath)

# Lire le fichier CSV
$servers = Import-Csv -Path $csvPath
$KBList = Import-Csv -Path $csvPath2

#Definir le repertoire d'envoi du rapport
$kbdestination = "C:\Temp\RapportKB"

#On recupere la date et l'heure
$date= Get-Date -Format "dd-MM-yyyy_HH-mm"
#Generer le path du fichier CSV de rapport
$reportcsvpath = Join-Path -Path $kbdestination -ChildPath "$NumLot-rapportKB_$date.csv"

#Creation d'une liste pour stocker les infos pour le report
$reportdata = @()

# Demander les informations d'identification une seule fois
$credential = Get-Credential

# Parcourir chaque serveur et exécuter la commande Get-Hotfix
Write-Host -ForegroundColor White "`n`nComparaison de la liste des KB a install avec les 5 derniers hotfixes de chaque serveur:"
foreach ($server in $servers) {
    $serverName = $server.ComputerName
    # On recupere la version de l'os (j'utilise invoke command car le get-ciminstance en remote directement ne marche pas)
    $osversion = Invoke-Command -ComputerName $serverName -Credential $credential -ScriptBlock {
        Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object Caption
    }
    #J'utilie le notmatch avec 2008|2012 car cela fonctionnait mieux que notlike
    if ($osversion.Caption -notmatch "Windows Server (2008|2012)"){
        Write-Host "Version du serveur: " $osversion.Caption
        # On recupere la date du dernier boot
        $lastBootTime = Invoke-Command -ComputerName $serverName -Credential $credential -ScriptBlock {
            Get-WmiObject -Class Win32_OperatingSystem | Select-Object -ExpandProperty LastBootUpTime
        }
        $lastBootTimeFormatted = [System.Management.ManagementDateTimeConverter]::ToDateTime($lastBootTime).ToString("dd/MM/yyyy HH:mm")
        # On regarde si un reboot est required
        $rebootRequired = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue
        if ($rebootRequired) {
            $rebootRequired = $true
        } else {
            $rebootRequired = $false
        }
        [bool]$ServerKBinstall = $false
        $arrayKB = @()
        try {
            Write-Host -ForegroundColor cyan "`nExecution de Get-Hotfix sur le serveur : $serverName"
            Write-Host "`nDate du dernier reboot:" $lastBootTimeFormatted
            try {
                # On recupere tous les correctifs
                $hotfixes = Invoke-Command -ComputerName $serverName -Credential $credential -ScriptBlock {
                    Get-Hotfix
                }
            
                # On filtre car parfois le kb est present mais pas installé du coup pas de date qui remonte dans installedon et cela genere une erreur non terminante
                $validHotfixes = $hotfixes | Where-Object {
                    # On check si il y a bien le installedon, j'utilise le psobjectproperties car le parse ne fonctionnait pas bien et les serveurs ou les kb etaient ok ne remontait pas avec le parse
                    $_.PSObject.Properties['InstalledOn'] -and $_.InstalledOn -ne $null -and $_.InstalledOn -isnot [string]
                }
                # Si aucun correctif avec une date valide n'est trouvé, afficher un message d'information
                if ($validHotfixes.Count -eq 0) {
                    Write-Host "Aucun correctif avec date valide n'a ete trouve. Aucun tri effectue."
                    $tophotfixes = $hotfixes | Select-Object -First 5
                } else {
                    # Trier uniquement les correctifs avec une date valide
                    $tophotfixes = $validHotfixes | Sort-Object InstalledOn -Descending | Select-Object -First 5
                }
            }catch {
                # Gestion des erreurs au cas où Get-Hotfix échouerait
                Write-Host "Erreur lors de la récupération des correctifs sur $serverName."
                $tophotfixes = $null
            }
            #On initialise un booleen pour dire que le check de l'eventvwr n'a pas ete fait
            [bool]$VerifKBeventvwr = $false
            #On compare les kb de la liste avec les KB installes sur le serveur
            foreach($tophotfixe in $tophotfixes){
                foreach($KB in $KBList){
                    if ($KB.KBID -eq $tophotfixe.HotFixId){
                        Write-Host  "Match Found: KB A install=" $KB.KBID " hotfixes installe=" $tophotfixe.HotFixId
                        #Vu qu'il y a eu un match entre le KB sur la liste et le KB sur le serveur on passe le booleen en true
                        $ServerKBinstall = $true
                        #On rajoute le nom du KB dans l'array
                        $arrayKB += $KB.KBID
                        #On va check l'eventvwr pour voir si il y a bien un event install ok
                        $kbevents = Invoke-Command -ComputerName $serverName -Credential $credential -ScriptBlock {
                            Get-WinEvent -LogName Setup | Where-Object { $_.Id -eq 2 } | 
                            Select-Object -ExpandProperty Message
                        }
                        foreach ($kbevent in $kbevents) {
                            if ($kbevent.Contains($KB.KBID)) {
                                Write-Host "Verification de l'eventvwr:"
                                Write-Host -ForegroundColor green "Event d'installation OK: $kbevent"
                                #La verif est ok on passe le booleen du check eventvwr en true
                                $VerifKBeventvwr = $true
                            }
                        }
                        if ($VerifKBeventvwr -eq $false){
                            Write-Host -ForegroundColor red "Pas d'evenement d'installation dans l'eventvwr pour le "$KB.KBID
                        }
                    }
                }
            }
            #On append au csv quand tout est ok
            if ($ServerKBinstall -eq $true){
                Write-Host -ForegroundColor green "Le serveur" $serverName "a installe les KB" $arrayKB "`n`n`n"
                $reportdata += [PSCustomObject]@{
                    ComputerName = $serverName
                    OS = $osversion.Caption
                    KB_Install_Pending = "N/A"
                    Reboot_Required = $rebootRequired
                    Last_Reboot = $lastBootTimeFormatted
                    KB_On_Server = "OK"
                    Verif_Install_EventVWR = $VerifKBeventvwr
                    KB_Installed = $arrayKB  -join ', '
                }
            }
            else{
                Write-Host -ForegroundColor red "`n`n Aucun KB de la liste n'a ete installe sur le " $serverName
                #On check si des KB sont en pending
                $updatePending = Invoke-Command -ComputerName $serverName -Credential $credential -ScriptBlock {
                    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
                    $UpdateSearcher = $UpdateSession.CreateupdateSearcher()
                    $Updates = @($UpdateSearcher.Search("IsHidden=0 and IsInstalled=0").Updates)
                    $Updates | Select-Object -ExpandProperty Title
                }
                if ($updatePending){
                    $kbPendingArray = @()
                    foreach ($update in $updatePending) {
                        if ($update -match "KB\d{7}") {
                            $kbPendingArray += $matches[0]
                        }
                    }
                    # Display the KB numbers
                    $kbPending = $kbPendingArray -join ', '
                }
                else{
                    $updatePending = $false
                }
                #On append au csv mais on signale que le check de l'eventvwr est ko
                $reportdata += [PSCustomObject]@{
                    ComputerName = $serverName
                    OS = $osversion.Caption
                    KB_Install_Pending = $kbPending
                    Reboot_Required = $rebootRequired
                    Last_Reboot = $lastBootTimeFormatted
                    KB_On_Server = "KO"
                    Verif_Install_EventVWR = $VerifKBeventvwr
                    KB_Installed = "KB non present sur le serveur"
                }
            }
        } catch {
            Write-Host -ForegroundColor yellow "Erreur lors de l'execution de Get-Hotfix sur $serverName : $_"
        }
}
else {
    Write-Host "Le serveur $serverName est un " $osversion.Caption " => Pas de KB a install`n"
    #On ajoute les infos au csv concernant les serveurs obsoletes
    $reportdata += [PSCustomObject]@{
        ComputerName = $serverName
        OS = $osversion.Caption
        KB_Install_Pending = "Os plus maintenu"
        Reboot_Required = "N/A"
        Last_Reboot = $lastBootTimeFormatted
        KB_On_Server = "N/A"
        Verif_Install_eventvwr = "N/A"
        KB_Installed = "N/A"
    }
}
}
#On genere le csv
$reportdata | Export-Csv -Path $reportcsvpath -NoTypeInformation -Encoding UTF8

Write-Host "Rapport d'analyse: $reportcsvpath"
Write-Host "Copie du rapport pour generation du mail"
