# Script PowerShell : Vérification et Rapport des KB installés

Ce script PowerShell permet de vérifier si des KB spécifiques sont installés sur une liste de serveurs définie dans un fichier CSV. Il génère ensuite un rapport détaillé en CSV.

## Fonctionnalités

- Vérifie si les serveurs sont encore maintenu (exclu les versions en dessous de server 2012).
- Compare une liste de KB fourni en csv avec les derniers correctifs installés sur les serveurs.
- Vérifie les événements d'installation dans l'Event Viewer pour confirmer le succès des installations.
- Signale si des KB sont en attente d'installation.
- Génère un rapport CSV contenant les informations pour chaque serveur analysé.

## Prérequis

- **Accès administrateur** sur les serveurs à vérifier.
- **PowerShell** version 5.1 ou ultérieure.
- Liste des serveurs et des KB à vérifier dans des fichiers CSV.

## Utilisation

### Fichier CSV Serveurs 

Attention a bien respecter la nomenclature Computers:


![image](https://github.com/user-attachments/assets/4086d028-232b-4f24-a13f-af5376a4ac77)

### Fichier CSV KB
Attention a bien respecter la nomenclature KBID:


![image](https://github.com/user-attachments/assets/ab32cec2-d057-4ceb-bd18-2ac2084a3bc0)

### Syntaxe

```powershell
.\checkinstallkb.ps1 -csvPath <chemin_vers_csv_serveurs> -csvPath2 <chemin_vers_csv_kb>

