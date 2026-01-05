# 🖥️ Remote Session Manager (PRT)

> Utilitaire graphique de gestion de sessions PowerShell distantes pour Windows.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue?logo=powershell)
![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

## 📋 Description

**Remote Session Manager** est un outil PowerShell avec interface WPF qui permet de gérer facilement des sessions distantes vers des machines Windows. Conçu pour les équipes support/helpdesk, il simplifie la connexion et l'administration à distance des postes de travail.

## ✨ Fonctionnalités

- 🔗 **Connexion PSSession** - Établissement de sessions PowerShell distantes
- 💻 **Console intégrée** - Exécutez des commandes à distance directement dans l'interface
- 📜 **Historique des commandes** - Navigation avec les flèches ↑/↓
- 🔍 **Scanner réseau** - Détection des postes via Active Directory ou scan IP
- ⏱️ **Timer de session** - Suivi du temps de connexion
- 🎨 **Thèmes personnalisables** - Interface claire ou sombre
- 📤 **Export des logs** - Sauvegarde de l'historique des commandes
- 🧹 **Nettoyage profil admin** - Suppression automatique du profil admin à la déconnexion

### Raccourcis clavier

| Raccourci | Action |
|-----------|--------|
| `Enter` | Envoyer la commande |
| `↑` / `↓` | Naviguer dans l'historique |
| `Ctrl+L` | Effacer la console |
| `Ctrl+S` | Exporter le log |
| `F5` | Rafraîchir |

## 🚀 Installation

### Prérequis

- Windows 10/11 ou Windows Server 2016+
- PowerShell 5.1 ou supérieur
- [WinRM] activé sur les machines distantes
- Droits d'administration sur les postes cibles

### Lancement

```powershell
# Méthode 1 : Lancer directement
.\RemoteSessionManager.ps1

# Méthode 2 : Via le lanceur
.\Lanceur.cmd
```

## 📖 Utilisation

1. **Entrez le numéro ou nom du poste** dans le champ de saisie
2. **Cliquez sur "Connecter"** pour établir la session
3. **Utilisez la console** pour exécuter des commandes à distance
4. **Utilisez les actions rapides** dans le panneau de droite
5. **Cliquez sur "Déconnecter"** pour fermer proprement la session

### Options de connexion

- **Préfixe PRT** : Par défaut, le préfixe "PRT" est ajouté automatiquement aux numéros (ex: PRT001)
- **Sans PRT** : Cochez la case pour utiliser un nom de machine personnalisé

## 🛠️ Configuration WinRM

Sur les machines distantes, WinRM doit être activé :

```powershell
# Activer WinRM (en admin)
Enable-PSRemoting -Force

# Vérifier la configuration
winrm quickconfig
```

## 📁 Structure du projet

```
├── RemoteSessionManager.ps1   # Script principal avec interface WPF
├── Lanceur.cmd                # Lanceur batch
└── README.md                  # Documentation
```

## 🔧 Fonctionnalités avancées

### Scanner réseau

Trois méthodes de scan disponibles :
1. **Active Directory** - Recherche des ordinateurs dans l'AD (recommandé)
2. **Scan plage IP** - Scan d'une plage d'adresses IP spécifique
3. **Scan plage PRT** - Scan des postes avec préfixe PRT

### Console externe

Possibilité d'ouvrir une fenêtre PowerShell séparée avec `Enter-PSSession` pour un accès interactif complet.

## 👤 Auteur

**Hotline6 By Micka** - Version 4.1

## 📄 License

Ce projet est sous licence MIT - voir le fichier [LICENSE](LICENSE) pour plus de détails.

---

⭐ *Si cet outil vous est utile, n'hésitez pas à le partager !*
