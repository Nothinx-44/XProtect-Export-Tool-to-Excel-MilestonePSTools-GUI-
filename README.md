<div align="center">

# 🎥 Milestone Toolkit

**Administration & export tool for Milestone XProtect VMS**
*Outil d'administration et d'export pour Milestone XProtect VMS*

Modern graphical interface (WPF) built on the **MilestonePSTools** PowerShell module —
export cameras, hardware, retention and recordings to **Excel / CSV**, capture snapshots,
manage groups and create alarms **in bulk**.

[![Total downloads](https://img.shields.io/github/downloads/Nothinx-44/XProtect-Export-Tool-to-Excel-MilestonePSTools-GUI-/total?label=Downloads%20%C2%B7%20T%C3%A9l%C3%A9chargements&labelColor=1E1E2E&color=A6E3A1&logo=github&style=for-the-badge)](../../releases)

![Windows](https://img.shields.io/badge/Windows-10%2F11%20%7C%20Server%202016%2B-0078D6?logo=windows)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1-5391FE?logo=powershell)
![Languages](https://img.shields.io/badge/UI-EN%20%7C%20FR-brightgreen)
![Milestone](https://img.shields.io/badge/Milestone-XProtect-orange)
![Latest release](https://img.shields.io/github/v/release/Nothinx-44/XProtect-Export-Tool-to-Excel-MilestonePSTools-GUI-?label=version&color=CBA6F7)

<img width="1145" height="893" alt="image" src="https://github.com/user-attachments/assets/f64781e2-9b82-49ad-8f03-6f4e19dd3806" />


**🌐 [English](#-english) · [Français](#-français)**

</div>

---

# 🇬🇧 English

## Why this tool

Milestone XProtect has no built-in way to easily export its configuration in a usable format.
Milestone Toolkit fills that gap and saves a considerable amount of time on large installations:

- 📊 **Full export** of cameras and hardware to Excel (with snapshots)
- 🕵️ **Audit** recording retention and statistics
- 📸 **Bulk snapshots** (live or historical)
- 🚨 **Bulk alarm creation** in a few clicks
- 🌍 **Bilingual** interface (English / French), dark theme

## Getting started

1. Download the latest version from the [**Releases**](../../releases) page
2. Unzip the archive
3. Double-click **`Demarrer Milestone Toolkit.bat`**

> On first launch, the startup screen checks dependencies and installs them automatically if
> Internet is available. No manual installation required.

## Features

### 📸 Snapshots

| Action | Description |
|--------|-------------|
| **Snapshot – Selection** | Capture the selected camera via the Milestone dialog |
| **Snapshot – All cameras** | Capture every camera in the system, with progress and cancellation |
| **Snapshot – PTZ presets** | Iterate over PTZ presets and capture an image at each position |

Each action supports two modes: **Live** (latest image) or **Historical** (closest image to a date/time).

### 🛠️ Management

| Action | Description |
|--------|-------------|
| **Hardware Export (Excel)** | Configurable Excel report: pick columns (hardware, video streams, retention, snapshots). Passwords are **excluded by default**. Reliable **recorded vs live** stream detection per camera. |
| **Group by Model** | Creates camera groups in Milestone, organized by model |
| **🚨 Create alarms** | **Bulk alarm creation** — see dedicated section below |

<img width="513" alt="Export column selector" src="https://github.com/user-attachments/assets/5bb4cf19-e0f3-4684-9fce-6d1d328feff4" />

### 📡 Monitoring

| Action | Description |
|--------|-------------|
| **Camera status** | Real-time status (OK / offline / error) via the Event Server. Disabled cameras are ignored. CSV export. |
| **Recording dates** | First and last available recording per camera, with total retention. CSV export. |

### 🔍 Diagnostics

| Action | Description |
|--------|-------------|
| **Recording stats (7 days)** | Recording and motion statistics per camera over 7 days (FPS, bitrate, resolution). CSV export. |
| **License information** | Licensed products, expiration dates and used channels |

## 🚨 Bulk alarm creation

The **Create alarms** button opens a complete, flexible window to generate several alarm
definitions at once — ideal for applying the same alarm (e.g. *camera offline*) across a whole fleet.

**Two modes:**
- **New alarm** — pick the event group, type, priority and category. The lists are
  **populated dynamically from your server** (no hard-coded values).
- **Duplicate an existing one** — reuses the settings of an already-configured alarm.

**Scope:**
- All cameras · a selection · **a single global alarm** or **one alarm per camera**
  (name via a `{camera}` template).

**Safe by design:** a pre-flight check validates the chosen event type **before** any bulk
creation, and alarms remain removable (Management Client or `Remove-VmsAlarmDefinition`).

## Requirements

- **Windows** 10 / 11 or Windows Server 2016+
- **PowerShell 5.1** (built into Windows)
- **Excel optional**: if absent, Hardware Export automatically falls back to the *ImportExcel*
  module (no Office installation required)
- Network access to the Milestone XProtect **Management Server**

> The **MilestonePSTools** and **ImportExcel** modules are installed automatically on first launch
> if Internet is available.

## Installation modes

### Online (default)
Modules are downloaded automatically from PowerShell Gallery. No action required.

### Offline (machine without Internet)
1. On a machine **with** Internet, click **Prepare offline** on the startup screen
2. Copy the entire project (including the `Dependencies/` folder) to the target machine
3. Launch normally — Offline mode is detected automatically

<img width="562" alt="Startup / dependency check screen" src="https://github.com/user-attachments/assets/da83ab43-56b6-4d38-97b6-607831601588" />

## Languages & updates

- **Languages** — English and French. Chosen on first launch, then stored in `config.json`
  (key `language`).
- **Automatic updates** — on startup, the tool checks the latest GitHub release and offers a
  one-click update. The archive is **verified** (pinned official GitHub repo + **SHA256** checksum
  published in the release notes), and the previous version is **backed up** with **automatic
  rollback** on failure.

## Configuration

`config.json`:

```json
{
    "outputDirectory": "./Output",
    "snapshotQuality": 95,
    "csvDelimiter": ";",
    "csvEncoding": "UTF8",
    "language": "fr",
    "autoLogin": false,
    "autoUpdate": {
        "enabled": true,
        "repo": "Nothinx-44/XProtect-Export-Tool-to-Excel-MilestonePSTools-GUI-"
    }
}
```

| Setting | Description | Default |
|---------|-------------|---------|
| `outputDirectory` | Output folder (snapshots, Excel, CSV) | `./Output` |
| `snapshotQuality` | JPEG quality of snapshots (1–100) | `95` |
| `csvDelimiter` | CSV field separator | `;` |
| `csvEncoding` | CSV file encoding | `UTF8` |
| `language` | UI language (`fr` / `en`), remembered after first launch | — |
| `autoLogin` | Auto-connect the Milestone dialog to the last server. Disabled by default (a stale server would break startup) | `false` |
| `autoUpdate.enabled` | Check for updates on startup | `true` |
| `autoUpdate.repo` | Informative: the update repo is **pinned in code** for security | — |

## Use cases

Camera fleet audit · retention verification · client-facing reports · system maintenance ·
bulk alarm deployment.

---

# 🇫🇷 Français

## Pourquoi cet outil

Milestone XProtect ne propose pas nativement d'export simple et exploitable de la configuration.
Milestone Toolkit comble ce manque et fait gagner un temps considérable sur les grosses installations :

- 📊 **Export complet** des caméras et du matériel vers Excel (snapshots inclus)
- 🕵️ **Audit** de la rétention et des enregistrements
- 📸 **Snapshots** en masse (live ou historique)
- 🚨 **Création d'alarmes en masse** en quelques clics
- 🌍 Interface **bilingue** (français / anglais), thème sombre

## Lancement

1. Télécharger la dernière version depuis la page [**Releases**](../../releases)
2. Décompresser l'archive
3. Double-cliquer sur **`Demarrer Milestone Toolkit.bat`**

> Au premier lancement, l'écran de démarrage vérifie les dépendances et les installe
> automatiquement si Internet est disponible. Aucune installation manuelle requise.

## Fonctionnalités

### 📸 Snapshots

| Action | Description |
|--------|-------------|
| **Snapshot – Sélection** | Capture la caméra sélectionnée via le dialogue Milestone |
| **Snapshot – Toutes les caméras** | Capture chaque caméra du système, avec progression et annulation |
| **Snapshot – Presets PTZ** | Parcourt les presets PTZ et capture une image à chaque position |

Chaque action supporte deux modes : **Live** (dernière image) ou **Historique** (image la plus proche d'une date/heure).

### 🛠️ Gestion

| Action | Description |
|--------|-------------|
| **Export Hardware (Excel)** | Rapport Excel configurable : sélection des colonnes (matériel, flux vidéo, rétention, snapshots). Les mots de passe sont **exclus par défaut**. Détection fiable du flux **enregistré vs live** par caméra. |
| **Grouper par Modèle** | Crée des groupes de caméras dans Milestone, organisés par modèle |
| **🚨 Créer des alarmes** | **Création d'alarmes en masse** — voir section dédiée ci-dessous |

### 📡 Monitoring

| Action | Description |
|--------|-------------|
| **État des caméras** | État temps réel (OK / hors ligne / erreur) via l'Event Server. Caméras désactivées ignorées. Export CSV. |
| **Dates d'enregistrement** | Premier et dernier enregistrement disponible par caméra, avec durée totale de rétention. Export CSV. |

### 🔍 Diagnostic

| Action | Description |
|--------|-------------|
| **Stats Enregistrement (7 j)** | Statistiques d'enregistrement et de mouvement par caméra sur 7 jours (FPS, bitrate, résolution). Export CSV. |
| **Informations Licence** | Produits licenciés, dates d'expiration et canaux utilisés |

## 🚨 Création d'alarmes en masse

Le bouton **Créer des alarmes** ouvre une fenêtre complète et flexible pour générer plusieurs
définitions d'alarme d'un coup — idéal pour appliquer une même alarme (ex. *caméra hors ligne*)
à tout un parc.

**Deux modes :**
- **Nouvelle alarme** — choix du groupe d'événement, du type, de la priorité et de la catégorie.
  Les listes sont **peuplées dynamiquement depuis votre serveur** (aucune valeur codée en dur).
- **Dupliquer une existante** — reprend les réglages d'une alarme déjà configurée.

**Portée au choix :**
- Toutes les caméras · une sélection · **une seule alarme globale** ou **une alarme par caméra**
  (nom via un modèle `{camera}`).

**Sûr par conception :** un test préalable valide le type d'événement choisi **avant** toute
création en masse, et les alarmes restent supprimables (Management Client ou `Remove-VmsAlarmDefinition`).

## Prérequis

- **Windows** 10 / 11 ou Windows Server 2016+
- **PowerShell 5.1** (inclus dans Windows)
- **Excel facultatif** : s'il est absent, l'export Hardware bascule automatiquement sur le module
  *ImportExcel* (aucune installation d'Office requise)
- Accès réseau au **Management Server** Milestone XProtect

> Les modules **MilestonePSTools** et **ImportExcel** sont installés automatiquement au premier
> lancement si Internet est disponible.

## Modes d'installation

### En ligne (par défaut)
Les modules sont téléchargés automatiquement depuis PowerShell Gallery. Aucune action requise.

### Hors ligne (machine sans Internet)
1. Sur une machine **avec** Internet, cliquer **Préparer offline** dans l'écran de démarrage
2. Copier le projet entier (avec le dossier `Dependencies/`) sur la machine cible
3. Lancer normalement — le mode Offline est détecté automatiquement

## Langues & mises à jour

- **Langues** — français et anglais. Le choix est demandé au premier lancement puis mémorisé dans
  `config.json` (clé `language`).
- **Mise à jour automatique** — au démarrage, l'outil vérifie la dernière release GitHub et propose
  la mise à jour en un clic. L'archive est **vérifiée** (dépôt GitHub officiel figé + empreinte
  **SHA256** publiée dans les notes de release), et l'ancienne version est **sauvegardée** avec
  **restauration automatique** en cas d'échec.

## Configuration

Voir le tableau de configuration dans la section anglaise ci-dessus — les paramètres de `config.json`
sont identiques : dossier de sortie, qualité des snapshots, délimiteur/encodage CSV, langue,
`autoLogin` et `autoUpdate`.

## Cas d'usage

Audit de parc caméras · vérification de la rétention · export client · maintenance système ·
déploiement d'alarmes en masse.

---

<div align="center">

## 📁 Project structure

```
Milestone Toolkit/
├── Demarrer Milestone Toolkit.bat   # Entry point (double-click)
├── config.json                      # Configuration
├── Save-Dependencies.ps1            # Offline preparation
├── Dependencies/                    # Offline modules (optional)
├── Logs/                            # Daily logs (auto, purged > 30 days)
├── Output/                          # Generated files (auto)
└── src/
    ├── Bootstrap.ps1                # Startup: language, checks, updates
    ├── App.ps1                      # UI loading and events
    ├── Version.ps1                  # Version number (single source)
    ├── UI/MainWindow.xaml           # WPF interface (Catppuccin dark theme)
    ├── Lang/                        # Translations fr.ps1 / en.ps1
    ├── Actions/                     # Snapshots, export, alarms, monitoring…
    └── Core/                        # Modules, updater, logging, console…
```

</div>

---

<div align="center">

## 🙏 Acknowledgements · Remerciements

Special thanks to **Tony Mattina** for his help in creating this tool.
*Un grand merci à **Tony Mattina** pour son aide dans la création de cet outil.*

Built on the **MilestonePSTools** PowerShell module for Milestone XProtect.

**Made by Vincent Le Bonhomme**

<sub>Milestone XProtect · export to Excel · MilestonePSTools GUI · CCTV audit tool · camera report ·
export XProtect cameras · hardware export · video surveillance reporting</sub>

</div>
