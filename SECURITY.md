# Security · Sécurité

Ce document décrit le fonctionnement de sécurité de **Milestone Toolkit** à destination des
équipes IT / sécurité. Objectif : permettre une **revue et une validation** de l'outil, et fournir
un **profil de déploiement durci** pour un usage en entreprise.

> 🇬🇧 A concise English summary is provided at the end of each section.

---

## 1. Ce que fait l'outil

Milestone Toolkit est une application **PowerShell + interface WPF** qui s'appuie sur le module
officiel **MilestonePSTools** pour interagir avec un serveur Milestone XProtect.

**Opérations en LECTURE** (rapports) : liste des caméras/matériel, rétention, statistiques
d'enregistrement, état des caméras, informations de licence, snapshots. Sorties en Excel / CSV / JPEG.

**Opérations en ÉCRITURE** (modifient la configuration VMS) :
- Création de **groupes de caméras** (« Grouper par modèle »)
- Création de **définitions d'alarme** en masse (réversibles via le Management Client)

> 🇬🇧 A PowerShell/WPF tool built on MilestonePSTools. Read operations = reports (Excel/CSV/snapshots).
> Write operations = camera groups and alarm definitions (both reversible in the Management Client).

---

## 2. Connexions réseau (liste exhaustive)

| Destination | Quand | But | Désactivable |
|---|---|---|---|
| **Serveur Milestone Management Server** | À chaque lancement | Fonction principale (SDK Milestone) | Non (c'est le cœur de l'outil) |
| `api.github.com` | Au démarrage | Vérifier s'il existe une mise à jour | **Oui** — `autoUpdate.enabled: false` |
| `github.com` / `codeload.github.com` | Clic sur « Mettre à jour » | Télécharger l'archive de la nouvelle version | **Oui** — idem |
| `www.powershellgallery.com` (+ CDN NuGet Microsoft) | Uniquement si un module requis est **absent** | Installer MilestonePSTools / ImportExcel | **Oui** — mode hors-ligne (modules embarqués) |

**En profil entreprise (hors-ligne + mise à jour désactivée), la SEULE connexion sortante est le
Management Server Milestone.** Aucun appel à GitHub ni à PowerShell Gallery.

> 🇬🇧 Only two external hosts are ever contacted: GitHub (auto-update, disableable) and PowerShell
> Gallery (module install, avoided in offline mode). In the enterprise profile, the **only** outbound
> connection is the Milestone Management Server.

---

## 3. Données manipulées

| Donnée | Sensibilité | Traitement |
|---|---|---|
| IP, MAC, modèle, firmware des caméras | Interne | Export Excel/CSV, dans le dossier de sortie local |
| Topologie / groupes / licences | Interne | Idem |
| **Mots de passe caméras** | **Sensible** | **Exclus par défaut.** Exportés en clair **uniquement** si la colonne est cochée explicitement |
| Snapshots (images) | Selon contexte | Fichiers JPEG dans le dossier de sortie local |
| Identifiants du serveur Milestone | Sensible | Saisis dans le dialogue Milestone ; jamais stockés par l'outil |

- Les fichiers générés restent **en local** (dossier `Output/` par défaut). L'outil ne les envoie nulle part.
- Journaux d'activité horodatés dans `Logs/` (purge automatique > 30 jours).

> 🇬🇧 Exports stay local. Camera passwords are excluded by default and only exported in clear text if
> the column is explicitly ticked. The tool never transmits generated files anywhere.

---

## 4. Permissions requises

- **Windows** : aucun droit administrateur requis pour l'usage courant. L'outil s'exécute en
  utilisateur standard ; les modules PowerShell s'installent en `CurrentUser`.
- **Compte Milestone** : appliquer le **moindre privilège**.
  - Rapports uniquement → un compte **lecture seule** suffit.
  - Fonctions d'écriture (groupes, alarmes) → un compte aux droits **limités à ces opérations**.
  - **Ne pas** utiliser un compte administrateur personnel.

> 🇬🇧 No admin rights needed for normal use. Use a **least-privilege** Milestone account (read-only for
> reporting; scoped write rights only if you use groups/alarms).

---

## 5. Modèle d'exécution (transparence)

Points à connaître pour une revue de sécurité, et comment les neutraliser :

| Comportement | Raison | Durcissement |
|---|---|---|
| Lancement en `-ExecutionPolicy Bypass` | Exécuter des scripts non signés téléchargés en `.zip` | **Signer le code** (Authenticode) puis passer en `AllSigned`/`RemoteSigned` |
| `Unblock-File` au 1er lancement | Lever le « Mark of the Web » sur les fichiers du `.zip` | Distribuer via un partage interne de confiance |
| **Auto-update** depuis GitHub | Confort de mise à jour | **Désactiver** (`autoUpdate.enabled: false`) et figer une version revue |
| Installation de modules depuis PSGallery | Confort d'installation | **Mode hors-ligne** (modules embarqués, aucun accès Internet) |

Le dépôt de mise à jour est **figé dans le code** : une modification de `config.json` ne peut pas
rediriger les mises à jour vers un autre dépôt. Les téléchargements se font en **HTTPS**.

> 🇬🇧 The tool runs with ExecutionPolicy Bypass and unblocks files (to run the downloaded .zip). Both
> can be removed via **code signing**. Auto-update and PSGallery installs can be fully disabled
> (enterprise offline profile). The update repository is **pinned in code**; downloads are HTTPS-only.

---

## 6. Déploiement durci en entreprise

Voir **[docs/DEPLOIEMENT-ENTREPRISE.md](docs/DEPLOIEMENT-ENTREPRISE.md)** pour la procédure complète :
profil hors-ligne, désactivation de la mise à jour, compte de moindre privilège, version figée, et
une **checklist de validation** pour l'IT.

> 🇬🇧 See **docs/DEPLOIEMENT-ENTREPRISE.md** for the hardened, offline enterprise deployment guide and
> an IT validation checklist.

---

## 7. Signaler une vulnérabilité · Reporting a vulnerability

Merci de signaler tout problème de sécurité de manière **privée** via un
[Security Advisory GitHub](../../security/advisories/new) plutôt que par une issue publique.

*Please report security issues privately via a GitHub Security Advisory rather than a public issue.*
