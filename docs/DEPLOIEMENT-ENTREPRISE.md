# Déploiement en entreprise — profil durci

Ce guide décrit comment déployer **Milestone Toolkit** de façon **maîtrisée et hors-ligne**, adaptée
à un environnement d'entreprise avec validation IT / sécurité.

Objectif du profil durci : **aucun appel Internet** (ni GitHub, ni PowerShell Gallery), **pas
d'auto-mise à jour**, et **moindre privilège** côté Milestone.

---

## 1. Préparer une version validée (une seule fois)

1. Récupérer une version depuis la page **Releases** et la faire **revoir par l'IT/sécurité**.
2. Sur un poste **avec** Internet, préparer les dépendances hors-ligne :
   - Lancer l'outil → écran de démarrage → bouton **« Préparer offline »**
   - Cela télécharge **MilestonePSTools** et **ImportExcel** dans le dossier `Dependencies/`.
3. La version ainsi préparée (dossier complet **avec `Dependencies/`**) devient le **paquet de
   référence** validé, à copier tel quel sur les postes cibles.

> Une fois `Dependencies/` présent, l'outil bascule automatiquement en **mode hors-ligne** : il
> charge les modules localement et **ne contacte plus PowerShell Gallery**.

---

## 2. Configurer le profil entreprise

Éditer `config.json` à la racine pour **désactiver la mise à jour automatique** :

```json
{
    "outputDirectory": "./Output",
    "snapshotQuality": 95,
    "csvDelimiter": ";",
    "csvEncoding": "UTF8",
    "language": "fr",
    "autoLogin": false,
    "autoUpdate": {
        "enabled": false
    }
}
```

Effet de `"autoUpdate": { "enabled": false }` : au démarrage, l'outil **ne contacte plus GitHub**
(la vérification de mise à jour retourne immédiatement, sans aucune requête réseau).

Un modèle prêt à l'emploi est fourni : **[config.entreprise.exemple.json](config.entreprise.exemple.json)**.

---

## 3. Compte Milestone en moindre privilège

- **Rapports uniquement** → un compte **lecture seule** suffit (aucune écriture dans la config VMS).
- **Groupes / alarmes** (fonctions d'écriture) → un compte aux droits **strictement limités** à ces
  opérations, sur le périmètre concerné.
- **Ne jamais** utiliser un compte administrateur nominatif.
- Laisser `"autoLogin": false` : le serveur n'est pas mémorisé/reconnecté automatiquement.

---

## 4. Distribution

- Héberger le paquet validé sur un **partage interne de confiance** (pas de retéléchargement Internet).
- Les postes reçoivent une **version figée** ; les mises à jour sont **gérées par l'IT** (remplacement
  du paquet après revue d'une nouvelle version), pas par l'outil.

---

## 5. Flux réseau en profil entreprise

| Destination | Contactée ? |
|---|---|
| Serveur Milestone Management Server | ✅ Oui (fonction principale) |
| api.github.com / github.com | ❌ Non (auto-update désactivé) |
| www.powershellgallery.com | ❌ Non (modules embarqués) |

**Seul le Management Server est contacté.** À vérifier au besoin par capture réseau lors de la revue.

---

## 6. Checklist de validation IT / sécurité

| # | Contrôle | OK |
|---|---|---|
| 1 | Version revue par l'IT et figée (paquet de référence) | ☐ |
| 2 | `Dependencies/` présent (mode hors-ligne actif) | ☐ |
| 3 | `autoUpdate.enabled` = `false` dans `config.json` | ☐ |
| 4 | Compte Milestone en moindre privilège (lecture seule si rapports) | ☐ |
| 5 | `autoLogin` = `false` | ☐ |
| 6 | Distribution depuis un partage interne de confiance | ☐ |
| 7 | Capture réseau : seul le Management Server est contacté | ☐ |
| 8 | Dossier `Output/` (exports) sur un emplacement maîtrisé | ☐ |
| 9 | Export des mots de passe caméras interdit/encadré (colonne décochée) | ☐ |
| 10 | *(optionnel)* Scripts signés Authenticode → `Bypass` retiré | ☐ |

---

## 7. (Optionnel) Signature de code

Pour supprimer le `-ExecutionPolicy Bypass` :

1. L'entreprise fournit un **certificat de signature de code** (interne ou public).
2. Signer les fichiers `.ps1` avec `Set-AuthenticodeSignature`.
3. Distribuer avec une politique `AllSigned` ou `RemoteSigned`, et adapter le lanceur `.bat` en
   conséquence (retrait du `Bypass` et de `Unblock-File`).

Sur demande, une procédure détaillée + un script de signature peuvent être fournis.
