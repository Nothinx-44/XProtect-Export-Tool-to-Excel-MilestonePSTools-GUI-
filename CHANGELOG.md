# Release Notes

## v4.9.12
> Mises a jour comptees dans les telechargements GitHub

### Ameliorations
- **Show-StartupCheck.ps1** : la mise a jour telecharge desormais le **fichier .zip attache** a la release (au lieu de l'archive source auto-generee). GitHub comptabilise ces telechargements — les mises a jour via le logiciel comptent donc dans le total. **Repli automatique** sur l'archive source si aucun asset n'est present (anciennes releases).
- Validation d'URL elargie pour couvrir les deux formats (asset et zipball). Integrite assuree par HTTPS + depot fige.
- **CI** : le workflow construit un `.zip` propre (`git archive`) et l'**attache** a chaque release. Un vrai bouton de telechargement + un compteur fiable, sans intervention.

> Note : la ligne `sha256:` dans les notes a ete abandonnee — une seule empreinte ne peut pas convenir a la fois aux anciennes versions (qui verifient le zipball) et aux nouvelles (qui verifient l'asset). L'integrite reste assuree par le telechargement HTTPS depuis le depot fige.

---

## v4.9.11
> Nouveau : creation d'alarmes en masse. + Export Hardware : detection fiable du flux enregistre vs live.

### Nouveautes
- **Creation d'alarmes en masse** (nouveau bouton dans Gestion). Une fenetre permet de :
  - creer une **nouvelle alarme** en choisissant le type d'evenement, la priorite et la categorie — les listes sont **peuplees dynamiquement depuis le serveur** (pas de valeurs codees en dur, chaque installation voit ses propres options) ;
  - ou **dupliquer une alarme existante** (reprend type / priorite / categorie) ;
  - avec une **portee** au choix : toutes les cameras, une selection, en **une seule alarme globale** ou **une alarme par camera** (nom via un modele `{camera}`).
  - Groupes d'evenement lus via `EventTypeGroupValues` (GUID attendus par le serveur) et types d'evenement via la hierarchie de configuration — contourne `ValidateItem` (non supporte sur certains SDK).
  - **Test prealable** : avant toute creation en masse, une alarme jetable valide le type d'evenement choisi. Si le serveur le refuse (evenement non declencheur d'alarme), l'utilisateur est prevenu clairement AVANT de creer des alarmes en echec.
  - Progression et annulation gerees ; les alarmes sont creees via `New-VmsAlarmDefinition` (reversibles dans le Management Client ou via `Remove-VmsAlarmDefinition`).

### Corrections
- **Export-HardwareReport.ps1** : l'identification des flux enregistrement / live etait faussee par un `if/elseif` qui ne classait chaque flux que dans UNE seule categorie. Un flux a la fois enregistre ET live par defaut (cas le plus courant) etait compte uniquement comme "enregistre", laissant les colonnes live vides ou decalees ; en multi-track, le 2e flux enregistre etait compte comme "supplementaire".
- Nouvelle logique : les flux sont d'abord regroupes **par camera**, puis l'enregistre et le live sont identifies **independamment** via les vrais drapeaux de configuration (`Recorded`, `LiveDefault`), avec preference pour la piste d'enregistrement primaire en multi-track. Un meme flux peut legitimement etre a la fois enregistre et live.
- La comparaison "meme flux physique" se base desormais sur `StreamReferenceId` (au lieu du nom).

Resultat : les colonnes Codec / Resolution / FPS (enregistrement et live) refletent la configuration reelle de chaque camera, y compris mono-flux, bi-flux et multi-track.

---

## v4.9.10
> Connexion : auto-login du dialogue Milestone desactive par defaut

### Corrections
- **App.ps1** : le dialogue de connexion Milestone est desormais ouvert avec `-DisableAutoLogin`. Si un ancien serveur (inaccessible) avait ete memorise avec "Auto login", le dialogue se connectait tout seul et l'application plantait au demarrage sans laisser changer de serveur. Le dialogue s'affiche maintenant normalement (adresse pre-remplie) et attend la validation. Reactivable via `"autoLogin": true` dans `config.json`.

---

## v4.9.9
> Quatrieme audit — bugs PTZ et stats video, demarrage plus rapide, README a jour

### Corrections
- **Get-RecordingStats.ps1** : les colonnes FPS/Bitrate/Resolution etaient TOUJOURS 'N/A' — `$cam | Get-VideoDeviceStatistics` liait l'Id de la camera au parametre `RecordingServerId` (aucun resultat). Remplace par un appel groupe unique (parallelise par serveur d'enregistrement) avec lookup par `DeviceId` : le bug est corrige ET l'action est beaucoup plus rapide.
- **Invoke-PtzPreset.ps1** : les echecs (`Write-Error` non-terminant) etaient invisibles — la console est masquee et le `catch` de l'appelant ne se declenchait jamais. Remplace par `throw` : les erreurs PTZ apparaissent desormais dans le journal.
- **Invoke-PtzPreset.ps1** : la verification de position comparait `Abs(Abs(a)-Abs(b))` — une camera a pan -0.5 etait jugee arrivee sur un preset a +0.5. Corrige en `Abs(a-b)`.
- **Export-HardwareReport.ps1** : liberation COM complete (sheet + workbook + excel) suivie d'un GC force — evite les processus Excel.exe zombies apres export.
- **Export-HardwareReport.ps1** : plus de `Install-Module` silencieux en pleine action — le telechargement d'ImportExcel est annonce dans le journal et refuse en mode Offline.

### Ameliorations
- **Demarrer Milestone Toolkit.bat** : le deblocage des fichiers (`Unblock-File` recursif sur tout le projet, y compris le SDK) ne s'execute plus qu'une seule fois (fichier temoin `.unblocked`, efface par l'updater apres mise a jour). Demarrage nettement plus rapide.
- **Write-ActivityLog.ps1** : purge automatique des logs de plus de 30 jours au demarrage (`Remove-OldLogs`).
- **CI** : le workflow de release refuse desormais la publication si `src/Version.ps1` ne correspond pas au tag (protege l'auto-updater).
- **Export-HardwareReport.ps1** : le calcul des valeurs de colonnes est factorise (`Get-CameraRowValues`), source unique pour les chemins COM et ImportExcel (2 regex IP divergentes unifiees).
- **Core/ConsoleWindow.ps1** (nouveau) : P/Invoke console mutualise (`Hide-Console`/`Show-Console`), fin de la duplication Bootstrap/App.
- Libelles : ImportExcel n'est plus presente comme "optionnel" ; le message d'installation automatique annonce le telechargement et sa duree.

### Documentation
- **README** : structure du projet reelle (Bootstrap, Lang/, Version.ps1, Updater...), pretendu parallelisme des snapshots retire, Excel indique facultatif (fallback ImportExcel), sections langues / mise a jour automatique / SHA256 ajoutees, `config.json` documente en entier.

---

## v4.9.8
> Audit securite et robustesse — mises a jour verifiees, erreurs visibles, refactor

### Securite
- **Show-StartupCheck.ps1** : le depot de mise a jour est desormais **code en dur**. Toute valeur `autoUpdate.repo` differente dans `config.json` est rejetee — un fichier local altere ne peut plus rediriger les mises a jour vers un depot tiers (execution de code arbitraire).
- **Show-StartupCheck.ps1** : avant tout telechargement, l'URL de l'archive est validee (HTTPS, hote `github.com`, chemin du depot de confiance). Verification **SHA256** optionnelle si une ligne `sha256: <hash>` figure dans les notes de release.

### Corrections
- **Updater.ps1** : **sauvegarde du code avant ecrasement + restauration automatique** si la copie de la mise a jour echoue (plus de risque d'installation cassee).
- **Bootstrap.ps1 / App.ps1** : les erreurs fatales et l'echec de connexion au serveur s'affichent en `MessageBox`. La console etant masquee, un `Read-Host` restait invisible et l'app semblait se figer.
- **Initialize-Modules.ps1** : `Join-Path` a 3 arguments (casse en PS 5.1) corrige dans le chemin d'extraction nupkg a sous-dossier unique.

### Ameliorations
- **App.ps1** : le bouton Annuler est desormais rendu et cliquable avant le lancement du travail. Pump du journal UI limite a ~50 ms au lieu d'un rendu synchrone par ligne (actions verbeuses plus fluides).
- **src/Version.ps1** (nouveau) : version centralisee, dot-sourcee par Bootstrap et App — fin de la duplication du numero.
- **src/Core/RequiredModules.ps1** (nouveau) : liste des modules requis centralisee (`Get-RequiredModules`), utilisee par Initialize-Modules, Show-StartupCheck et Save-Dependencies.
- **Initialize-Modules.ps1** : suppression des branches "module optionnel" mortes (tous les modules sont requis).
- Retrait d'un fichier de log commite par erreur (fuite de chemin local).

---

## v4.9.7
> Troisieme audit qualite — annulation, erreurs cachees et traductions manquantes

### Corrections
- **Get-SnapshotAll.ps1** : le bloc `catch` utilisait `SA_LogFailed` qui n'affichait que le nom de la camera, sans le motif de l'echec. Remplace par `SA_LogError` (cle `"{0}": {1}` deja definie mais jamais utilisee) — l'erreur exacte est maintenant visible dans le journal.
- **fr.ps1 / en.ps1** : suppression de la cle `SA_LogFailed` devenue inutilisee apres le correctif ci-dessus.
- **Export-HardwareReport.ps1** : deux messages log etaient hardcodes en francais dans le chemin ImportExcel (`'Creation du fichier Excel via sous-processus...'` et `"ATTENTION: Impossible de supprimer l'ancien fichier Excel..."`). Remplaces par les nouvelles cles `EH_LogSubproc` et `EH_LogExcelLocked` (ajoutees dans fr.ps1 et en.ps1).
- **Export-HardwareReport.ps1** : les trois boucles de pre-capture du chemin ImportExcel (snapshot live, snapshot J-7, images de reference) n'appelaient jamais `$Cancel`. L'utilisateur pouvait cliquer "Annuler" sans effet pendant le telechargement des images. Ajout de `if (& $Cancel) { break }` en tete de chaque boucle — comportement coherent avec le chemin COM.
- **Get-PlaybackReport.ps1** : `Get-PlaybackInfo -Parallel` est un appel bloquant (30-60 s sur de grands systemes). Aucun message n'etait affiche pendant ce temps. Ajout d'un log apres le retour de l'appel (`PR_LogDataReceived`) pour signaler la fin de l'attente et le debut du traitement par camera. Cle ajoutee dans fr.ps1 et en.ps1.

---

## v4.9.6
> Deuxieme audit qualite — nettoyage de code mort, performance, UX et robustesse

### Corrections
- **fr.ps1** : trois apostrophes manquantes corrigees — `App_Closing` (`l'application`), `Act_Playback` (`Dates d'enregistrement`), `PR_LogFound` (`plages d'enregistrement`). Ces clés etaient differentes des corrections du tour precedent.
- **Updater.ps1** : le `catch {}` silencieux sur `Copy-Item` est remplace par un log dans `%TEMP%\MilestoneToolkitUpdate_error.txt`. Si la copie des fichiers de mise a jour echoue (fichier verrouille, droits insuffisants), l'erreur est maintenant tracee au lieu de passer inapercue.
- **Initialize-Modules.ps1 line 53** : `(Join-Path ...)\*` remplace par `(Join-Path ... '*')` — forme idiomatique et coherente avec la ligne 56 du meme fichier.

### Nettoyage de code mort
- **Bootstrap.ps1** : suppression des fonctions `Invoke-AutoUpdate`, `Get-GitHubLatestRelease` et `Get-ComparableVersion` (77 lignes) — jamais appelees depuis Bootstrap.ps1 ; la verification de mise a jour est assuree par Show-StartupCheck.
- **Show-StartupCheck.ps1** : suppression du scriptblock `$script:_SC_DownloadModuleNuGet` (54 lignes) — defini mais jamais reference nulle part dans le projet.

### Ameliorations
- **Bootstrap.ps1** — persistence de la langue : la langue choisie au selecteur est sauvegardee dans `config.json` (cle `language`). Aux prochains demarrages, le selecteur est ignore et la langue sauvegardee est utilisee directement.
- **Set-CameraGroupByModel.ps1** — complexite O(n×m) -> O(n+m) : `Get-VmsDeviceGroupMember` est desormais appele une seule fois par groupe (et non une fois par camera). Les IDs existants sont charges dans un `HashSet` avant la boucle — irrelevant sur les petits systemes, significatif sur les grandes installations.
- **Invoke-PtzPreset.ps1** — l'attente de 2 500 ms apres repositionnement PTZ ne bloque plus le thread UI. Remplace par une boucle de 50 ms pompant le `Dispatcher` WPF, ce qui maintient l'interface reactive pendant la stabilisation de la camera.
- **Save-Dependencies.ps1** : TLS 1.2 force en debut de script — requis par PowerShell Gallery sous PS 5.1.

---

## v4.9.5
> Correctifs qualite et robustesse — audit complet du code

### Corrections

- **Bootstrap.ps1** : ajout de `-TimeoutSec 10` sur `Invoke-RestMethod` pour eviter un blocage indefini si l'API GitHub ne repond pas.
- **Show-StartupCheck.ps1** : meme correctif sur `Invoke-RestMethod` pour la verification de mise a jour au demarrage.
- **Show-StartupCheck.ps1** : suppression d'une branche morte (`elseif ($allOk -and $anyUpdate ...)`) — la propriete `UpdateAvailable` n'etait jamais definie sur les objets module, rendant cette branche inatteignable.
- **Write-ActivityLog.ps1** : remplacement de `Add-Content -Encoding UTF8` (ecrit UTF-8 avec BOM sous PowerShell 5.1) par `[System.IO.File]::AppendAllText` avec `UTF8Encoding($false)` — les fichiers de log sont desormais UTF-8 sans BOM.
- **App.ps1** — logCallback : ajout de `.TrimStart()` avant le matching regex `^ERREUR` / `^AVERTISSEMENT`. Les messages indentes (ex : `  ERREUR: ...`) etaient incorrectement classes comme INFO et affiches en blanc au lieu du rouge/jaune attendu.
- **App.ps1** — Write-UILog : limitation du journal a 2 000 entrees. Au-dela, la plus ancienne ligne est supprimee pour eviter l'accumulation memoire sur les longues sessions.
- **Get-RecordingStats.ps1** : le bloc `catch {}` silencieux sur `Get-VideoDeviceStatistics` est remplace par un log AVERTISSEMENT (`RS_LogLiveWarn`) — coherent avec les blocs catch des stats enregistrement et mouvement.
- **Get-PtzPresetSnapshot.ps1** : sanitisation du nom de fichier des snapshots PTZ — les noms de camera et de preset contenant des caracteres invalides (`\ / : * ? " < > |`) sont maintenant remplaces par `_` avant construction du nom de fichier.
- **fr.ps1 / en.ps1** — `EH_LogNoExcel` : corrige le prefixe `ERREUR:` en `AVERTISSEMENT:` — Excel absent n'est pas une erreur fatale puisque le fallback ImportExcel suit immediatement.
- **fr.ps1 / en.ps1** : ajout de la cle `RS_LogLiveWarn` utilisee par le correctif Get-RecordingStats.

---

## v4.9.4

### Nouveautes
- Export Hardware : nouvelle option "Images de référence" dans le groupe Options. Permet de sélectionner un dossier sur le poste contenant une image par caméra (fichier nommé `NomCamera.ext` ou `NomCamera_<suffixe>.ext`, ex. snapshots horodatés — le plus récent est utilisé en cas de doublon). Ces images sont insérées dans une colonne dédiée du rapport Excel, selon le même fonctionnement que les colonnes Snapshot J-7 et Snapshot Live.


## v4.7.1
> Correctifs de compatibilite PSTools et export Excel sans Office

### Corrections
- Détection de `MilestonePSTools` améliorée : l'application prévient maintenant l'utilisateur si une version déjà installée est présente et indique si une mise à jour est disponible.
- L'outil inclut désormais `ImportExcel` comme dépendance de base.
- Export Hardware : si Excel n'est pas installé, l'application tente désormais d'utiliser le module PowerShell `ImportExcel` pour produire le fichier `.xlsx`.
- Export sans Excel : les snapshots ne sont plus inclus dans l'export en fallback `ImportExcel`, afin d'éviter des erreurs de génération.

### Ameliorations
- Ajout d'une vérification de mise à jour GitHub au démarrage via `config.json`.
- `ImportExcel` est pris en charge comme module optionnel et installé automatiquement si nécessaire.
- Messages de fallback et d'erreur clarifiés dans l'UI et le journal.

## v4.7
- `ImportExcel` est pris en charge comme module optionnel et installé automatiquement si nécessaire.
- Messages de fallback et d'erreur clarifiés dans l'UI et le journal.

## v4.7
> Correctifs telechargement offline, robustesse globale et audit qualite

### Corrections — telechargement offline
- **Preparer offline** : `Save-Module` remplace par un telechargement direct via `WebClient` + `Expand-Archive`. Aucune dependance sur PowerShellGet ou NuGet provider — fonctionne sur toute installation fraiche de Windows.
- **TLS 1.2** force des le demarrage de Bootstrap.ps1 (avant tout chargement de module).
- Nom de fichier temp unique par run (GUID) pour eviter les conflits d'extraction entre deux tentatives consecutives.

### Corrections — robustesse
- **Fuite de ressources** : `WebClient` desormais dispose dans un bloc `try/finally`.
- **Fuite de threads** : pool de runspaces (snapshots paralleles) ferme dans un `try/finally` dans `Get-SnapshotAll` et `Export-HardwareReport`.
- **Timeout snapshots** : les operations de capture en parallele s'interrompent apres 10 minutes si une camera ne repond pas, au lieu de bloquer indefiniment.
- **Get-SnapshotDateTime** : retourne desormais `@{ Ok; Time }` au lieu de trois types differents (`$null`, `$false`, `[datetime]`) — comparaisons fiables dans tous les cas.
- **Export CSV** : `Export-Csv` entoure d'un `try/catch` dans Get-CameraStatus, Get-PlaybackReport et Get-RecordingStats — echec visible plutot que silencieux.
- **Set-CameraGroupByModel** : `New-VmsDeviceGroup` et `Add-VmsDeviceGroupMember` dans un `try/catch` — erreur par modele loggee sans interrompre les autres.
- **Get-LicenseInfo** : proprietes `CarePlus`/`CarePremium` utilisent desormais `$script:T` pour respecter la langue selectionnee.

### Ameliorations
- **Version centralisee** : `$script:AppVersion` defini une seule fois dans `Bootstrap.ps1` et `App.ps1`, reference partout via `$ver` dans les fichiers de langue — une seule ligne a modifier pour changer de version.
- **Nouvelles cles de traduction** : `SA_LogTimeout`, `EH_LogSnapTimeout`, `GM_LogModelError`, `LI_LogCareProp`.

---

## v4.5.1
> Packaging Inno Setup — installeur Windows natif

### Nouveautes
- **Installeur EXE** : Inno Setup compile deux variantes publiees automatiquement sur chaque release GitHub :
  - `*-Online-Setup.exe` (~500 Ko) : leger, telecharge MilestonePSTools au premier lancement.
  - `*-Offline-Setup.exe` (~80 Mo) : module MilestonePSTools bundle dans l'EXE, fonctionne sans Internet.
- **Raccourci Menu Demarrer** et optionnellement Bureau. Lance l'application sans aucune fenetre console (via `Start.vbs`).
- **Pas d'UAC requis** : installation dans `AppData\Local` par defaut. Option "Pour tous les utilisateurs" disponible pour les admins.
- **Deblocage automatique** : `Unblock-File` applique sur tous les scripts lors de l'installation — le flag Zone.Identifier n'est plus un probleme.
- **Mise a jour silencieuse** : reinstaller un EXE plus recent met a jour l'installation existante (meme AppId Inno Setup).
- **Desinstallation propre** : entree dans Programmes & fonctionnalites, suppression des dossiers `Output/`, `Logs/`, `Dependencies/`.
- **GitHub Actions** (`.github/workflows/release.yml`) : push d'un tag `vX.X` → build automatique des deux EXE → publication sur la page Releases.
- **Build local** : `installer\build.ps1 -Mode Online` ou `-Mode Offline` pour compiler sans passer par CI.

---

## v4.5
> Support multilingue FR / EN

### Nouveautes
- **Selecteur de langue au demarrage** : une fenetre de choix FR / EN s'affiche avant la verification des dependances. La langue selectionnee s'applique a toute l'application pour la session.
- **Interface entierement traduite** : tous les textes de l'UI sont traduits — fenetre de demarrage, fenetre principale (categories, boutons, statuts, combo mode capture), fenetre de selection des colonnes Export Hardware.
- **Excel traduit** : les en-tetes de colonnes du fichier Excel suivent la langue choisie. Le nom du fichier change egalement (`Liste_des_Cameras.xlsx` en francais, `Camera_List.xlsx` en anglais).
- **CSV traduits** : les noms de colonnes et les noms de fichiers CSV s'adaptent a la langue (`Etat_Cameras.csv` / `Camera_Status.csv`, `Dates_Enregistrement.csv` / `Recording_Dates.csv`, `Stats_Enregistrement.csv` / `Recording_Stats.csv`).
- **Logs traduits** : tous les messages de journal d'activite (progression, erreurs, avertissements) sont emis dans la langue selectionnee.
- **Architecture i18n** : fichiers `src/Lang/fr.ps1` et `src/Lang/en.ps1` centralisant ~200 cles de traduction. Extensible a d'autres langues sans modifier le code applicatif.

---

## v4.4
> Export Hardware — selection des colonnes et securite mot de passe

### Ameliorations
- **Export Hardware — Fenetre de selection des colonnes** : remplacement des dialogues MessageBox par une fenetre WPF avec cases a cocher groupees par categorie (Informations hardware, Flux video, Retention, Options). Seules les colonnes cochees apparaissent dans le fichier Excel.
- **Export Hardware — Mots de passe exclus par defaut** : la colonne `Mot de passe` est desormais decochee par defaut. Elle doit etre activee explicitement. L'appel `Get-VmsCameraReport -IncludePlainTextPassword` n'est effectue que si la colonne est selectionnee.
- **Export Hardware — Optimisation des appels API** : `Get-VmsCameraStream` et `Get-PlaybackInfo` ne sont appeles que si les colonnes flux video ou retention sont selectionnees, evitant des requetes inutiles.
- **Export Hardware — Boutons Tout cocher / Tout decocher** : selection rapide de toutes les colonnes ou remise a zero en un clic.
- **Export Hardware — Suppression de l'avertissement mot de passe** : la fenetre de confirmation redondante est supprimee, la case `Mot de passe (!)` marquee en orange suffit comme signal d'intention.

---

## v4.3
> Enrichissement de l'export Hardware

### Ameliorations
- **Export Hardware — Flux video** : ajout de 7 nouvelles colonnes groupees par couleur dans le fichier Excel : `Codec (Enreg.)`, `Resolution (Enreg.)`, `FPS (Enreg.)`, `Codec (Live)`, `Resolution (Live)`, `FPS (Live)`, `Flux supplementaires`. Si le flux live et le flux enregistre sont identiques, les colonnes Live restent vides pour eviter la redondance.
- **Export Hardware — Retention disponible** : nouvelle colonne `Retention disponible` indiquant la duree totale d'archive accessible par camera, calculee via `Get-PlaybackInfo`.
- **Export Hardware — Adresse IP** : suppression automatique du protocole (`http://`) et du port (`:8000`) — seule l'adresse IP pure (4 blocs) est affichee.
- **En-tetes Excel colores par groupe** : hardware (sable), flux video (bleu marine), retention (vert), snapshot (violet) pour une lecture rapide.

---

## v4.2
> Nouvelle categorie Monitoring

### Nouveautes
- **Etat des cameras** (`Get-ItemState`) : etat temps reel de chaque camera via l'Event Server. Les cameras Responding sont affichees en vert, les erreurs en rouge, les cameras desactivees sont automatiquement ignorees. Export CSV `Etat_Cameras.csv`.
- **Dates d'enregistrement** (`Get-PlaybackInfo`) : premier et dernier enregistrement disponible par camera avec la duree de retention calculee. Utilise le mode parallele natif du SDK pour les grands systemes. Export CSV `Dates_Enregistrement.csv`.
- **Nouvelle categorie MONITORING** dans la sidebar (couleur cyan) separee de DIAGNOSTIC pour distinguer les donnees temps reel des analyses historiques.

---

## v4.1
> Optimisation des performances

### Ameliorations
- **Snapshot - Toutes les cameras** : les snapshots sont desormais recuperes en parallele via un pool de threads (jusqu'a 12 simultanees). Le temps de capture est divise par le nombre de cameras actives simultanement au lieu d'etre sequentiel.
- **Export Hardware - Snapshots** : meme amelioration — les snapshots sont pre-telecharges en parallele avant la construction du fichier Excel. Le log affiche chaque snapshot des qu'il est recu, dans l'ordre d'arrivee.
- **Configuration** : suppression du parametre `installMode` dans `config.json` — le mode Online/Offline est toujours auto-detecte (presence du dossier `Dependencies/`).

---

## v4.0
> Refonte de l'export hardware + nouvelles fonctions de diagnostic + mode capture historique

### Nouveautes

#### Export Hardware
- Format de sortie change de **CSV vers Excel (.xlsx)** avec mise en forme (en-tete colore, bordures, colonne figee)
- Option d'inclure un **snapshot integre dans la cellule** de chaque camera (ancre, redimensionne automatiquement)
- Suppression des colonnes GPS et Activation (inutiles)
- Demande de confirmation avant l'inclusion des snapshots (operation longue)

#### Mode capture historique
- Nouveau selecteur **Live / Image historique** dans la sidebar
- En mode historique : choix de la date et de l'heure (hh:mm)
- Applicable aux trois actions snapshot : Selection, Toutes les cameras, Presets PTZ
- Utilise `Get-Snapshot -Behavior GetNearest` pour trouver l'image la plus proche

#### Diagnostic (nouvelle categorie)
- **Stats Enregistrement (7j)** : sequences, pourcentage de temps enregistre, duree totale, detection de mouvement et statistiques live (FPS, bitrate, resolution) par camera. Export CSV.
- **Informations Licence** : produits licencies, date d'expiration avec alerte si < 30 jours, canaux utilises/total avec alerte si >= 90 %. Utilise `Get-VmsLicensedProducts`.

### Corrections
- Rapport Stockage supprime (les donnees d'espace disque ne sont pas accessibles sans droits WMI sur le serveur distant)
- `Get-VmsLicensedProducts` : gestion correcte des valeurs d'expiration non-date (`"Unrestricted"`)
- Proprietes internes du SDK Milestone filtrees de l'affichage licence (Path, ParentPath, ServerId, etc.)

---

## v3.0
> Fork initial — refonte UX complete

### Nouveautes
- **Journal d'activite colore** : remplacement du TextBox monochrome par un RichTextBox avec codes couleur par niveau (erreur, avertissement, succes, action)
- **Bouton Annuler** : toutes les actions longues peuvent etre interrompues entre chaque iteration
- **Barre de progression** : deterministe (avancement reel) pour les actions iteratives
- **Selecteur de dossier de sortie** : changement du repertoire sans modifier `config.json`
- **Snapshot - Presets PTZ** : progression basee sur le nombre total de presets toutes cameras confondues
- **Grouper par Modele** : verification des doublons avant ajout au groupe

### Corrections
- Encodage du titre de fenetre corrige (`—` affiché correctement via entite XML `&#x2014;`)
- Banniere de demarrage : distinction correcte entre mode hors ligne (cache local) et connexion Internet disponible
- ComboBox mode capture : template sombre complet pour etre lisible sur fond sombre
- `Get-CameraRecordingStats` : parametres corriges (`-StartTime`/`-EndTime`, proprietes `PercentRecorded`/`TimeRecorded`)
- `Get-VmsStorageRetention` : retourne un `[TimeSpan]` directement, `.TotalDays` utilise
- Confirmation avant inclusion des mots de passe en clair dans l'export hardware
