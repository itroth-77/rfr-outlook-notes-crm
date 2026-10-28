# RDV & Notes — complément Outlook 2.3.9

## Ajustement 2.3.9 — un seul bouton mobile clair

L'utilisateur a confirmé l'ouverture réelle de l'APK depuis le lien HTTPS de la
2.3.8. Ce lien et sa navigation au clic sont conservés sans modification.
Le bouton affiche maintenant le logo de l'application avec « Ouvrir l’application
mobile » et « RDV & Notes » sur deux lignes. L'aide « Si Outlook bloque… », le
bouton Copier, le lien brut et le code de presse-papiers sont supprimés. Aucun
message n'apparaît après un clic normal ; les protections contre le changement
de RDV et les brouillons non enregistrés restent actives. 98 tests simulés,
compilation et manifeste validés. Aucun changement d'APK ni d'Hyperion.

Le libellé natif « Journal » d'Outlook n'est pas modifié : le manifeste déclare
déjà RDV & Notes, mais la surface MobileLogEventAppointmentAttendee expose
l'action standard Log/View gérée par Outlook. EventLogged n'est activé qu'après
la création effective d'une note, pas pour forcer un libellé.

## Historique 2.3.8

## Ajustement 2.3.8 — lien mobile HTTPS et copie sans volet

Le test réel du 28 octobre a confirmé : le lien HTTPS copié ouvre l'APK, mais
la tentative `intent://` depuis le complément ne lance rien. Cette tentative
est supprimée. « Ouvrir dans RDV & Notes » pointe maintenant exactement vers le
même HTTPS que « Copier le lien », sous forme d'une vraie ancre HTML. Pas de
redirection programmée, pas de popup, pas de schéma personnalisé. Le contrôle du
RDV et du brouillon n'intercepte que les clics dangereux, pas le clic normal.

L'aide dépliable et le champ non cliquable sont supprimés. Le bouton Copier est
visible directement. Si le presse-papiers est refusé, un véritable lien HTTPS
cliquable est affiché et peut être copié par appui long. Aucun accès automatique
au presse-papiers, changement Intune, APK ou Hyperion.

**Ouverture directe non encore validée sur Outlook Android** : le passage du
WebView Outlook vers le système reste contrôlé par Outlook. Le complément ne
dispose pas d'OpenBrowserWindowApi sur Android et ne peut pas garantir la sortie.
La suppression de cr-api dans AppDomains ne garantit pas non plus ce comportement
sur mobile (la documentation décrit cette navigation pour les plateformes bureau).
Retester le clic, puis Copier en secours, sur le même téléphone. Le résultat de
99 tests simulés ne valide pas le lancement d'une application native réelle.

Référence Microsoft :
https://learn.microsoft.com/en-us/office/dev/add-ins/reference/requirement-sets/open-browser-window-api-requirement-sets

## Historique 2.3.7

## Ajustements 2.3.6 / 2.3.7 — information de série sans doublon

Le volet « Portée et exceptions de la série » est supprimé, sans remplacement
par un autre volet dépliable ni par une autre phrase dans Dynamics : l'information
figure déjà dans la fiche du rendez-vous Outlook, affichée une seule fois.
Les avertissements bloquants restent visibles et les précautions sur les séances
personnalisées restent dans la confirmation avant modification. Aucun changement
de logique Dynamics/Graph, de permission, d'APK ou d'Hyperion. 96 tests simulés.

## Historique 2.3.5

## Correctif 2.3.5 — Outlook en mode modification (28 octobre 2026)

La capture de l'occurrence IT Meeting a identifié un manque dans la 2.3.4 :
`item.itemId` appartient au mode lecture, pas à `AppointmentCompose`. Le
complément utilise maintenant `getItemIdAsync` en mode modification pour lire
l'identifiant de l'élément déjà enregistré, puis le convertir au format REST.
Il ne sauvegarde jamais automatiquement l'élément Outlook avec `saveAsync` :
les modifications Outlook en attente, les horaires, les invités et la récurrence
restent inchangés. Aucun changement de permission, de GUID, d'APK ou d'Hyperion.

Les contrôles asynchrones relisent cet identifiant après les chargements et
immédiatement avant les PATCH Graph/Dataverse et le POST de note. Le même objet
Office réutilisé pour un autre élément ne suffit donc pas à autoriser l'écriture.
Un élément non enregistré ou un ancien client sans `getItemIdAsync` est bloqué
avec un message utile, sans sauvegarde ni appariement de secours par titre.

94 tests simulés réussis, dont lecture/invité/mobile, occurrence en modification,
conversion EWS/REST, changement pendant confirmation, même wrapper Office avec
autre ID, refus asynchrone avant écriture, élément non enregistré et ancien
client. Compilation Vite et manifeste 2.3.5.0 validés. Publication autorisée par
l'utilisateur. Reste à retester le RDV réel montré sur la capture dans Outlook.

Référence API Microsoft :
https://learn.microsoft.com/en-us/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-getitemidasync-member(1)

## Historique 2.3.4

## Mise à jour du 28 octobre 2026

Publication 2.3.4 autorisée par l'utilisateur. Manifeste 2.3.4.0, même GUID,
mêmes permissions et logo. Uniquement le rendez-vous ouvert dans Outlook :
pas d'agenda par semaine, ni de fonctions audio/transcription/génération de CR.

- Actualiser relit le contexte Office, puis Graph et Dynamics. Plus de boucle
  demandant de fermer/rouvrir le complément. Quand ItemChanged est disponible,
  un changement réel recharge le RDV hors opération active ; sinon Actualiser
  reste disponible. Un retour de focus contrôle également l'identité.
- Les réponses tardives ne peuvent pas sélectionner l'ancien rendez-vous ni
  modifier le nouveau. L'identité est revérifiée avant les PATCH/POST. Les
  brouillons de notes sont gardés uniquement en mémoire, par compte et séance,
  jusqu'à fermeture du volet. Pas de sauvegarde persistante de ces textes.
- En-tête fixe avec Actualiser ; version dans À propos ; sujet, horaires et
  récurrence dans la fiche RDV. Pas d'activation quand le RDV Dynamics est trouvé.
  Catégorie déjà active mais CRM absent : Vérifier sans refaire le suivi.
- Qualification, Concernant, note manuelle et notes/CR de cette séance. Éditeur
  Concernant et note dépliables, historique chargé seulement à la demande, texte
  de chaque note dépliable. Les règles sûres de série et de notes restent celles
  décrites ci-dessous.
- Ouverture Android : lien intent explicite vers `fr.roth.compterendu`, portant
  la même route HTTPS `/open?eventId=...&view=crm` reconnue par l'APK. Aucun jeton,
  contenu CRM, changement d'APK ou d'Hyperion. Pas de page intermédiaire dans ce
  parcours. Le domaine cr-api n'est plus déclaré comme navigation interne du
  manifeste : réinstaller le XML pour appliquer ce changement.

**Limite mobile à tester, non résoluble de façon garantie par le complément :**
Outlook Android ne propose pas l'API OpenBrowserWindowApi. Certains WebViews
refusent aussi les intents externes ; l'essai sur le téléphone reste nécessaire.
Le lien HTTPS peut être copié depuis l'aide et ouvert dans le navigateur du
téléphone. Les liens reçus par mail restent inchangés. Aucune redirection
automatique différée, aucun schéma natif injecté dans une page Hyperion, aucune
modification des règles de sécurité Outlook/Intune.

Contrôles : 84 tests simulés, compilation Vite et validation du manifeste Office.
Aucun rendez-vous réel ni appel IA pour les tests. Tester le changement A/B/A,
le brouillon, l'organisateur et l'invité sur série/exception, les refus de droits,
la note/les CR de la séance et l'ouverture Android sur la flotte pilote.

Références officielles :
- https://learn.microsoft.com/en-us/javascript/api/outlook/office.mailbox#outlook-office-mailbox-addhandlerasync-member(1)
- https://learn.microsoft.com/en-us/office/dev/add-ins/reference/requirement-sets/open-browser-window-api-requirement-sets
- https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests#specify-domains-you-want-to-open-in-the-add-in-window
- https://developer.chrome.com/docs/android/intents

## Historique 2.3.3

Préparation locale du 27 octobre 2026 ; publication autorisée le 28 octobre 2026.
L'ancien nom RFR Notes CRM est
remplacé par RDV & Notes sans changer le GUID du complément ni l'authentification.

Le manifeste 2.3.3.1 reprend le logo microphone/Roth/Dynamics fourni le 28 octobre.
Le dessin complet est conservé ; seuls les PNG sont redimensionnés pour les
tailles bureau/mobile et haute densité. Le source est dans
`assets/rdv-notes-logo-source.png` ; `scripts/prepare-icons.py` (Pillow, local)
permet de reproduire les tailles. Les anciennes URL d'icônes restent disponibles
pour les manifestes déjà installés. Charger le nouveau XML pour changer ce logo.

Périmètre : suivi Dynamics, Qualification, Concernant, note manuelle simple,
historique et lien vers le RDV dans l'APK. Aucun enregistrement audio,
transcription, génération ou rattachement de CR dans le complément.

L'historique affiche les 20 dernières notes identifiées pour la seule séance
sélectionnée, pas toutes celles du Concernant. Les nouvelles notes gardent leur
rattachement au Concernant et portent une référence `RFR-RDV` fondée sur le GUID
Dataverse de la séance, commune à l'APK 0.2.9. Ce n'est pas le GUID du maître.
Les notes directement rattachées au RDV sont également affichées. Les anciennes
notes sans cette référence restent accessibles dans Dynamics : aucun appariement
par sujet/date, déplacement ou migration massive n'est effectué.
Référence API : https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query/filter-rows

Le manifeste ne met pas lui-même le JavaScript à jour : son URL est celle du site
partagé existant. Le déploiement GitHub Pages du site doit réussir avant de
tester les nouveautés dans Outlook. Cette publication concerne les utilisateurs
actuels du complément ; aucun déploiement de l'APK ou d'Hyperion n'en découle.

Les RDV récurrents sont résolus sur le maître de la copie Outlook de l'utilisateur
et sur le maître Dynamics via seriesid, jamais par réutilisation d'un identifiant
Exchange comme GUID CRM. Catégories existantes, exceptions, droits et ETag sont
préservés. Confirmation explicite « toute la série », sans repli silencieux sur
une occurrence. Les modifications du Concernant et de Qualification sont
recontrôlées côté serveur ; aucune adaptation du schéma de production.

Sur bureau/web, deux surfaces sont déclarées : organisateur et invité. Sur mobile,
le point d'extension Microsoft MobileLogEventAppointmentAttendee reste utilisé ;
la visibilité réelle selon Outlook et le type de RDV doit être testée. Les tests
simulés ne garantissent pas que Microsoft affichera le bouton dans toute vue.

Références :
- https://learn.microsoft.com/en-us/javascript/api/manifest/extensionpoint
- https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/mobile-log-appointments
- https://learn.microsoft.com/en-us/graph/api/resources/event
- https://learn.microsoft.com/en-us/power-apps/developer/data-platform/update-recurring-appointment

Ce projet remplace la page Microsoft « Say hello » par un formulaire mobile de
compte rendu. Il utilise le compte Microsoft déjà ouvert dans Outlook et crée
une **Note Dataverse (`annotation`)** reliée au dossier indiqué par le champ
**Concernant** du rendez-vous CRM (compte, contact, opportunité, etc.). Si le
rendez-vous n’a aucun Concernant, la saisie et l’enregistrement sont désactivés.

## Sécurité

- aucun mot de passe ni secret applicatif dans le code ;
- aucun compte technique ou compte `adminsynchro` ;
- authentification SSO par Nested App Authentication (NAA) ;
- opérations Dataverse exécutées avec les droits du collaborateur connecté ;
- aucune liaison automatique si le rendez-vous CRM est ambigu.

## Configuration requise

1. Une inscription d’application **SPA** dans Microsoft Entra.
2. L’autorisation déléguée Dynamics CRM `user_impersonation`.
3. L’autorisation déléguée Microsoft Graph `Calendars.ReadWrite`, nécessaire
   pour activer le suivi d’un rendez-vous Outlook encore absent du CRM.
4. Une URI de redirection SPA NAA : `brk-multihub://DOMAINE-DE-L-APPLICATION`.
5. Un hébergement HTTPS accessible depuis le XCover.
6. Les utilisateurs doivent pouvoir lire et modifier les rendez-vous, lire les
   métadonnées et créer des notes sur les dossiers Concernant dans Dataverse.

## Préparer les paramètres

Copier `.env.example` vers `.env`, puis remplacer :

- `VITE_CLIENT_ID` : ID de l’application Entra ;
- `VITE_TENANT_ID` : ID du tenant Roth France ;
- `VITE_DATAVERSE_URL` : URL racine de l’environnement Dynamics ;
- `APP_URL` : URL HTTPS publique de ce formulaire.

Ces quatre valeurs sont des identifiants de configuration publics. Aucun secret
client ne doit être créé ni copié dans ce projet.

## Hébergement gratuit avec GitHub Pages

Pour le prototype, le projet peut être publié gratuitement sur GitHub Pages.
Le dépôt ne contient ni secret ni donnée CRM. Les appels CRM sont faits depuis
le téléphone directement vers Microsoft Dataverse avec le jeton de l’utilisateur.

1. Créer le dépôt GitHub public `itroth-77/rfr-outlook-notes-crm`.
2. Placer le contenu de ce dossier à la racine du dépôt.
3. Dans **Settings > Pages**, sélectionner **GitHub Actions** comme source.
4. Dans **Settings > Secrets and variables > Actions > Variables**, créer :
   - `VITE_CLIENT_ID` ;
   - `VITE_TENANT_ID`.

   Le workflow contient déjà les deux valeurs publiques de l’application :
   `https://roth.crm4.dynamics.com` et
   `https://itroth-77.github.io/rfr-outlook-notes-crm`.
5. Envoyer les fichiers sur la branche `main`. Le workflow
   `Publier RFR Notes CRM` teste, compile et publie automatiquement le site.
6. Le manifeste final devient téléchargeable à l’adresse
   `APP_URL/manifest.xml`.

Pour NAA, l’inscription Entra doit contenir l’URI de redirection SPA
`brk-multihub://itroth-77.github.io`. Le chemin du dépôt ne doit pas être
ajouté à cette URI : Microsoft demande uniquement l’origine du site.

## Construire localement

```powershell
npm install
npm test
npm run build
```

Publier le contenu de `dist` sur l’hébergement HTTPS, puis générer et valider le
manifeste :

```powershell
npm run manifest
npm run validate:manifest
```

Supprimer le complément de phase 1, charger le nouveau `manifest.xml`, puis
fermer complètement et relancer Outlook Mobile.

## Correspondance avec le rendez-vous CRM

Le formulaire recherche les rendez-vous Dataverse situés à plus ou moins
15 minutes de l’heure Outlook. Il compare ensuite le sujet, le début et la fin.
Une correspondance forte et unique est présélectionnée. Dans les autres cas,
l’utilisateur doit choisir explicitement le rendez-vous CRM avant que le bouton
d’enregistrement soit activé.

Ce choix évite d’attacher silencieusement une note au mauvais client.

Depuis la version 2.1.1, la liste est également limitée aux rendez-vous pour
lesquels l’utilisateur Dataverse connecté est organisateur, participant
obligatoire ou participant facultatif. Les rendez-vous des autres utilisateurs
ne sont ni proposés ni présélectionnés.

À partir de la version 2.2.0, aucun choix manuel entre plusieurs rendez-vous
n’est proposé. Le sujet et les horaires doivent fournir une correspondance
forte et unique. Lorsqu’aucune correspondance n’existe, le panneau peut ajouter
la catégorie Outlook `Tracked To Dynamics 365` avec Microsoft Graph, puis
recherche automatiquement le rendez-vous dans Dataverse pendant deux minutes.
La synchronisation serveur Dynamics restant asynchrone, elle peut prendre
jusqu’à environ quinze minutes ; le panneau peut alors être rouvert plus tard.

## Champs CRM

À partir de la version 2.1.0, le panneau affiche la qualification du rendez-vous
CRM sélectionné et permet de la modifier. Les choix ne sont pas codés en dur :
ils sont lus dans les métadonnées Dataverse du champ
`appointment.rfr_os_appointmentqualification`, puis la valeur est enregistrée
sur le rendez-vous avec les droits de l’utilisateur connecté.

Le Concernant peut être affecté ou remplacé depuis le panneau après une
recherche parmi les comptes, contacts ou opportunités actifs. La création de
note reste désactivée tant qu’aucun Concernant n’est enregistré.

La recherche du Concernant accepte également les jokers `*` et `%`. Chaque
fragment saisi doit être présent dans les champs recherchables. Pour un contact,
la recherche couvre le nom, le prénom et l’adresse e-mail. Pour une opportunité,
elle couvre son numéro et le nom de l’Affaire liée ; les résultats sont affichés
sous la forme `403151 — MA REXEL`. Par exemple, `%EDEO%GNY%` retrouve
`CEDEO LAGNY` pour un compte.

L’ordre des fragments est respecté dans le texte affiché. Ainsi,
`*ma*26*rexel*` exige que `ma` apparaisse avant `26`, puis `rexel`, dans
`numéro — Affaire`.

Les notes de chronologie enregistrées en HTML par Dynamics sont rendues comme
du contenu mis en forme après suppression des balises, attributs et URL
potentiellement dangereux.

## Diagnostic mobile

À partir de la version 2.0.4, le complément affiche l’étape précise de son
démarrage et retente une fois la lecture du rendez-vous si Outlook tarde à
répondre. Un échec affiche un code technique et un bouton **Réessayer**.

La version 2.0.5 charge également la bibliothèque Office après l’affichage de
l’interface. Une indisponibilité d’Office.js ne peut ainsi plus bloquer le volet
avant même le démarrage du diagnostic.

La version 2.0.6 n’attend plus obligatoirement la résolution de
`Office.onReady()`. Lorsque le rendez-vous est déjà exposé dans
`Office.context.mailbox.item`, il est lu directement. Dans le cas contraire,
le complément attend en parallèle la fin de l’initialisation ou l’apparition du
rendez-vous.

Si Outlook ferme le volet avant de pouvoir afficher l’erreur, la dernière étape
atteinte est conservée uniquement dans le stockage local du volet. Elle est
présentée à la prochaine ouverture réussie. Le diagnostic ne contient ni sujet,
ni participant, ni texte de note.

Le cas de référence ayant motivé ce diagnostic a été observé avec Outlook pour
Android 5.2640.
