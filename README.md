# RDV & Notes — complément Outlook 2.3.3

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
