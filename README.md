# RFR Notes CRM — phase 2

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
3. Une URI de redirection SPA NAA : `brk-multihub://DOMAINE-DE-L-APPLICATION`.
4. Un hébergement HTTPS accessible depuis le XCover.
5. Les utilisateurs doivent pouvoir lire et modifier les rendez-vous, lire les
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

## Champs CRM

À partir de la version 2.1.0, le panneau affiche la qualification du rendez-vous
CRM sélectionné et permet de la modifier. Les choix ne sont pas codés en dur :
ils sont lus dans les métadonnées Dataverse du champ
`appointment.rfr_os_appointmentqualification`, puis la valeur est enregistrée
sur le rendez-vous avec les droits de l’utilisateur connecté.

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
