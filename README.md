# RFR Notes CRM — phase 2

Ce projet remplace la page Microsoft « Say hello » par un formulaire mobile de
compte rendu. Il utilise le compte Microsoft déjà ouvert dans Outlook et crée
une **Note Dataverse (`annotation`)** reliée au rendez-vous CRM (`appointment`).

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
5. Les utilisateurs doivent pouvoir lire les rendez-vous et créer des notes
   dans Dataverse.

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
