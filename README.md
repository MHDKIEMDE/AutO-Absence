# Système de Gestion des Autorisations d'Absence

Ce script Google Apps Script permet d'automatiser le processus de demande et de validation des autorisations d'absence au sein d'une organisation.

## Fonctionnalités

- **Traitement des soumissions de formulaire** pour les demandes d'absence
- **Flux de validation à plusieurs niveaux** (Supérieur Hiérarchique → RH → Présidence)
- **Génération automatique de documents** basés sur un modèle
- **Organisation des fichiers** dans Google Drive (par année, par employé, et par statut)
- **Notifications par email** à chaque étape du processus
- **Protection des décisions** une fois prises
- **Traçabilité complète** avec journalisation des actions

## Prérequis

- Un formulaire Google Forms pour la soumission des demandes d'absence
- Une feuille de calcul Google Sheets liée au formulaire
- Un document modèle Google Docs avec des placeholders
- Les dossiers Google Drive suivants:
  - Dossier principal (ID: `1H5GtauuqauEf_Ze4ZQ3kkoNEgyDJLsRw`)
  - Dossier des employés (ID: `1D3xmuzg4jN9kfr_pfBly1E2fuaEO2-nO`)
  - Document modèle (ID: `1b_RRz7ie06UHuhfB4_5Sqeh33yhAD8gWeYVaAjrnz94`)

## Configuration

Avant utilisation, vous devez configurer:

1. Les adresses email pour les notifications:
   ```javascript
   var emailRH = "@gmail.com"; // À remplacer par l'email RH réel
   var emailPresidence = "@gmail.com"; // À remplacer par l'email de la présidence
   ```

2. Vérifier que les IDs des dossiers correspondent à votre structure Drive:
   ```javascript
   var mainFolderId = "1H5GtauuqauEf_Ze4ZQ3kkoNEgyDJLsRw"; // ID du dossier principal
   var employeeFolderId = "1D3xmuzg4jN9kfr_pfBly1E2fuaEO2-nO"; // ID du dossier des employés
   var templateDocId = "1b_RRz7ie06UHuhfB4_5Sqeh33yhAD8gWeYVaAjrnz94"; // ID du modèle de document
   ```

## Structure du document modèle

Le document modèle doit contenir les placeholders suivants:
- `{{matricule}}`
- `{{nom}}`
- `{{prenom}}`
- `{{serviceFonction}}`
- `{{typeAbsence}}`
- `{{dateDebut}}`
- `{{heureDebut}}`
- `{{dateFin}}`
- `{{heureFin}}`
- `{{motifAbsence}}`
- `{{typePermission}}`
- `{{motifPermission}}`
- `{{nbreJours}}`
- `{{statusSuperieurHierachique}}`
- `{{statusRH}}`
- `{{statusPresidence}}`
- `{{commentairePresidence}}`

## Structure de la feuille de calcul

La feuille de calcul doit avoir la structure suivante:
1. Les colonnes 1-16 contiennent les données du formulaire
2. Colonne 17: Validation du Supérieur Hiérarchique
3. Colonne 18: Validation RH
4. Colonne 19: Validation Présidence
5. Colonne 20: Commentaire Présidence
6. Colonne 21: ID du document généré
7. Colonne 22: Email RH
8. Colonne 23: Email Présidence

## Organisation des dossiers

Le script crée automatiquement une structure de dossiers organisée:
```
Dossier principal
└── Dossier des employés
    └── Autorisation d'absence - [ANNÉE]
        └── [NOM PRÉNOM]
            ├── En Attente/
            ├── Validée/
            └── Rejetée/
```

## Déploiement

1. Ouvrez votre feuille de calcul Google Sheets liée au formulaire
2. Allez dans Extensions > Apps Script
3. Copiez et collez l'intégralité du code dans l'éditeur de script
4. Sauvegardez le projet
5. Configurez les déclencheurs:
   - `onFormSubmit`: à déclencher lors de la soumission du formulaire
   - `onEditValidation`: à déclencher lors de la modification de la feuille

### Configuration des déclencheurs

Pour configurer les déclencheurs:
1. Dans l'éditeur Apps Script, cliquez sur l'icône en forme d'horloge (⏰)
2. Cliquez sur "+ Ajouter un déclencheur"
3. Pour `onFormSubmit`:
   - Sélectionnez "onFormSubmit" comme fonction à exécuter
   - "Depuis la feuille de calcul" comme source d'événement
   - "Lors de la soumission du formulaire" comme type d'événement
4. Pour `onEditValidation`:
   - Sélectionnez "onEditValidation" comme fonction à exécuter
   - "Depuis la feuille de calcul" comme source d'événement
   - "Lors de la modification" comme type d'événement
5. Cliquez sur "Enregistrer" pour chaque déclencheur

## Processus de validation

1. Un employé soumet une demande d'absence via le formulaire
2. Le script crée un document basé sur le modèle et l'enregistre dans le dossier "En Attente"
3. Le supérieur hiérarchique est notifié par email
4. Lorsque le supérieur valide, le RH est notifié
5. Lorsque le RH valide, la Présidence est notifiée
6. Lorsque la Présidence valide, le demandeur est notifié de l'approbation finale
7. Si un validateur rejette la demande, le processus s'arrête et le demandeur est notifié du rejet

## Fonctions principales

- `onFormSubmit(e)`: Traite les nouvelles soumissions de formulaire
- `onEditValidation(e)`: Gère les validations dans la feuille de calcul
- `sendEmail(recipient, subject, body)`: Envoie des notifications par email
- `handleValidation(role, statut, docId)`: Met à jour le statut dans le document
- `notifyNextValidator(sheet, row, colonne, docId)`: Notifie le prochain validateur dans la chaîne
- `handleRejection(sheet, row, role, docId)`: Gère le rejet d'une demande
- `moveDocument(doc, sourceFolder, targetFolder)`: Déplace les documents entre les dossiers
- `getOrCreateFolder(parentFolder, folderName)`: Crée ou récupère un dossier
- `protectValidationCell(sheet, row, colonne)`: Protège une cellule après validation
- `checkValidationCascade(sheet, row, colonne, statut)`: Vérifie que l'ordre de validation est respecté
- `notifyRequestorRejection(sheet, row, role, docId)`: Notifie le demandeur en cas de rejet

## Maintenance et débogage

Le script inclut une journalisation détaillée via `Logger.log()` pour faciliter le débogage:
- Données reçues du formulaire
- Création et accès aux dossiers
- Création et mise à jour des documents
- Envoi d'emails
- Erreurs rencontrées

Vous pouvez consulter ces logs dans l'éditeur Apps Script via le menu "Exécution" > "Journaux d'exécution".

### Résolution des problèmes courants

1. **Emails non reçus**:
   - Vérifiez que les adresses email sont correctement saisies
   - Vérifiez les quotas d'envoi d'emails pour Google Apps Script

2. **Erreurs de permission**:
   - Vérifiez que le compte exécutant le script a accès à tous les dossiers et documents
   - Vérifiez si des autorisations supplémentaires sont requises

3. **Document non généré**:
   - Vérifiez l'ID du document modèle
   - Vérifiez les logs pour des erreurs spécifiques

## Sécurité

Le script inclut:
- Protection des cellules de validation une fois les décisions prises
- Validation de l'ordre de cascade (Supérieur → RH → Présidence)
- Gestion des erreurs pour éviter les interruptions du script
- Validation des données pour éviter les valeurs nulles ou indéfinies

## Personnalisation avancée

### Modification des statuts de validation
Pour changer les libellés des statuts (actuellement "Favorable" et "Non Favorable"):
```javascript
// Chercher dans le code où ces statuts sont définis ou utilisés
// Par exemple dans handleValidation(), handleRejection(), etc.
```

### Ajout d'un niveau de validation supplémentaire
Nécessite des modifications dans:
- La structure de la feuille
- La fonction `checkValidationCascade()`
- La fonction `notifyNextValidator()`
- Le document modèle (ajout d'un nouveau placeholder)

## Limitations connues

- Le script est limité par les quotas Apps Script (ex: emails par jour, temps d'exécution)
- Les validations doivent être faites dans l'ordre séquentiel
- Une fois une décision prise, elle ne peut plus être modifiée

## Informations de licence et crédits

Ce script est fourni tel quel, sans garantie. Vous êtes libre de l'adapter aux besoins spécifiques de votre organisation.
