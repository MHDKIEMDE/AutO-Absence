// Fonction pour envoyer un email
function sendEmail(recipient, subject, body) {
    try {
        // Envoie un email HTML
        GmailApp.sendEmail(recipient, subject, "", { htmlBody: body });
        Logger.log(`[EMAIL ENVOYÉ] À : ${recipient} | Sujet : ${subject}`);
    } catch (error) {
        Logger.log(`[ERREUR EMAIL] Échec d'envoi à : ${recipient} | Message : ${error.message}`);
    }
}

// Fonction pour créer ou récupérer un dossier
function getOrCreateFolder(parentFolder, folderName) {
    try {
        // Recherche ou création d'un dossier
        var folders = parentFolder.getFoldersByName(folderName);
        var folder = folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
        Logger.log(`[DOSSIER] Accès ou création du dossier : ${folderName}`);
        return folder;
    } catch (error) {
        Logger.log(`[ERREUR DOSSIER] Impossible d'accéder ou de créer le dossier : ${error.message}`);
        return null;
    }
}

// Fonction pour déplacer un document entre deux dossiers
function moveDocument(doc, sourceFolder, targetFolder) {
    try {
        sourceFolder.removeFile(doc);
        targetFolder.addFile(doc);
        Logger.log(`[DOCUMENT DÉPLACÉ] Depuis : ${sourceFolder.getName()} | Vers : ${targetFolder.getName()}`);
    } catch (error) {
        Logger.log(`[ERREUR DOCUMENT] Impossible de déplacer le document : ${error.message}`);
    }
}

// Fonction déclenchée lors de la soumission du formulaire
function onFormSubmit(e) {
    var data = e.values;

    // Extraction des données du formulaire
    var horodateur = data[0];
    var email = data[1];
    var matricule = data[2];
    var nom = data[3];
    var prenom = data[4];
    var serviceFonction = data[5];
    var typeAbsence = data[6];
    var dateDebut = data[7];
    var heureDebut = data[8];
    var dateFin = data[9];
    var heureFin = data[10];
    var motifAbsence = data[11];
    var typePermission = data[12];
    var motifPermission = data[13];
    var nombreJours = data[14];
    var emailSuperieur = data[15];
    var emailRH = "@gmail.com"; // À remplacer par l'email RH réel
    var emailPresidence = "@gmail.com"; // À remplacer par l'email de la présidence
    var statusSuperieurHierarchique = "En attente"; // Initialisation par défaut
    var statusRH = "En attente"; // Initialisation par défaut
    var statusPresidence = "En attente"; // Initialisation par défaut
    var commentairePresidence = data[19];

    Logger.log(`[FORMULAIRE SOUMIS] Demandeur : ${nom} ${prenom} | Email : ${email}`);

    if (!email || !emailSuperieur) {
        Logger.log("[ERREUR] Email du demandeur ou du supérieur manquant !");
        return;
    }

    // Gestion des dossiers sur Google Drive
    var mainFolderId = "1H5GtauuqauEf_Ze4ZQ3kkoNEgyDJLsRw"; // ID du dossier principal
    var employeeFolderId = "1D3xmuzg4jN9kfr_pfBly1E2fuaEO2-nO"; // ID du dossier des employés
    var templateDocId = "1b_RRz7ie06UHuhfB4_5Sqeh33yhAD8gWeYVaAjrnz94"; // ID du modèle de document

    var mainFolder = DriveApp.getFolderById(mainFolderId);
    var employeeFolder = DriveApp.getFolderById(employeeFolderId);
    var demandeYear = dateDebut.split("/")[2];
    var yearFolder = getOrCreateFolder(employeeFolder, `Autorisation d'absence - ${demandeYear}`);
    var demandeurFolder = getOrCreateFolder(yearFolder, `${nom} ${prenom}`);
    var pendingFolder = getOrCreateFolder(demandeurFolder, "En Attente");
    var approvedFolder = getOrCreateFolder(demandeurFolder, "Validée");
    var rejectedFolder = getOrCreateFolder(demandeurFolder, "Rejetée");

    // Création d'un document de demande basé sur le modèle
    var tempDoc = DriveApp.getFileById(templateDocId).makeCopy(`Demande_${prenom}_${nom}`, pendingFolder);
    var tempDocId = tempDoc.getId();
    var doc = DocumentApp.openById(tempDocId);
    var body = doc.getBody();

    // Remplacement des placeholders dans le modèle
    var placeholders = {
        "{{matricule}}": matricule,
        "{{nom}}": nom,
        "{{prenom}}": prenom,
        "{{serviceFonction}}": serviceFonction,
        "{{typeAbsence}}": typeAbsence,
        "{{dateDebut}}": dateDebut,
        "{{heurDebut}}": heureDebut,
        "{{dateFin}}": dateFin,
        "{{heurFin}}": heureFin,
        "{{motifAbsence}}": motifAbsence,
        "{{typePermission}}": typePermission,
        "{{motifPermission}}": motifPermission,
        "{{nbreJours}}": nombreJours,
        "{{statusSuperieurHirachique}}": statusSuperieurHierarchique,
        "{{statusRh}}": statusRH,
        "{{StatusPresidence}}": statusPresidence,
        "{{commentairePresidence}}":commentairePresidence,
    };

    for (var key in placeholders) {
        body.replaceText(key, placeholders[key]);
    }

    doc.saveAndClose();
    Logger.log(`[DOCUMENT CRÉÉ] ID : ${tempDocId} | URL : ${tempDoc.getUrl()}`);

    // Enregistrer l'ID du document et les emails pour les validations futures dans la feuille de calcul
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow(); // La dernière ligne où les données ont été ajoutées
    sheet.getRange(lastRow, 21).setValue(tempDocId); // ID du document en colonne 21
    sheet.getRange(lastRow, 22).setValue(emailRH); // Email RH en colonne 22
    sheet.getRange(lastRow, 23).setValue(emailPresidence); // Email Présidence en colonne 23
    Logger.log(`[INFOS DOCUMENT ENREGISTRÉES] ID et emails de validation enregistrés pour la ligne ${lastRow}`);

    // Notifications par email - uniquement au supérieur hiérarchique initialement
    notifyValidator(emailSuperieur, "Supérieur Hiérarchique", tempDoc.getUrl(), SpreadsheetApp.getActiveSpreadsheet().getUrl(), tempDocId);
    sendEmail(email, "Confirmation de votre demande d'absence", 
        `Bonjour ${prenom},<br><br>Votre demande a été soumise avec succès.<br> <a href="${tempDoc.getUrl()}">Consulter votre document</a>`);
}

// Fonction pour notifier un validateur
function notifyValidator(email, role, docUrl, sheetUrl, docId) {
    try {
        Logger.log(`[NOTIFICATION VALIDATEUR] Envoi au ${role} : ${email}`);
        var message = `Bonjour,<br><br>Une nouvelle demande d'absence nécessite votre validation.<br>` +
            ` <a href="${docUrl}">Consulter le document</a><br>` +
            ` <a href="${sheetUrl}">Accéder à la feuille des réponses</a><br>` +
            `Merci de valider ou rejeter la demande directement sur la feuille.<br><br>Cordialement.`;
        sendEmail(email, `Nouvelle demande d'absence à valider - Rôle : ${role}`, message);
    } catch (error) {
        Logger.log(`[ERREUR NOTIFICATION VALIDATEUR] Échec d'envoi : ${error.message}`);
    }
}

// Fonction pour gérer les validations sur la feuille
function onEditValidation(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;

    // Vérifie si une réponse de validation a été modifiée (colonnes 17-19)
    if (range.getColumn() >= 17 && range.getColumn() <= 19) {
        var row = range.getRow();
        var role = "";
        var colonne = range.getColumn();
        
        // Déterminer le rôle en fonction de la colonne modifiée
        if (colonne == 17) {
            role = "Supérieur Hiérarchique";
        } else if (colonne == 18) {
            role = "RH";
        } else if (colonne == 19) {
            role = "Présidence";
        }
        
        var oldValue = e.oldValue; // Récupérer l'ancienne valeur
        var newValue = range.getValue(); // Statut modifié ("Favorable" ou "Non Favorable")

        // Vérifier si la décision a déjà été prise (différente de "En attente")
        if (oldValue !== undefined && oldValue !== null && oldValue !== "" && oldValue !== "En attente") {
            // Restaurer l'ancienne valeur
            range.setValue(oldValue);
            SpreadsheetApp.getUi().alert("Impossible de modifier une décision déjà prise. La décision précédente (" + oldValue + ") a été restaurée.");
            Logger.log(`[TENTATIVE DE MODIFICATION] Rôle : ${role}, Ancienne valeur : ${oldValue}, Tentative : ${newValue}`);
            return;
        }
        
        var docId = sheet.getRange(row, 21).getValue(); // Colonne U pour l'ID du document

        Logger.log(`[DEBUG] Validation modifiée : Rôle : ${role}, Statut : ${newValue}, docId : ${docId}`);

        // Vérifier si l'ID du document existe
        if (!docId || docId.trim() === "") {
            Logger.log("[ERREUR] ID du document manquant ou invalide dans la cellule !");
            return;
        }
        
        // Vérifier la cascade de validation
        if (!checkValidationCascade(sheet, row, colonne, newValue)) {
            Logger.log("[ERREUR VALIDATION] Ordre de validation non respecté !");
            // Réinitialiser la cellule et informer l'utilisateur
            range.setValue("En attente");
            SpreadsheetApp.getUi().alert("Impossible de valider à cette étape. Veuillez respecter l'ordre de validation: Supérieur → RH → Présidence");
            return;
        }
        
        // Si rejet, traiter immédiatement
        if (newValue === "Non Favorable") {
            handleValidation(role, newValue, docId);
            // Désactiver les validations suivantes si rejet
            disableNextValidations(sheet, row, colonne);
            // Protéger cette cellule après avoir pris la décision
            protectValidationCell(sheet, row, colonne);
            return;
        }
        
        // Si favorable, traiter la validation et notifier le prochain validateur si nécessaire
        handleValidation(role, newValue, docId);
        
        // Notifier le prochain validateur dans la cascade si statut favorable
        notifyNextValidator(sheet, row, colonne, docId);
        
        // Protéger cette cellule après avoir pris la décision
        protectValidationCell(sheet, row, colonne);
    }
}

// Fonction pour protéger une cellule de validation après décision
function protectValidationCell(sheet, row, column) {
    try {
        // Créer une protection pour la cellule spécifique
        var range = sheet.getRange(row, column);
        var protection = range.protect();
        
        // Supprimer tous les éditeurs sauf le propriétaire de la feuille
        var me = Session.getEffectiveUser();
        protection.addEditor(me);
        protection.removeEditors(protection.getEditors());
        
        // Ajouter un message d'avertissement
        protection.setDescription("Décision de validation verrouillée");
        
        Logger.log(`[PROTECTION] Cellule verrouillée : Ligne ${row}, Colonne ${column}`);
    } catch (error) {
        Logger.log(`[ERREUR PROTECTION] Impossible de protéger la cellule : ${error.message}`);
    }
}

// Fonction pour vérifier l'ordre de validation en cascade
function checkValidationCascade(sheet, row, colonne, statut) {
    // Si c'est un rejet, on peut toujours le faire
    if (statut === "Non Favorable") {
        return true;
    }
    
    // Pour Supérieur Hiérarchique (colonne 17), toujours autorisé
    if (colonne == 17) {
        return true;
    }
    
    // Pour RH (colonne 18), vérifier que Supérieur a approuvé
    if (colonne == 18) {
        var statusSuperieur = sheet.getRange(row, 17).getValue();
        return statusSuperieur === "Favorable";
    }
    
    // Pour Présidence (colonne 19), vérifier que RH a approuvé
    if (colonne == 19) {
        var statusRH = sheet.getRange(row, 18).getValue();
        return statusRH === "Favorable";
    }
    
    return false;
}

// Fonction pour désactiver les validations suivantes en cas de rejet
function disableNextValidations(sheet, row, colonne) {
    // Si Supérieur rejette, désactiver RH et Présidence
    if (colonne == 17) {
        sheet.getRange(row, 18).setValue("N/A");
        sheet.getRange(row, 19).setValue("N/A");
    }
    
    // Si RH rejette, désactiver Présidence
    if (colonne == 18) {
        sheet.getRange(row, 19).setValue("N/A");
    }
}

// Fonction pour notifier le prochain validateur dans la cascade
function notifyNextValidator(sheet, row, colonne, docId) {
    // Si Supérieur valide, notifier RH
    if (colonne == 17) {
        var emailRH = sheet.getRange(row, 22).getValue(); // Email RH en colonne 22
        if (emailRH) {
            var docUrl = DriveApp.getFileById(docId).getUrl();
            var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
            notifyValidator(emailRH, "RH", docUrl, sheetUrl, docId);
        }
    }
    
    // Si RH valide, notifier Présidence
    if (colonne == 18) {
        var emailPresidence = sheet.getRange(row, 23).getValue(); // Email Présidence en colonne 23
        if (emailPresidence) {
            var docUrl = DriveApp.getFileById(docId).getUrl();
            var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
            notifyValidator(emailPresidence, "Présidence", docUrl, sheetUrl, docId);
        }
    }
}

// Fonction pour gérer la validation des documents
function handleValidation(role, statut, docId) {
    Logger.log(`[VALIDATION] Début pour rôle : ${role}, statut : ${statut}, ID : ${docId}`);

    if (!docId || docId.trim() === "") {
        Logger.log("[ERREUR] ID du document manquant ou invalide !");
        return;
    }

    try {
        var doc = DriveApp.getFileById(docId);
        Logger.log(`[DOCUMENT] Document trouvé : ${doc.getName()}`);
        
        // Trouver les dossiers du document
        var parentFolders = doc.getParents();
        if (!parentFolders.hasNext()) {
            Logger.log("[ERREUR] Document sans dossier parent !");
            return;
        }
        
        var currentFolder = parentFolders.next();
        var parentFolder = currentFolder.getParents().next();
        
        // Trouver les dossiers de destination dans le même dossier parent
        var pendingFolder = null;
        var approvedFolder = null;
        var rejectedFolder = null;
        
        var folders = parentFolder.getFolders();
        while (folders.hasNext()) {
            var folder = folders.next();
            var folderName = folder.getName();
            
            if (folderName === "En Attente") {
                pendingFolder = folder;
            } else if (folderName === "Validée") {
                approvedFolder = folder;
            } else if (folderName === "Rejetée") {
                rejectedFolder = folder;
            }
        }
        
        if (!pendingFolder || !approvedFolder || !rejectedFolder) {
            Logger.log("[ERREUR] Impossible de trouver tous les dossiers nécessaires !");
            return;
        }

        // Mettre à jour le document avec le nouveau statut
        var docContent = DocumentApp.openById(docId);
        var body = docContent.getBody();
        
        // Mettre à jour le statut dans le document en fonction du rôle
        if (role === "Supérieur Hiérarchique") {
            body.replaceText("{{statusSuperieurHirachique}}", statut);
        } else if (role === "RH") {
            body.replaceText("{{statusRh}}", statut);
        } else if (role === "Présidence") {
            body.replaceText("{{StatusPresidence}}", statut);
        }
        
        docContent.saveAndClose();
        
        // Déplacer le document si nécessaire
        if (statut === "Favorable" && role === "Présidence") {
            // Si la présidence approuve, déplacer vers le dossier des validées
            moveDocument(doc, pendingFolder, approvedFolder);
            Logger.log(`[VALIDATION] Document validé par la présidence et déplacé vers Validée.`);
        } else if (statut === "Non Favorable") {
            // Si quelqu'un rejette, déplacer vers le dossier des rejetées
            moveDocument(doc, pendingFolder, rejectedFolder);
            Logger.log(`[REJET] Document rejeté par ${role} et déplacé vers Rejetée.`);
        }

        // Notifier le demandeur du changement de statut
        notifyRequestor(doc, role, statut);

    } catch (error) {
        Logger.log(`[ERREUR GESTION VALIDATION] : ${error.message}`);
    }
}

// Fonction pour notifier le demandeur des changements de statut
function notifyRequestor(doc, role, statut) {
    try {
        // Récupérer les informations du document
        var docContent = DocumentApp.openById(doc.getId());
        var body = docContent.getBody();
        
        // Extraire le nom et le prénom depuis le document ou son nom
        var docName = doc.getName(); // Format: Demande_Prenom_Nom
        var nameParts = docName.replace("Demande_", "").split("_");
        var prenom = nameParts[0] || "Demandeur";
        var nom = nameParts[1] || "";
        
        // Trouver l'email du demandeur depuis la feuille
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        var dataRange = sheet.getDataRange();
        var values = dataRange.getValues();
        var email = "";
        
        // Chercher l'entrée correspondant à cet ID de document
        for (var i = 0; i < values.length; i++) {
            if (values[i][20] === doc.getId()) { // Colonne 21 (index 20)
                email = values[i][1]; // Colonne 2 (index 1) pour l'email
                break;
            }
        }
        
        if (!email) {
            Logger.log("[ERREUR] Impossible de trouver l'email du demandeur !");
            return;
        }
        
        // Envoyer une notification
        var message = "";
        if (statut === "Favorable" && role === "Présidence") {
            // Notification finale d'approbation
            var subject = "Votre demande d'absence a été approuvée";
            message = `Bonjour ${prenom},<br><br>
                Votre demande d'absence a été complètement approuvée.<br>
                <a href="${doc.getUrl()}">Consulter votre document</a><br><br>
                Cordialement.`;
        } else if (statut === "Non Favorable") {
            // Notification de rejet
            var subject = "Votre demande d'absence a été rejetée";
            message = `Bonjour ${prenom},<br><br>
                Votre demande d'absence a été rejetée par ${role}.<br>
                <a href="${doc.getUrl()}">Consulter votre document</a><br><br>
                Cordialement.`;
        } else {
            // Notification d'étape de validation
            var subject = `Mise à jour de votre demande d'absence - Validée par ${role}`;
            message = `Bonjour ${prenom},<br><br>
                Votre demande d'absence a été validée par ${role} et passe à l'étape suivante de validation.<br>
                <a href="${doc.getUrl()}">Consulter votre document</a><br><br>
                Cordialement.`;
        }
        
        sendEmail(email, subject, message);
        Logger.log(`[NOTIFICATION DEMANDEUR] Email envoyé à ${email} concernant le statut ${statut} par ${role}`);
        
    } catch (error) {
        Logger.log(`[ERREUR NOTIFICATION DEMANDEUR] : ${error.message}`);
    }
}

https://github.com/MHDKIEMDE/AutO.git