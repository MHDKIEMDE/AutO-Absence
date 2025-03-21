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
        // Vérification si le nom du dossier est valide
        if (!folderName || folderName.trim() === "" || folderName.includes("undefined")) {
            Logger.log(`[ERREUR DOSSIER] Nom de dossier invalide : "${folderName}"`);
            // Utiliser un nom par défaut si nécessaire
            folderName = "Dossier_" + new Date().getTime();
        }
        
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
    // Logger les données complètes pour le débogage
    Logger.log("Données reçues du formulaire: " + JSON.stringify(e.values));
    
    var data = e.values || [];

    // Extraction des données du formulaire avec valeurs par défaut pour éviter null/undefined
    var horodateur = data[0]|| "";
    var email = data[1]
    var matricule = data[2]
    var nom = data[3]
    var prenom = data[4]
    var serviceFonction = data[5]
    var typeAbsence = data[6]
    var dateDebut = data[7]
    var heureDebut = data[8]
    var dateFin = data[9]
    var heureFin = data[10]
    var motifAbsence = data[11]
    var typePermission = data[12]
    var motifPermission = data[13]
    var nbreJours = data[14]
    var emailSuperieur = data[15]
    var emailRH = "@gmail.com"; // À remplacer par l'email RH réel
    var emailPresidence = "@gmail.com"; // À remplacer par l'email de la présidence
    var statusSuperieurHierachique = "En attente"; // Initialisation par défaut
    var statusRH = "En attente"; // Initialisation par défaut
    var statusPresidence = "En attente"; // Initialisation par défaut
    var commentairePresidence = data[19]

    Logger.log(`[FORMULAIRE SOUMIS] Demandeur : ${nom} ${prenom} | Email : ${email}`);

    // Vérifier si les données essentielles sont présentes
    if (!email || !emailSuperieur) {
        Logger.log("[ERREUR] Email du demandeur ou du supérieur manquant !");
        return;
    }

    if (!dateDebut || dateDebut.trim() === "") {
        Logger.log("[ERREUR] Date de début manquante ou invalide !");
        return;
    }

    // Extraire l'année de la date de début avec gestion d'erreur
    var demandeYear = "";
    try {
        if (dateDebut && dateDebut.includes("/")) {
            var dateParts = dateDebut.split("/");
            demandeYear = dateParts.length >= 3 ? dateParts[2] : new Date().getFullYear().toString();
        } else {
            demandeYear = new Date().getFullYear().toString();
        }
    } catch (error) {
        Logger.log(`[ERREUR DATE] Impossible d'extraire l'année : ${error.message}`);
        demandeYear = new Date().getFullYear().toString();
    }
    
    // Créer un nom de dossier employé valide
    var employeeFolderName = "Employé_Inconnu";
    if (nom.trim() !== "" || prenom.trim() !== "") {
        employeeFolderName = `${nom.trim() || "Nom"} ${prenom.trim() || "Prénom"}`;
    }

    // Gestion des dossiers sur Google Drive
    var mainFolderId = "1H5GtauuqauEf_Ze4ZQ3kkoNEgyDJLsRw"; // ID du dossier principal
    var employeeFolderId = "1D3xmuzg4jN9kfr_pfBly1E2fuaEO2-nO"; // ID du dossier des employés
    var templateDocId = "1b_RRz7ie06UHuhfB4_5Sqeh33yhAD8gWeYVaAjrnz94"; // ID du modèle de document

    var mainFolder = DriveApp.getFolderById(mainFolderId);
    var employeeFolder = DriveApp.getFolderById(employeeFolderId);
    var yearFolder = getOrCreateFolder(employeeFolder, `Autorisation d'absence - ${demandeYear}`);
    var demandeurFolder = getOrCreateFolder(yearFolder, employeeFolderName);
    var pendingFolder = getOrCreateFolder(demandeurFolder, "En Attente");
    var approvedFolder = getOrCreateFolder(demandeurFolder, "Validée");
    var rejectedFolder = getOrCreateFolder(demandeurFolder, "Rejetée");

    // Vérifier si le dossier en attente a bien été créé
    if (!pendingFolder) {
        Logger.log("[ERREUR] Impossible de créer ou d'accéder au dossier En Attente");
        return;
    }

    // Création d'un document de demande basé sur le modèle
    try {
        var tempDoc = DriveApp.getFileById(templateDocId).makeCopy(`Demande_${prenom || "Prenom"}_${nom || "Nom"}`, pendingFolder);
        var tempDocId = tempDoc.getId();
        var doc = DocumentApp.openById(tempDocId);
        var body = doc.getBody();

        // Remplacement des placeholders dans le modèle
        var placeholders = {
            "{{matricule}}": matricule || "",
            "{{nom}}": nom || "",
            "{{prenom}}": prenom || "",
            "{{serviceFonction}}": serviceFonction || "",
            "{{typeAbsence}}": typeAbsence || "",
            "{{dateDebut}}": dateDebut || "",
            "{{heureDebut}}": heureDebut || "",
            "{{dateFin}}": dateFin || "",
            "{{heureFin}}": heureFin || "",
            "{{motifAbsence}}": motifAbsence || "",
            "{{typePermission}}": typePermission || "",
            "{{motifPermission}}": motifPermission || "",
            "{{nbreJours}}": nbreJours || "",
            "{{statusSuperieurHierachique}}": statusSuperieurHierachique || "",
            "{{statusRH}}": statusRH || "",
            "{{statusPresidence}}": statusPresidence || "",
            "{{commentairePresidence}}": commentairePresidence || "",
        };

        // Remplacer chaque placeholder en vérifiant qu'il existe dans le document
        for (var key in placeholders) {
            try {
                if (body.findText(key)) {
                    body.replaceText(key, placeholders[key]);
                }
            } catch (error) {
                Logger.log(`[ERREUR REMPLACEMENT] Impossible de remplacer ${key}: ${error.message}`);
            }
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
            `Bonjour ${prenom || "Demandeur"},<br><br>Votre demande a été soumise avec succès.<br> <a href="${tempDoc.getUrl()}">Consulter votre document</a>`);
    } catch (error) {
        Logger.log(`[ERREUR CRÉATION DOCUMENT] ${error.message}`);
    }
}

// Fonction pour notifier un validateur
function notifyValidator(email, role, docUrl, sheetUrl, docId) {
    try {
        if (!email || email.trim() === "") {
            Logger.log(`[ERREUR NOTIFICATION] Email du ${role} manquant ou invalide`);
            return;
        }
        
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

// Fonction pour notifier le prochain validateur dans la cascade
function notifyNextValidator(sheet, row, colonne, docId) {
    try {
        // Déterminer quel est le prochain validateur en fonction de la colonne qui vient d'être modifiée
        var nextRole = "";
        var nextEmail = "";
        
        if (colonne == 17) { // Si c'est le supérieur qui vient de valider
            nextRole = "RH";
            nextEmail = sheet.getRange(row, 22).getValue(); // Email RH en colonne 22
        } else if (colonne == 18) { // Si c'est le RH qui vient de valider
            nextRole = "Présidence";
            nextEmail = sheet.getRange(row, 23).getValue(); // Email Présidence en colonne 23
        } else {
            // Si c'est la présidence qui vient de valider, c'est la fin du processus
            // Notifier le demandeur de l'approbation finale
            var email = sheet.getRange(row, 2).getValue(); // Email du demandeur en colonne 2
            var doc = DriveApp.getFileById(docId);
            
            if (email) {
                var message = `Bonjour,<br><br>Votre demande d'absence a été entièrement validée.<br>` +
                    `<a href="${doc.getUrl()}">Consulter votre document</a><br><br>` +
                    `Cordialement.`;
                sendEmail(email, "Demande d'absence approuvée", message);
                Logger.log(`[NOTIFICATION APPROBATION FINALE] Email envoyé à ${email}`);
            }
            
            // Déplacer le document vers le dossier des validées
            try {
                var parentFolders = doc.getParents();
                if (parentFolders.hasNext()) {
                    var currentFolder = parentFolders.next();
                    var parentFolder = currentFolder.getParents().next();
                    
                    // Trouver le dossier des validées
                    var folders = parentFolder.getFolders();
                    while (folders.hasNext()) {
                        var folder = folders.next();
                        if (folder.getName() === "Validée") {
                            moveDocument(doc, currentFolder, folder);
                            Logger.log(`[DOCUMENT DÉPLACÉ] Document validé déplacé vers Validée.`);
                            break;
                        }
                    }
                }
            } catch (error) {
                Logger.log(`[ERREUR DÉPLACEMENT] Impossible de déplacer le document : ${error.message}`);
            }
            
            return;
        }
        
        // Notifier le prochain validateur s'il existe un email
        if (nextEmail && nextEmail.trim() !== "" && nextEmail.includes("@")) {
            var doc = DriveApp.getFileById(docId);
            var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
            notifyValidator(nextEmail, nextRole, doc.getUrl(), sheetUrl, docId);
            Logger.log(`[NOTIFICATION PROCHAINE ÉTAPE] Rôle : ${nextRole}, Email : ${nextEmail}`);
        } else {
            Logger.log(`[ERREUR NOTIFICATION] Email du ${nextRole} manquant ou invalide : ${nextEmail}`);
        }
    } catch (error) {
        Logger.log(`[ERREUR NOTIFICATION PROCHAINE ÉTAPE] : ${error.message}`);
    }
}

// Fonction pour protéger une cellule de validation après décision
function protectValidationCell(sheet, row, colonne) {
    try {
        var protection = sheet.getRange(row, colonne).protect();
        protection.setDescription(`Protection de décision - Ligne ${row}, Colonne ${colonne}`);
        // Supprimer tous les éditeurs sauf l'utilisateur actuel (propriétaire)
        var me = Session.getEffectiveUser();
        protection.removeEditors(protection.getEditors());
        protection.addEditor(me);
        Logger.log(`[PROTECTION] Cellule protégée : Ligne ${row}, Colonne ${colonne}`);
    } catch (error) {
        Logger.log(`[ERREUR PROTECTION] Impossible de protéger la cellule : ${error.message}`);
    }
}

// Fonction pour gérer la validation dans le document
function handleValidation(role, statut, docId) {
    try {
        var doc = DocumentApp.openById(docId);
        var body = doc.getBody();
        
        // Déterminer quel placeholder remplacer en fonction du rôle
        var placeholder = "";
        if (role === "Supérieur Hiérachique") {
            placeholder = "{{statusSuperieurHierachique}}";
        } else if (role === "RH") {
            placeholder = "{{statusRH}}";
        } else if (role === "Présidence") {
            placeholder = "{{statusPresidence}}";
        }
        
        // Remplacer le placeholder s'il existe
        if (placeholder && body.findText(placeholder)) {
            body.replaceText(placeholder, statut);
            Logger.log(`[DOCUMENT MIS À JOUR] Statut du ${role} mis à jour : ${statut}`);
        } else {
            Logger.log(`[ERREUR MISE À JOUR] Placeholder ${placeholder} non trouvé dans le document`);
        }
        
        doc.saveAndClose();
    } catch (error) {
        Logger.log(`[ERREUR DOCUMENT] Impossible de mettre à jour le document : ${error.message}`);
    }
}

// Fonction modifiée pour gérer les validations sur la feuille
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
            role = "Supérieur Hiérachique";
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
        
        // Si rejet, traiter immédiatement et sortir de la boucle
        if (newValue === "Non Favorable") {
            handleRejection(sheet, row, role, docId);
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

// Fonction pour gérer les rejets spécifiquement
function handleRejection(sheet, row, role, docId) {
    Logger.log(`[REJET] Traitement du rejet par ${role}, docId: ${docId}`);
    
    // Désactiver les validations suivantes
    if (role === "Supérieur Hiérarchique") {
        sheet.getRange(row, 18).setValue("N/A");
        sheet.getRange(row, 19).setValue("N/A");
    } else if (role === "RH") {
        sheet.getRange(row, 19).setValue("N/A");
    }
    
    // Appliquer le rejet dans le document
    handleValidation(role, "Non Favorable", docId);
    
    // Notifier le demandeur spécifiquement pour un rejet
    notifyRequestorRejection(sheet, row, role, docId);
}

// Fonction pour notifier le demandeur en cas de rejet
function notifyRequestorRejection(sheet, row, role, docId) {
    try {
        var doc = DriveApp.getFileById(docId);
        var docName = doc.getName();
        var nameParts = docName.replace("Demande_", "").split("_");
        var prenom = nameParts[0] || "Demandeur";
        var nom = nameParts[1] || "";
        
        // Récupérer l'email du demandeur depuis la feuille
        var email = sheet.getRange(row, 2).getValue(); // Colonne 2 pour l'email
        
        if (!email) {
            Logger.log("[ERREUR] Impossible de trouver l'email du demandeur !");
            return;
        }
        
        // Message de rejet spécifique
        var subject = `Votre demande d'absence a été rejetée par ${role}`;
        var message = `Bonjour ${prenom},<br><br>
            Votre demande d'absence a été rejetée par ${role}.<br>
            <a href="${doc.getUrl()}">Consulter votre document</a><br><br>
            <strong>Veuillez noter que cette demande ne peut plus être traitée.</strong> Si nécessaire, merci de soumettre une nouvelle demande.<br><br>
            Cordialement.`;
        
        sendEmail(email, subject, message);
        Logger.log(`[NOTIFICATION REJET] Email envoyé à ${email} concernant le rejet par ${role}`);
        
        // Déplacer le document vers le dossier des rejetées
        try {
            var parentFolders = doc.getParents();
            if (parentFolders.hasNext()) {
                var currentFolder = parentFolders.next();
                var parentFolder = currentFolder.getParents().next();
                
                // Trouver le dossier des rejetées
                var folders = parentFolder.getFolders();
                while (folders.hasNext()) {
                    var folder = folders.next();
                    if (folder.getName() === "Rejetée") {
                        moveDocument(doc, currentFolder, folder);
                        Logger.log(`[DOCUMENT DÉPLACÉ] Document rejeté déplacé vers Rejetée.`);
                        break;
                    }
                }
            }
        } catch (error) {
            Logger.log(`[ERREUR DÉPLACEMENT] Impossible de déplacer le document : ${error.message}`);
        }
        
    } catch (error) {
        Logger.log(`[ERREUR NOTIFICATION REJET] : ${error.message}`);
    }
}

// Fonction pour vérifier l'ordre de validation en cascade
function checkValidationCascade(sheet, row, colonne, statut) {
    // Pour Supérieur Hiérarchique (colonne 17), toujours autorisé
    if (colonne == 17) {
        return true;
    }
    
    // Vérifier que le supérieur hiérarchique a validé avant toute autre validation
    var statusSuperieurHierachique = sheet.getRange(row, 17).getValue();
    if (statusSuperieurHierachique !== "Favorable") {
        // Si le supérieur n'a pas encore validé ou a rejeté, personne d'autre ne peut valider
        return false;
    }
    
    // Pour RH (colonne 18), vérifier que Supérieur a approuvé
    if (colonne == 18) {
        return true; // Si on arrive ici, c'est que le supérieur a déjà approuvé
    }
    
    // Pour Présidence (colonne 19), vérifier que RH a approuvé
    if (colonne == 19) {
        var statusRH = sheet.getRange(row, 18).getValue();
        return statusRH === "Favorable";
    }
    
    return false;
}