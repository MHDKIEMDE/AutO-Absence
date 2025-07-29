const CONFIG = {
    emails: {
      rh: "",
      presidence: ["",""]
    }
  };
  
  // Fonction utilitaire pour obtenir une valeur sécurisée
  function getSafeValue(value, defaultValue = "") {
      return value !== null && value !== undefined ? value.toString().trim() : defaultValue;
  }
  
  // Fonction utilitaire pour extraire l'année d'une date
  function extractYear(dateValue) {
      try {
          if (!dateValue) return new Date().getFullYear().toString();
          
          // Si c'est un objet Date
          if (dateValue instanceof Date) {
              return dateValue.getFullYear().toString();
          }
          
          // Si c'est une string
          if (typeof dateValue === 'string' && dateValue.includes("/")) {
              const dateParts = dateValue.split("/");
              return dateParts.length >= 3 ? dateParts[2] : new Date().getFullYear().toString();
          }
          
          return new Date().getFullYear().toString();
      } catch (error) {
          Logger.log(`[ERREUR DATE] ${error.message}`);
          return new Date().getFullYear().toString();
      }
  }
  
  // Fonction pour envoyer un email
  function sendEmail(recipient, subject, body) {
      try {
          if (!recipient || !recipient.includes("@")) {
              Logger.log(`[ERREUR EMAIL] Email invalide : ${recipient}`);
              return false;
          }
          
          // Envoie un email HTML
          GmailApp.sendEmail(recipient, subject, "", { htmlBody: body });
          Logger.log(`[EMAIL ENVOYÉ] À : ${recipient} | Sujet : ${subject}`);
          return true;
      } catch (error) {
          Logger.log(`[ERREUR EMAIL] Échec d'envoi à : ${recipient} | Message : ${error.message}`);
          return false;
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
          return true;
      } catch (error) {
          Logger.log(`[ERREUR DOCUMENT] Impossible de déplacer le document : ${error.message}`);
          return false;
      }
  }
  
  // Fonction pour formater la date et l'heure actuelles
  function getFormattedDateTime() {
      var now = new Date();
      var day = ("0" + now.getDate()).slice(-2);
      var month = ("0" + (now.getMonth() + 1)).slice(-2);
      var year = now.getFullYear();
      var hours = ("0" + now.getHours()).slice(-2);
      var minutes = ("0" + now.getMinutes()).slice(-2);
  
      return `${day}/${month}/${year} ${hours}:${minutes}`;
  }
  
  // CORRIGÉ: Fonction pour obtenir l'email de présidence
  function getPresidenceEmail(superieurEmail) {
      try {
          let presEmails = CONFIG.emails.presidence;
          
          // Vérification de sécurité
          if (!presEmails || presEmails.length === 0) {
              Logger.log("[ERREUR] Aucun email de présidence configuré");
              return null;
          }
          
          // Si l'email du supérieur correspond à un email de la présidence,
          // utiliser l'autre email de la présidence
          if (presEmails.includes(superieurEmail)) {
              // Trouver l'autre email de présidence
              const otherEmail = presEmails.filter(email => email !== superieurEmail)[0];
              return otherEmail || presEmails[0]; // Fallback sur le premier si pas d'autre
          }
          
          // Par défaut, retourner le premier email de présidence
          return presEmails[0];
      } catch (error) {
          Logger.log(`[ERREUR getPresidenceEmail] ${error.message}`);
          return CONFIG.emails.presidence && CONFIG.emails.presidence.length > 0 ? CONFIG.emails.presidence[0] : null;
      }
  }
  
  // CORRIGÉ: Fonction déclenchée lors de la soumission du formulaire
  function onFormSubmit(e) {
      try {
          // Logger les données complètes pour le débogage
          Logger.log("Données reçues du formulaire: " + JSON.stringify(e.values));
          
          var data = e.values || [];
  
          // CORRIGÉ: Extraction sécurisée des données du formulaire
          var horodateur = getSafeValue(data[0]);
          var email = getSafeValue(data[1]);
          var matricule = getSafeValue(data[2]);
          var nom = getSafeValue(data[3]);
          var prenom = getSafeValue(data[4]);
          var serviceFonction = getSafeValue(data[5]);
          var typeAbsence = getSafeValue(data[6]);
          var dateDebut = data[7]; // Gardé tel quel pour gestion Date/String
          var heureDebut = getSafeValue(data[8]);
          var dateFin = data[9]; // Gardé tel quel pour gestion Date/String
          var heureFin = getSafeValue(data[10]);
          var motifAbsence = getSafeValue(data[11]);
          var typePermission = getSafeValue(data[12]);
          var motifPermission = getSafeValue(data[13]);
          var nbreJours = getSafeValue(data[14]);
          var emailSuperieur = getSafeValue(data[15]);
          
          // Utilisation du tableau de configuration pour les emails RH et Présidence
          var emailRH = CONFIG.emails.rh;
          var emailPresidence = getPresidenceEmail(emailSuperieur);
          
          Logger.log(`[FORMULAIRE SOUMIS] Demandeur : ${nom} ${prenom} | Email : ${email}`);
  
          // CORRIGÉ: Vérifications de données essentielles AVANT traitement
          if (!email || !email.includes("@")) {
              Logger.log("[ERREUR] Email du demandeur manquant ou invalide !");
              return;
          }
          
          if (!emailSuperieur || !emailSuperieur.includes("@")) {
              Logger.log("[ERREUR] Email du supérieur manquant ou invalide !");
              return;
          }
          
          if (!dateDebut) {
              Logger.log("[ERREUR] Date de début manquante !");
              return;
          }
          
          if (!emailPresidence) {
              Logger.log("[ERREUR] Impossible d'obtenir l'email de présidence !");
              return;
          }
  
          // CORRIGÉ: Extraction sécurisée de l'année
          var demandeYear = extractYear(dateDebut);
          
          // Créer un nom de dossier employé valide
          var employeeFolderName = "Employé_Inconnu";
          if (nom || prenom) {
              employeeFolderName = `${nom || "Nom"} ${prenom || "Prénom"}`;
          }
  
          // ✅ MODIFIÉ: Gestion des dossiers sur Google Drive avec le nouveau dossier
          var mainFolderId = "1H5GtauuqauEf_Ze4ZQ3kkoNEgyDJLsRw"; // ID du dossier principal (inchangé)
          var employeeFolderId = "1YS5_UiwR3I8y0f61A7WtyhoGU1T0mpsv"; // ✅ NOUVEAU: ID du dossier des employés
          var templateDocId = "1b_RRz7ie06UHuhfB4_5Sqeh33yhAD8gWeYVaAjrnz94"; // ID du modèle de document (inchangé)
  
          Logger.log(`[NOUVEAU DOSSIER] Utilisation du dossier employés : ${employeeFolderId}`);
  
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
  
          Logger.log(`[STRUCTURE DOSSIERS] Création réussie dans le nouveau dossier`);
          Logger.log(`[DOSSIER ANNÉE] ${yearFolder.getName()}`);
          Logger.log(`[DOSSIER EMPLOYÉ] ${demandeurFolder.getName()}`);
          Logger.log(`[DOSSIER EN ATTENTE] ${pendingFolder.getName()}`);
  
          // Création d'un document de demande basé sur le modèle
          var tempDoc = DriveApp.getFileById(templateDocId).makeCopy(`Demande_${prenom || "Prenom"}_${nom || "Nom"}`, pendingFolder);
          var tempDocId = tempDoc.getId();
          var doc = DocumentApp.openById(tempDocId);
          var body = doc.getBody();
  
          // CORRIGÉ: Gestion des dates pour les placeholders
          var dateDebutStr = dateDebut instanceof Date ? 
              dateDebut.toLocaleDateString('fr-FR') : 
              getSafeValue(dateDebut);
          
          var dateFinStr = dateFin instanceof Date ? 
              dateFin.toLocaleDateString('fr-FR') : 
              getSafeValue(dateFin);
  
          var placeholders = {
              "{{matricule}}": matricule,
              "{{nom}}": nom,
              "{{prenom}}": prenom,
              "{{serviceFonction}}": serviceFonction,
              "{{typeAbsence}}": typeAbsence,
              "{{dateDebut}}": dateDebutStr,
              "{{heureDebut}}": heureDebut,
              "{{dateFin}}": dateFinStr,
              "{{heureFin}}": heureFin,
              "{{motifAbsence}}": motifAbsence,
              "{{typePermission}}": typePermission,
              "{{motifPermission}}": motifPermission,
              "{{nbreJours}}": nbreJours,
              "{{statusSuperieurHierarchique}}": "Statut SH: ",
              "{{statusRH}}": "Statut RH:",
              "{{statusPresidence}}": "Statut PR:",
              "{{commentairePresidence}}": "Comment PR:"
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
          Logger.log(`[EMPLACEMENT DOCUMENT] Document créé dans le nouveau dossier : ${pendingFolder.getName()}`);
  
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
          if (!email || !email.includes("@")) {
              Logger.log(`[ERREUR NOTIFICATION] Email du ${role} manquant ou invalide : ${email}`);
              return false;
          }
          
          Logger.log(`[NOTIFICATION VALIDATEUR] Envoi au ${role} : ${email}`);
          var message = `Bonjour,<br><br>Une nouvelle demande d'absence nécessite votre validation.<br>` +
              ` <a href="${docUrl}">Consulter le document</a><br>` +
              ` <a href="${sheetUrl}">Accéder à la feuille des réponses</a><br>` +
              `Merci de valider ou rejeter la demande directement sur la feuille.<br><br>Cordialement.`;
          
          return sendEmail(email, `Nouvelle demande d'absence à valider - Rôle : ${role}`, message);
      } catch (error) {
          Logger.log(`[ERREUR NOTIFICATION VALIDATEUR] Échec d'envoi : ${error.message}`);
          return false;
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
              
              if (email && email.includes("@")) {
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
                              Logger.log(`[DOCUMENT DÉPLACÉ] Document validé déplacé vers Validée dans le nouveau dossier.`);
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
          if (nextEmail && nextEmail.includes("@")) {
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
  
  // CORRIGÉ: Fonction pour protéger une cellule de validation après décision
  function protectValidationCell(sheet, row, colonne) {
      try {
          var protection = sheet.getRange(row, colonne).protect();
          protection.setDescription(`Protection de décision - Ligne ${row}, Colonne ${colonne}`);
          
          // Gestion sécurisée des permissions
          try {
              var me = Session.getEffectiveUser();
              if (me) {
                  protection.removeEditors(protection.getEditors());
                  protection.addEditor(me);
              }
          } catch (permError) {
              Logger.log(`[AVERTISSEMENT PROTECTION] Impossible de gérer les permissions : ${permError.message}`);
              // Continuer sans erreur fatale
          }
          
          Logger.log(`[PROTECTION] Cellule protégée : Ligne ${row}, Colonne ${colonne}`);
          return true;
      } catch (error) {
          Logger.log(`[ERREUR PROTECTION] Impossible de protéger la cellule : ${error.message}`);
          return false;
      }
  }

// SOLUTION CORRIGÉE pour handleValidation
function handleValidation(role, statut, docId, decisionTime) {
    try {
        var doc = DocumentApp.openById(docId);
        var body = doc.getBody();
        
        // DIAGNOSTIC: D'abord, voir ce que contient le document
        Logger.log(`[DIAGNOSTIC] Contenu du document: ${body.getText().substring(0, 500)}...`);
        
        // Méthode 1: Chercher les placeholders exacts utilisés lors de la création
        var searchTexts = {
            "Supérieur Hiérarchique": ["Statut SH:", "{{statusSuperieurHierarchique}}"],
            "RH": ["Statut RH:", "{{statusRH}}"],
            "Présidence": ["Statut PR:", "{{statusPresidence}}"]
        };
        
        var prefixes = {
            "Supérieur Hiérarchique": "SH",
            "RH": "RH", 
            "Présidence": "PR"
        };
        
        var prefix = prefixes[role];
        var searchOptions = searchTexts[role];
        
        if (!prefix || !searchOptions) {
            Logger.log(`[ERREUR MISE À JOUR] Rôle non reconnu : ${role}`);
            return false;
        }
        
        var statutWithTime = `Statut ${prefix}: ${statut} (${decisionTime})`;
        var updated = false;
        
        // Essayer chaque option de recherche
        for (var i = 0; i < searchOptions.length; i++) {
            var searchText = searchOptions[i];
            Logger.log(`[RECHERCHE] Tentative avec: "${searchText}"`);
            
            var foundText = body.findText(searchText);
            if (foundText) {
                var element = foundText.getElement();
                var startOffset = foundText.getStartOffset();
                var endOffset = foundText.getEndOffsetInclusive();
                
                // Remplacer le texte trouvé
                element.asText().deleteText(startOffset, endOffset);
                element.asText().insertText(startOffset, statutWithTime);
                
                Logger.log(`[DOCUMENT MIS À JOUR] ${role}: "${searchText}" remplacé par "${statutWithTime}"`);
                updated = true;
                break;
            } else {
                Logger.log(`[RECHERCHE] "${searchText}" non trouvé`);
            }
        }
        
        // Si aucune des options n'a fonctionné, méthode alternative robuste
        if (!updated) {
            Logger.log(`[MÉTHODE ALTERNATIVE] Recherche de "En attente" pour ${role}`);
            
            // Déterminer quelle occurrence de "En attente" remplacer
            var targetOccurrence = 0;
            if (role === "Supérieur Hiérarchique") targetOccurrence = 1;
            else if (role === "RH") targetOccurrence = 2;
            else if (role === "Présidence") targetOccurrence = 3;
            
            var currentOccurrence = 0;
            var searchResult = null;
            
            // Chercher toutes les occurrences de "En attente"
            while ((searchResult = body.findText("En attente", searchResult)) !== null) {
                currentOccurrence++;
                Logger.log(`[OCCURRENCE] Trouvé occurrence #${currentOccurrence} de "En attente"`);
                
                if (currentOccurrence === targetOccurrence) {
                    var element = searchResult.getElement();
                    var startOffset = searchResult.getStartOffset();
                    var endOffset = searchResult.getEndOffsetInclusive();
                    
                    element.asText().deleteText(startOffset, endOffset);
                    element.asText().insertText(startOffset, `${statut} (${decisionTime})`);
                    
                    Logger.log(`[DOCUMENT MIS À JOUR ALTERNATIVE] ${role}: occurrence #${targetOccurrence} de "En attente" remplacée`);
                    updated = true;
                    break;
                }
            }
        }
        
        // Méthode de dernier recours: ajout à la fin du document
        if (!updated) {
            Logger.log(`[DERNIER RECOURS] Ajout du statut à la fin du document`);
            body.appendParagraph(`\n${role}: ${statut} (${decisionTime})`);
            updated = true;
        }
        
        doc.saveAndClose();
        return updated;
        
    } catch (error) {
        Logger.log(`[ERREUR DOCUMENT] Impossible de mettre à jour le document : ${error.message}`);
        return false;
    }
}

  // CORRIGÉ: Fonction pour gérer les rejets spécifiquement
  function handleRejection(sheet, row, role, docId, decisionTime) {
      try {
          Logger.log(`[REJET] Traitement du rejet par ${role}, Horodatage : ${decisionTime}, docId: ${docId}`);
          
          // Désactiver les validations suivantes
          if (role === "Supérieur Hiérarchique") {
              sheet.getRange(row, 18).setValue("N/A");
              sheet.getRange(row, 19).setValue("N/A");
              Logger.log(`[REJET SUPÉRIEUR] Validations RH et Présidence désactivées avec N/A`);
          } else if (role === "RH") {
              sheet.getRange(row, 19).setValue("N/A");
              Logger.log(`[REJET RH] Validation Présidence désactivée avec N/A`);
          }
          
          var rejectionStatus = "Non Favorable";
          
          // Appliquer le rejet dans le document avec horodatage
          handleValidation(role, rejectionStatus, docId, decisionTime);
          
          // Notifier le demandeur spécifiquement pour un rejet
          notifyRequestorRejection(sheet, row, role, docId, decisionTime);
          
          return true;
      } catch (error) {
          Logger.log(`[ERREUR REJET] ${error.message}`);
          return false;
      }
  }
  
  // Fonction pour notifier le demandeur en cas de rejet
  function notifyRequestorRejection(sheet, row, role, docId, decisionTime) {
      try {
          var doc = DriveApp.getFileById(docId);
          var docName = doc.getName();
          var nameParts = docName.replace("Demande_", "").split("_");
          var prenom = nameParts[0] || "Demandeur";
          var nom = nameParts[1] || "";
          
          // Récupérer l'email du demandeur depuis la feuille
          var email = sheet.getRange(row, 2).getValue(); // Colonne 2 pour l'email
          
          if (!email || !email.includes("@")) {
              Logger.log("[ERREUR] Impossible de trouver l'email du demandeur ou email invalide !");
              return false;
          }
          
          // Message de rejet spécifique incluant l'horodatage
          var subject = `Votre demande d'absence a été rejetée par ${role}`;
          var message = `Bonjour ${prenom},<br><br>
              Votre demande d'absence a été rejetée par ${role} le ${decisionTime}.<br>
              <a href="${doc.getUrl()}">Consulter votre document</a><br><br>
              <strong>Veuillez noter que cette demande ne peut plus être traitée.</strong> <br>Si nécessaire, merci de soumettre une nouvelle demande.<br><br>
              Cordialement.`;
          
          sendEmail(email, subject, message);
          Logger.log(`[NOTIFICATION REJET] Email envoyé à ${email} concernant le rejet par ${role} à ${decisionTime}`);
          
          // Déplacer le document vers le dossier des rejetées dans le nouveau dossier
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
                          Logger.log(`[DOCUMENT DÉPLACÉ] Document rejeté déplacé vers Rejetée dans le nouveau dossier à ${decisionTime}.`);
                          break;
                      }
                  }
              }
          } catch (error) {
              Logger.log(`[ERREUR DÉPLACEMENT] Impossible de déplacer le document : ${error.message}`);
          }
          
          return true;
      } catch (error) {
          Logger.log(`[ERREUR NOTIFICATION REJET] : ${error.message}`);
          return false;
      }
  }
  
  // CORRIGÉ: Fonction pour vérifier l'ordre de validation en cascade
  function checkValidationCascade(sheet, row, colonne, statut) {
      try {
          // Pour Supérieur Hiérarchique (colonne 17), toujours autorisé
          if (colonne == 17) {
              return true;
          }
          
          // Vérifier que le supérieur hiérarchique a validé avant toute autre validation
          var statusSuperieurHierarchique = sheet.getRange(row, 17).getValue();
          
          // CORRIGÉ: Vérification robuste de la valeur
          if (!statusSuperieurHierarchique || statusSuperieurHierarchique.toString().trim() !== "Favorable") {
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
              return statusRH && statusRH.toString().trim() === "Favorable";
          }
          
          return false;
      } catch (error) {
          Logger.log(`[ERREUR VALIDATION CASCADE] ${error.message}`);
          return false;
      }
  }
  
  // CORRIGÉ: Fonction pour gérer les validations sur la feuille
  function onEditValidation(e) {
    try {
        var sheet = e.source.getActiveSheet();
        var range = e.range;

        // Vérifie si une réponse de validation a été modifiée (colonnes 17-19)
        if (range.getColumn() >= 17 && range.getColumn() <= 19) {
            var row = range.getRow();
            var role = "";
            var colonne = range.getColumn();
            var decisionTime = getFormattedDateTime();
            
            // Déterminer le rôle en fonction de la colonne modifiée
            if (colonne == 17) {
                role = "Supérieur Hiérarchique";
            } else if (colonne == 18) {
                role = "RH";
            } else if (colonne == 19) {
                role = "Présidence";
            }
            
            var oldValue = e.oldValue;
            var newValue = range.getValue();

            // Vérifier si la décision a déjà été prise
            if (oldValue !== undefined && oldValue !== null && oldValue !== "" && oldValue !== "En attente") {
                range.setValue(oldValue);
                SpreadsheetApp.getUi().alert("Impossible de modifier une décision déjà prise. La décision précédente (" + oldValue + ") a été restaurée.");
                Logger.log(`[TENTATIVE DE MODIFICATION] Rôle : ${role}, Ancienne valeur : ${oldValue}, Tentative : ${newValue}`);
                return;
            }
            
            var docId = sheet.getRange(row, 21).getValue();

            Logger.log(`[DEBUG] Validation modifiée : Rôle : ${role}, Statut : ${newValue}, Horodatage : ${decisionTime}, docId : ${docId}`);

            // Vérifier si l'ID du document existe
            if (!docId || docId.toString().trim() === "") {
                Logger.log("[ERREUR] ID du document manquant ou invalide dans la cellule !");
                return;
            }
            
            // Vérifier la cascade de validation
            if (!checkValidationCascade(sheet, row, colonne, newValue)) {
                Logger.log("[ERREUR VALIDATION] Ordre de validation non respecté !");
                range.setValue("En attente");
                SpreadsheetApp.getUi().alert("Impossible de valider à cette étape. Veuillez respecter l'ordre de validation: Supérieur → RH → Présidence");
                return;
            }
            
            // Corriger les erreurs d'orthographe
            if (newValue === "Nom Favorable") {
                newValue = "Non Favorable";
                range.setValue(newValue);
                Logger.log(`[CORRECTION] Statut "Nom Favorable" corrigé en "Non Favorable"`);
            }
            
            // Validation des statuts autorisés
            if (newValue !== "Favorable" && newValue !== "Non Favorable") {
                Logger.log(`[ERREUR] Statut non autorisé : ${newValue}`);
                range.setValue("En attente");
                SpreadsheetApp.getUi().alert("Statut non autorisé. Veuillez choisir 'Favorable' ou 'Non Favorable'.");
                return;
            }
            
            // ✅ CORRECTION PRINCIPALE: Notifications d'abord, document après
            
            // Si rejet
            if (newValue === "Non Favorable") {
                // D'abord les notifications de rejet
                notifyRequestorRejection(sheet, row, role, docId, decisionTime);
                
                // Désactiver les validations suivantes
                if (role === "Supérieur Hiérarchique") {
                    sheet.getRange(row, 18).setValue("N/A");
                    sheet.getRange(row, 19).setValue("N/A");
                } else if (role === "RH") {
                    sheet.getRange(row, 19).setValue("N/A");
                }
                
                // Protéger la cellule
                protectValidationCell(sheet, row, colonne);
                
                // Essayer de mettre à jour le document (sans bloquer si ça échoue)
                try {
                    handleValidation(role, newValue, docId, decisionTime);
                } catch (docError) {
                    Logger.log(`[AVERTISSEMENT] Mise à jour document échouée: ${docError.message}`);
                }
                
                return;
            }
            
            // Si favorable - ENVOYER LES NOTIFICATIONS D'ABORD
            Logger.log(`[PRIORITÉ] Envoi des notifications avant mise à jour du document`);
            
            // ✅ NOTIFICATIONS EN PREMIER (même si le document échoue)
            notifyNextValidator(sheet, row, colonne, docId);
            
            // Protéger la cellule
            protectValidationCell(sheet, row, colonne);
            
            // Essayer la mise à jour du document APRÈS les notifications
            try {
                var updateSuccess = handleValidation(role, newValue, docId, decisionTime);
                if (updateSuccess) {
                    Logger.log(`[DOCUMENT] Mise à jour réussie pour ${role}`);
                } else {
                    Logger.log(`[AVERTISSEMENT] Mise à jour document échouée pour ${role}, mais notifications envoyées`);
                }
            } catch (docError) {
                Logger.log(`[AVERTISSEMENT] Erreur mise à jour document pour ${role}: ${docError.message}`);
                Logger.log(`[INFO] Les notifications ont été envoyées malgré l'erreur du document`);
            }
        }
    } catch (error) {
        Logger.log(`[ERREUR CRITIQUE onEditValidation] ${error.message}`);
        try {
            SpreadsheetApp.getUi().alert("Une erreur inattendue s'est produite. Veuillez vérifier les logs ou contacter l'administrateur.");
        } catch (uiError) {
            Logger.log(`[ERREUR UI] Impossible d'afficher l'alerte : ${uiError.message}`);
        }
    }
}

// FONCTION DE DIAGNOSTIC pour comprendre la structure du document
function debugDocumentStructure(docId) {
    try {
        var doc = DocumentApp.openById(docId);
        var body = doc.getBody();
        var fullText = body.getText();
        
        Logger.log(`[DEBUG] Contenu complet du document:`);
        Logger.log(fullText);
        
        // Chercher tous les placeholders possibles
        var placeholders = [
            "{{statusSuperieurHierarchique}}", "{{statusRH}}", "{{statusPresidence}}",
            "Statut SH:", "Statut RH:", "Statut PR:",
            "En attente"
        ];
        
        placeholders.forEach(function(placeholder) {
            var found = body.findText(placeholder);
            if (found) {
                Logger.log(`[DEBUG] TROUVÉ: "${placeholder}" à la position ${found.getStartOffset()}`);
            } else {
                Logger.log(`[DEBUG] NON TROUVÉ: "${placeholder}"`);
            }
        });
        
        return fullText;
    } catch (error) {
        Logger.log(`[ERREUR DEBUG] ${error.message}`);
        return null;
    }
}

// FONCTION UTILITAIRE pour analyser et réparer le template
function analyzeAndFixTemplate(templateDocId) {
    try {
        var doc = DocumentApp.openById(templateDocId);
        var body = doc.getBody();
        var fullText = body.getText();
        
        Logger.log(`[ANALYSE TEMPLATE] Contenu actuel:`);
        Logger.log(fullText);
        
        // Vérifier si les placeholders de statut existent
        var statusPlaceholders = [
            "{{statusSuperieurHierarchique}}",
            "{{statusRH}}",
            "{{statusPresidence}}"
        ];
        
        var missingPlaceholders = [];
        statusPlaceholders.forEach(function(placeholder) {
            var found = body.findText(placeholder);
            if (!found) {
                missingPlaceholders.push(placeholder);
            }
        });
        
        if (missingPlaceholders.length > 0) {
            Logger.log(`[TEMPLATE] Placeholders manquants: ${missingPlaceholders.join(', ')}`);
            
            // Ajouter les placeholders manquants à la fin du document
            body.appendParagraph("\n--- SECTION VALIDATION ---");
            body.appendParagraph("Statut Supérieur Hiérarchique: En attente");
            body.appendParagraph("Statut RH: En attente"); 
            body.appendParagraph("Statut Présidence: En attente");
            
            doc.saveAndClose();
            Logger.log(`[TEMPLATE] Placeholders ajoutés au template`);
        } else {
            Logger.log(`[TEMPLATE] Tous les placeholders sont présents`);
        }
        
        return true;
    } catch (error) {
        Logger.log(`[ERREUR ANALYSE TEMPLATE] ${error.message}`);
        return false;
    }
}