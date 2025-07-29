// WORKFLOW PRINCIPAL - Gestion des demandes d'absence

// MÉTADONNÉES - Génère les données cachées en arrière-plan
function generateBackgroundMetadata(sheet, row, docId) {
    try {
      // Récupère les emails nécessaires
      var emailSuperieur = sheet.getRange(row, 16).getValue();    // Colonne P
      var emailRH = CONFIG.emails.rh;
      var emailPresidence = getPresidenceEmail(emailSuperieur);
  
      // Enregistre dans les colonnes cachées (U, V, W, X, Y)
      var metaStartCol = 21;
      sheet.getRange(row, metaStartCol).setValue(docId);                    // U: ID document
      sheet.getRange(row, metaStartCol + 1).setValue(emailRH);              // V: Email RH  
      sheet.getRange(row, metaStartCol + 2).setValue(emailPresidence);      // W: Email Présidence
      sheet.getRange(row, metaStartCol + 3).setValue(getFormattedDateTime()); // X: Timestamp
      sheet.getRange(row, metaStartCol + 4).setValue("Créé");               // Y: Statut
      
      Logger.log(`[MÉTADONNÉES] Générées pour ligne ${row}`);
      return true;
    } catch (error) {
      Logger.log(`[ERREUR MÉTADONNÉES] ${error.message}`);
      return false;
    }
  }
  
  // SOUMISSION FORMULAIRE - Point d'entrée principal
  function onFormSubmit(e) {
    try {
      Logger.log("=== DÉBUT TRAITEMENT FORMULAIRE ===");
      
      var data = e.values || [];
      
      // ÉTAPE 1: Extraction des données du formulaire
      var horodateur = getSafeValue(data[0]);
      var email = getSafeValue(data[1]);
      var matricule = getSafeValue(data[2]);
      var nom = getSafeValue(data[3]);
      var prenom = getSafeValue(data[4]);
      var serviceFonction = getSafeValue(data[5]);
      var typeAbsence = getSafeValue(data[6]);
      var dateDebut = data[7];
      var heureDebut = getSafeValue(data[8]);
      var dateFin = data[9];
      var heureFin = getSafeValue(data[10]);
      var motifAbsence = getSafeValue(data[11]);
      var typePermission = getSafeValue(data[12]);
      var motifPermission = getSafeValue(data[13]);
      var nbreJours = getSafeValue(data[14]);
      var emailSuperieur = getSafeValue(data[15]);
  
      Logger.log(`[FORMULAIRE] ${nom} ${prenom} | ${email}`);
  
      // ÉTAPE 2: Validations essentielles
      if (!email || !email.includes("@")) {
        Logger.log("[ERREUR] Email demandeur invalide");
        return;
      }
      if (!emailSuperieur || !emailSuperieur.includes("@")) {
        Logger.log("[ERREUR] Email supérieur invalide");
        return;
      }
      if (!dateDebut) {
        Logger.log("[ERREUR] Date de début manquante");
        return;
      }
  
      // ÉTAPE 3: Organisation des dossiers
      var demandeYear = extractYear(dateDebut);
      var employeeFolderName = `${nom || "Nom"} ${prenom || "Prénom"}`;
  
      // Récupération des dossiers Google Drive
      var mainFolder = DriveApp.getFolderById(CONFIG.folders.mainId);
      var employeeFolder = DriveApp.getFolderById(CONFIG.folders.employeeId);
      var yearFolder = getOrCreateFolder(employeeFolder, `Autorisation d'absence - ${demandeYear}`);
      var demandeurFolder = getOrCreateFolder(yearFolder, employeeFolderName);
      var pendingFolder = getOrCreateFolder(demandeurFolder, "En Attente");
  
      if (!pendingFolder) {
        Logger.log("[ERREUR] Impossible de créer le dossier En Attente");
        return;
      }
  
      // ÉTAPE 4: Création du document depuis le template
      var tempDoc = DriveApp.getFileById(CONFIG.folders.templateId).makeCopy(
        `Demande_${prenom}_${nom}_${getFormattedDateTime().replace(/[/:]/g, "-")}`,
        pendingFolder
      );
      var tempDocId = tempDoc.getId();
      var doc = DocumentApp.openById(tempDocId);
      var body = doc.getBody();
  
      // ÉTAPE 5: Remplacement des placeholders dans le document
      var dateDebutStr = dateDebut instanceof Date ? dateDebut.toLocaleDateString("fr-FR") : getSafeValue(dateDebut);
      var dateFinStr = dateFin instanceof Date ? dateFin.toLocaleDateString("fr-FR") : getSafeValue(dateFin);
  
      // Dictionnaire des remplacements
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
        "{{statusSuperieurHierarchique}}": "En attente",
        "{{statusRH}}": "En attente", 
        "{{statusPresidence}}": "En attente"
      };
  
      // Application des remplacements
      for (var key in placeholders) {
        try {
          if (body.findText(key)) {
            body.replaceText(key, placeholders[key]);
          }
        } catch (error) {
          Logger.log(`[ERREUR REMPLACEMENT] ${key}: ${error.message}`);
        }
      }
  
      doc.saveAndClose();
      Logger.log(`[DOCUMENT CRÉÉ] ID : ${tempDocId}`);
  
      // ÉTAPE 6: Génération des métadonnées cachées
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var lastRow = sheet.getLastRow();
      generateBackgroundMetadata(sheet, lastRow, tempDocId);
  
      // ÉTAPE 7: Envoi des notifications
      var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
      var nomComplet = `${prenom} ${nom}`;
      
      // Notification au supérieur hiérarchique
      EmailService.notifyValidator(emailSuperieur, "Supérieur Hiérarchique", nomComplet, tempDoc.getUrl(), sheetUrl);
      
      // Confirmation au demandeur
      EmailService.confirmSubmission(email, nomComplet, tempDoc.getUrl());
  
      Logger.log("=== FIN TRAITEMENT FORMULAIRE ===");
  
    } catch (error) {
      Logger.log(`[ERREUR CRITIQUE onFormSubmit] ${error.message}`);
    }
  }
  
  // VALIDATION - Gestion des modifications dans les colonnes Q, R, S
  function onEditValidation(e) {
    try {
      var sheet = e.source.getActiveSheet();
      var range = e.range;
  
      // Vérifie si c'est une colonne de validation (Q=17, R=18, S=19)
      if (range.getColumn() >= 17 && range.getColumn() <= 19) {
        var row = range.getRow();
        var colonne = range.getColumn();
        var newValue = range.getValue();
        var oldValue = e.oldValue;
  
        // Ignore la ligne d'en-tête
        if (row === 1) return;
  
        // ÉTAPE 1: Détermine le rôle selon la colonne
        var roles = {17: "Supérieur Hiérarchique", 18: "RH", 19: "Présidence"};
        var role = roles[colonne];
  
        Logger.log(`[VALIDATION] ${role} - Valeur : ${newValue}`);
  
        // ÉTAPE 2: Empêche la modification d'une décision déjà prise
        if (oldValue && oldValue !== "" && oldValue !== "En attente") {
          range.setValue(oldValue);
          SpreadsheetApp.getUi().alert("Décision déjà prise et non modifiable.");
          return;
        }
  
        // ÉTAPE 3: Récupère l'ID du document depuis les métadonnées
        var docId = sheet.getRange(row, 21).getValue(); // Colonne U
        
        if (!docId) {
          Logger.log("[ERREUR] ID document manquant");
          return;
        }
  
        // ÉTAPE 4: Vérifie l'ordre de validation (cascade)
        if (!checkValidationOrder(sheet, row, colonne)) {
          range.setValue("En attente");
          SpreadsheetApp.getUi().alert("Ordre de validation : Supérieur → RH → Présidence");
          return;
        }
  
        // ÉTAPE 5: Valide le statut saisi
        if (newValue !== "Favorable" && newValue !== "Non Favorable") {
          range.setValue("En attente");
          SpreadsheetApp.getUi().alert("Statut autorisé : 'Favorable' ou 'Non Favorable'");
          return;
        }
  
        // ÉTAPE 6: Traite selon le statut
        var decisionTime = getFormattedDateTime();
        
        if (newValue === "Non Favorable") {
          handleRejection(sheet, row, role, docId, decisionTime);
        } else {
          handleApproval(sheet, row, role, docId, decisionTime, colonne);
        }
  
        // ÉTAPE 7: Protège la cellule pour éviter modification
        protectValidationCell(sheet, row, colonne);
      }
  
    } catch (error) {
      Logger.log(`[ERREUR onEditValidation] ${error.message}`);
    }
  }
  
  // CONTRÔLES - Vérifie l'ordre de validation en cascade
  function checkValidationOrder(sheet, row, colonne) {
    try {
      // Supérieur Hiérarchique (colonne 17) toujours autorisé
      if (colonne == 17) return true;
  
      // RH et Présidence doivent attendre validation du supérieur
      var statusSuperieur = sheet.getRange(row, 17).getValue();
      if (statusSuperieur !== "Favorable") return false;
  
      // RH (colonne 18) autorisé si supérieur OK
      if (colonne == 18) return true;
  
      // Présidence (colonne 19) autorisé si RH OK
      if (colonne == 19) {
        var statusRH = sheet.getRange(row, 18).getValue();
        return statusRH === "Favorable";
      }
  
      return false;
    } catch (error) {
      Logger.log(`[ERREUR checkValidationOrder] ${error.message}`);
      return false;
    }
  }
  
  // REJET - Traite les demandes refusées
  function handleRejection(sheet, row, role, docId, decisionTime) {
    try {
      Logger.log(`[REJET] Par ${role} à ${decisionTime}`);
  
      // Désactive les validations suivantes avec N/A
      if (role === "Supérieur Hiérarchique") {
        sheet.getRange(row, 18).setValue("N/A");  // RH
        sheet.getRange(row, 19).setValue("N/A");  // Présidence
      } else if (role === "RH") {
        sheet.getRange(row, 19).setValue("N/A");  // Présidence
      }
  
      // Notifie le demandeur du rejet
      var demandeurEmail = sheet.getRange(row, 2).getValue();
      var demandeurNom = `${sheet.getRange(row, 4).getValue()} ${sheet.getRange(row, 5).getValue()}`;
      var doc = DriveApp.getFileById(docId);
      
      EmailService.notifyRejection(demandeurEmail, demandeurNom, doc.getUrl(), role);
  
      // Déplace le document vers "Rejetée"
      moveDocumentToFolder(doc, "Rejetée");
  
      return true;
    } catch (error) {
      Logger.log(`[ERREUR handleRejection] ${error.message}`);
      return false;
    }
  }
  
  // APPROBATION - Traite les demandes approuvées
  function handleApproval(sheet, row, role, docId, decisionTime, colonne) {
    try {
      Logger.log(`[APPROBATION] Par ${role} à ${decisionTime}`);
  
      // Si c'est la dernière validation (Présidence = colonne 19)
      if (colonne == 19) {
        var demandeurEmail = sheet.getRange(row, 2).getValue();
        var demandeurNom = `${sheet.getRange(row, 4).getValue()} ${sheet.getRange(row, 5).getValue()}`;
        var doc = DriveApp.getFileById(docId);
        
        // Notifie l'approbation finale
        EmailService.notifyApproval(demandeurEmail, demandeurNom, doc.getUrl(), role);
  
        // Déplace vers "Validée"
        moveDocumentToFolder(doc, "Validée");
      } else {
        // Notifie le prochain validateur dans la chaîne
        EmailService.notifyNextValidator(sheet, row, colonne, docId);
      }
  
      return true;
    } catch (error) {
      Logger.log(`[ERREUR handleApproval] ${error.message}`);
      return false;
    }
  }
  
  // INITIALISATION - Configure le système (exécuter une seule fois)
  function initializeSystem() {
    try {
      Logger.log("=== INITIALISATION SYSTÈME ===");
      
      // ÉTAPE 1: Configure la protection des colonnes
      setupSheetProtection();
      
      // ÉTAPE 2: Supprime les anciens triggers
      var triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(function(trigger) {
        if (trigger.getHandlerFunction() === 'onEditValidation') {
          ScriptApp.deleteTrigger(trigger);
        }
      });
      
      // ÉTAPE 3: Crée le trigger pour les modifications
      ScriptApp.newTrigger('onEditValidation')
        .onEdit()
        .create();
      
      Logger.log("[INITIALISATION] Système configuré avec succès");
      
    } catch (error) {
      Logger.log(`[ERREUR INITIALISATION] ${error.message}`);
    }
  }