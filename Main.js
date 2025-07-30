// WORKFLOW PRINCIPAL - Gestion des demandes d'absence PE-A/PE-B/PO

// METADONNEES - Genere les donnees cachees en arriere-plan
function generateBackgroundMetadata(sheet, row, docId, extractedData) {
  try {
    // Recupere les emails necessaires selon le type PE-A/PE-B/PO
    var emailSuperieur = extractedData.emailSuperieur;
    var emailRH = CONFIG.emails.rh;
    var emailPresidence = getPresidenceEmail(emailSuperieur);

    // Enregistre dans les colonnes de metadonnees (AG, AH, AI, AJ, AK)
    var metaCols = CONFIG.columns.metadata;
    sheet.getRange(row, metaCols.DOC_ID + 1).setValue(docId);                    // AG: ID document
    sheet.getRange(row, metaCols.EMAIL_RH + 1).setValue(emailRH);              // AH: Email RH  
    sheet.getRange(row, metaCols.EMAIL_PRESIDENCE + 1).setValue(emailPresidence); // AI: Email Presidence
    sheet.getRange(row, metaCols.TIMESTAMP + 1).setValue(getFormattedDateTime()); // AJ: Timestamp
    sheet.getRange(row, metaCols.STATUS + 1).setValue("Cree");               // AK: Statut
    
    Logger.log("METADONNEES Generees pour ligne " + row + " - Type: " + extractedData.demandType);
    return true;
  } catch (error) {
    Logger.log("ERREUR METADONNEES " + error.message);
    return false;
  }
}

// SOUMISSION FORMULAIRE - Point d'entree principal PE-A/PE-B/PO
function onFormSubmit(e) {
  try {
    Logger.log("DEBUT TRAITEMENT FORMULAIRE PE-A/PE-B/PO");
    
    var data = e.values || [];
    
    // ETAPE 1: Detection automatique du type de demande
    var demandType = detectDemandType(data);
    Logger.log("DETECTION Type detecte : " + demandType.type + " - " + demandType.description);
    
    // ETAPE 2: Extraction des donnees selon le type detecte
    var extractedData = extractDemandData(data, demandType);
    
    if (!extractedData) {
      Logger.log("ERREUR Impossible d'extraire les donnees du formulaire");
      return;
    }
    
    Logger.log("FORMULAIRE " + extractedData.nom + " " + extractedData.prenom + " | " + extractedData.email + " | Type: " + extractedData.demandType);

    // ETAPE 3: Validations essentielles
    if (!extractedData.email || !extractedData.email.includes("@")) {
      Logger.log("ERREUR Email demandeur invalide");
      return;
    }
    if (!extractedData.emailSuperieur || !extractedData.emailSuperieur.includes("@")) {
      Logger.log("ERREUR Email superieur invalide");
      return;
    }
    if (!extractedData.dateDebut) {
      Logger.log("ERREUR Date de debut manquante");
      return;
    }

    // ETAPE 4: Organisation des dossiers
    var demandeYear = extractYear(extractedData.dateDebut);
    var employeeFolderName = (extractedData.nom || "Nom") + " " + (extractedData.prenom || "Prenom");

    // Recuperation des dossiers Google Drive
    var mainFolder = DriveApp.getFolderById(CONFIG.folders.mainId);
    var employeeFolder = DriveApp.getFolderById(CONFIG.folders.employeeId);
    var yearFolder = getOrCreateFolder(employeeFolder, "Autorisation d'absence - " + demandeYear);
    var demandeurFolder = getOrCreateFolder(yearFolder, employeeFolderName);
    var pendingFolder = getOrCreateFolder(demandeurFolder, "En Attente");

    if (!pendingFolder) {
      Logger.log("ERREUR Impossible de creer le dossier En Attente");
      return;
    }

    // ETAPE 5: Creation du document depuis le template
    var tempDoc = DriveApp.getFileById(CONFIG.folders.templateId).makeCopy(
      "Demande_" + extractedData.demandType + "_" + extractedData.prenom + "_" + extractedData.nom + "_" + getFormattedDateTime().replace(/[/:]/g, "-"),
      pendingFolder
    );
    var tempDocId = tempDoc.getId();
    var doc = DocumentApp.openById(tempDocId);
    var body = doc.getBody();

    // ETAPE 6: Remplacement des placeholders dans le document
    var dateDebutStr = extractedData.dateDebut instanceof Date ? 
      extractedData.dateDebut.toLocaleDateString("fr-FR") : 
      getSafeValue(extractedData.dateDebut);
    var dateFinStr = extractedData.dateFin instanceof Date ? 
      extractedData.dateFin.toLocaleDateString("fr-FR") : 
      getSafeValue(extractedData.dateFin);

    // Dictionnaire des remplacements avec donnees PE-A/PE-B/PO
    var placeholders = {
      "{{matricule}}": extractedData.matricule,
      "{{nom}}": extractedData.nom,
      "{{prenom}}": extractedData.prenom,
      "{{serviceFonction}}": extractedData.serviceFonction,
      "{{typeAbsence}}": extractedData.typePermission,
      "{{typeSpecifique}}": extractedData.demandType,
      "{{dateDebut}}": dateDebutStr,
      "{{heureDebut}}": extractedData.heureDebut || "",
      "{{dateFin}}": dateFinStr,
      "{{heureFin}}": extractedData.heureFin || "",
      "{{motifAbsence}}": extractedData.motifDetail || extractedData.motifSelection,
      "{{motifSelection}}": extractedData.motifSelection,
      "{{motifDetail}}": extractedData.motifDetail || "",
      "{{typePermission}}": extractedData.typePermission,
      "{{motifPermission}}": extractedData.motifDetail || extractedData.motifSelection,
      "{{nbreJours}}": extractedData.nbreJours,
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
        Logger.log("ERREUR REMPLACEMENT " + key + ": " + error.message);
      }
    }

    doc.saveAndClose();
    Logger.log("DOCUMENT CREE ID : " + tempDocId + " | Type: " + extractedData.demandType);

    // ETAPE 7: Generation des metadonnees cachees
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    generateBackgroundMetadata(sheet, lastRow, tempDocId, extractedData);

    // ETAPE 8: Envoi des notifications avec donnees PE-A/PE-B/PO
    var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    var nomComplet = extractedData.prenom + " " + extractedData.nom;
    
    // Prepare les donnees de la demande pour l'email
    var demandData = {
      demandType: extractedData.demandType,
      typePermission: extractedData.typePermission,
      nbreJours: extractedData.nbreJours || "Non specifie",
      dateDebut: dateDebutStr,
      heureDebut: extractedData.heureDebut || "",
      dateFin: dateFinStr,
      heureFin: extractedData.heureFin || "",
      motifDetail: extractedData.motifDetail || extractedData.motifSelection || "Non specifie",
      motifSelection: extractedData.motifSelection,
      serviceFonction: extractedData.serviceFonction || "",
      matricule: extractedData.matricule || ""
    };
    
    // Notification au superieur hierarchique avec informations completes PE-A/PE-B/PO
    EmailService.notifyValidator(extractedData.emailSuperieur, "Superieur Hierarchique", nomComplet, tempDoc.getUrl(), sheetUrl, demandData);
    
    // Confirmation au demandeur avec recapitulatif PE-A/PE-B/PO
    EmailService.confirmSubmission(extractedData.email, nomComplet, tempDoc.getUrl(), demandData);

    Logger.log("FIN TRAITEMENT FORMULAIRE " + extractedData.demandType);

  } catch (error) {
    Logger.log("ERREUR CRITIQUE onFormSubmit " + error.message);
  }
}

// VALIDATION - Gestion des modifications dans les colonnes AC, AD, AE, AF
function onEditValidation(e) {
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;

    // Verifie si c'est une colonne de validation (AC=29, AD=30, AE=31, AF=32)
    var validationCols = CONFIG.columns.validation;
    var colonne = range.getColumn();
    
    if (colonne >= (validationCols.SUPERIEUR + 1) && colonne <= (validationCols.COMMENTAIRE + 1)) {
      var row = range.getRow();
      var newValue = range.getValue();
      var oldValue = e.oldValue;

      // Ignore la ligne d'en-tete
      if (row === 1) return;

      // Ignore la colonne commentaire (AF) - elle est libre
      if (colonne === (validationCols.COMMENTAIRE + 1)) {
        Logger.log("COMMENTAIRE Ligne " + row + " : " + newValue);
        return;
      }

      // ETAPE 1: Determine le role selon la colonne
      var roles = {};
      roles[validationCols.SUPERIEUR + 1] = "Superieur Hierarchique";
      roles[validationCols.RH + 1] = "RH";
      roles[validationCols.PRESIDENCE + 1] = "Presidence";
      
      var role = roles[colonne];

      Logger.log("VALIDATION " + role + " - Ligne " + row + " - Valeur : " + newValue);

      // ETAPE 2: Empeche la modification d'une decision deja prise
      if (oldValue && oldValue !== "" && oldValue !== "En attente") {
        range.setValue(oldValue);
        SpreadsheetApp.getUi().alert("Decision deja prise et non modifiable.");
        return;
      }

      // ETAPE 3: Recupere l'ID du document depuis les metadonnees
      var docId = sheet.getRange(row, CONFIG.columns.metadata.DOC_ID + 1).getValue(); // Colonne AG
      
      if (!docId) {
        Logger.log("ERREUR ID document manquant en colonne " + String.fromCharCode(65 + CONFIG.columns.metadata.DOC_ID));
        return;
      }

      // ETAPE 4: Verifie l'ordre de validation (cascade)
      if (!checkValidationOrder(sheet, row, colonne - 1)) {  // -1 car colonne est basee sur 1
        range.setValue("En attente");
        SpreadsheetApp.getUi().alert("Ordre de validation : Superieur -> RH -> Presidence");
        return;
      }

      // ETAPE 5: Valide le statut saisi
      if (newValue !== "Favorable" && newValue !== "Non Favorable") {
        range.setValue("En attente");
        SpreadsheetApp.getUi().alert("Statut autorise : 'Favorable' ou 'Non Favorable'");
        return;
      }

      // ETAPE 6: Traite selon le statut
      var decisionTime = getFormattedDateTime();
      
      if (newValue === "Non Favorable") {
        handleRejection(sheet, row, role, docId, decisionTime);
      } else {
        handleApproval(sheet, row, role, docId, decisionTime, colonne - 1);  // -1 car colonne est basee sur 1
      }

      // ETAPE 7: Protege la cellule pour eviter modification
      protectValidationCell(sheet, row, colonne);
    }

  } catch (error) {
    Logger.log("ERREUR onEditValidation " + error.message);
  }
}

// CONTROLES - Verifie l'ordre de validation en cascade PE-A/PE-B/PO
function checkValidationOrder(sheet, row, colonne) {
  try {
    var validationCols = CONFIG.columns.validation;
    
    // Superieur Hierarchique toujours autorise
    if (colonne == validationCols.SUPERIEUR) return true;

    // RH et Presidence doivent attendre validation du superieur
    var statusSuperieur = sheet.getRange(row, validationCols.SUPERIEUR + 1).getValue();
    if (statusSuperieur !== "Favorable") return false;

    // RH autorise si superieur OK
    if (colonne == validationCols.RH) return true;

    // Presidence autorise si RH OK
    if (colonne == validationCols.PRESIDENCE) {
      var statusRH = sheet.getRange(row, validationCols.RH + 1).getValue();
      return statusRH === "Favorable";
    }

    return false;
  } catch (error) {
    Logger.log("ERREUR checkValidationOrder " + error.message);
    return false;
  }
}

// REJET - Traite les demandes refusees PE-A/PE-B/PO
function handleRejection(sheet, row, role, docId, decisionTime) {
  try {
    Logger.log("REJET Par " + role + " a " + decisionTime);

    var validationCols = CONFIG.columns.validation;

    // Desactive les validations suivantes avec N/A
    if (role === "Superieur Hierarchique") {
      sheet.getRange(row, validationCols.RH + 1).setValue("N/A");      // AD: RH
      sheet.getRange(row, validationCols.PRESIDENCE + 1).setValue("N/A");  // AE: Presidence
    } else if (role === "RH") {
      sheet.getRange(row, validationCols.PRESIDENCE + 1).setValue("N/A");  // AE: Presidence
    }

    // Met a jour le statut global
    sheet.getRange(row, CONFIG.columns.metadata.STATUS + 1).setValue("Rejete par " + role);

    // Recupere les donnees de la demande pour le recapitulatif email PE-A/PE-B/PO
    var demandData = EmailService.getDemandDataFromSheet(sheet, row);

    // Notifie le demandeur du rejet avec recapitulatif
    var demandeurEmail = sheet.getRange(row, CONFIG.columns.common.EMAIL + 1).getValue();
    var demandeurNom = sheet.getRange(row, CONFIG.columns.common.NOM + 1).getValue() + " " + sheet.getRange(row, CONFIG.columns.common.PRENOM + 1).getValue();
    var doc = DriveApp.getFileById(docId);
    
    EmailService.notifyRejection(demandeurEmail, demandeurNom, doc.getUrl(), role, demandData);

    // Deplace le document vers "Rejetee"
    moveDocumentToFolder(doc, "Rejetee");

    return true;
  } catch (error) {
    Logger.log("ERREUR handleRejection " + error.message);
    return false;
  }
}

// APPROBATION - Traite les demandes approuvees PE-A/PE-B/PO
function handleApproval(sheet, row, role, docId, decisionTime, colonne) {
  try {
    Logger.log("APPROBATION Par " + role + " a " + decisionTime);

    var validationCols = CONFIG.columns.validation;

    // Si c'est la derniere validation (Presidence)
    if (colonne == validationCols.PRESIDENCE) {
      // Met a jour le statut global
      sheet.getRange(row, CONFIG.columns.metadata.STATUS + 1).setValue("Approuve");

      // Recupere les donnees de la demande PE-A/PE-B/PO
      var demandData = EmailService.getDemandDataFromSheet(sheet, row);

      var demandeurEmail = sheet.getRange(row, CONFIG.columns.common.EMAIL + 1).getValue();
      var demandeurNom = sheet.getRange(row, CONFIG.columns.common.NOM + 1).getValue() + " " + sheet.getRange(row, CONFIG.columns.common.PRENOM + 1).getValue();
      var doc = DriveApp.getFileById(docId);
      
      // Notifie l'approbation finale avec recapitulatif PE-A/PE-B/PO
      EmailService.notifyApproval(demandeurEmail, demandeurNom, doc.getUrl(), role, demandData);

      // Deplace vers "Validee"
      moveDocumentToFolder(doc, "Validee");
    } else {
      // Met a jour le statut intermediaire
      sheet.getRange(row, CONFIG.columns.metadata.STATUS + 1).setValue("En cours - " + role + " valide");

      // Notifie le prochain validateur dans la chaine (avec donnees PE-A/PE-B/PO integrees)
      EmailService.notifyNextValidator(sheet, row, colonne, docId);
    }

    return true;
  } catch (error) {
    Logger.log("ERREUR handleApproval " + error.message);
    return false;
  }
}

// INITIALISATION - Configure le systeme PE-A/PE-B/PO (executer une seule fois)
function initializeSystem() {
  try {
    Logger.log("INITIALISATION SYSTEME PE-A/PE-B/PO");
    
    // ETAPE 1: Configure la protection des colonnes (AC, AD, AE, AF)
    setupSheetProtection();
    
    // ETAPE 2: Supprime les anciens triggers
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'onEditValidation') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    // ETAPE 3: Cree le trigger pour les modifications
    try {
      ScriptApp.newTrigger('onEditValidation').onEdit().create();
      Logger.log("Trigger cree avec succes");
    } catch (triggerError) {
      Logger.log("ERREUR Trigger: " + triggerError.message);
      Logger.log("Veuillez creer manuellement un trigger onEdit pour la fonction onEditValidation");
    }
    
    Logger.log("INITIALISATION Systeme configure avec succes");
    
  } catch (error) {
    Logger.log("ERREUR INITIALISATION " + error.message);
  }
}

// FONCTIONS DE DEBOGAGE

// Verifier si le trigger existe
function checkTriggers() {
  try {
    Logger.log("VERIFICATION TRIGGERS");
    var triggers = ScriptApp.getProjectTriggers();
    Logger.log("Nombre total de triggers: " + triggers.length);
    
    for (var i = 0; i < triggers.length; i++) {
      var trigger = triggers[i];
      Logger.log("Trigger " + (i+1) + ":");
      Logger.log("  - Fonction: " + trigger.getHandlerFunction());
      Logger.log("  - Type: " + trigger.getEventType());
    }
    
    // Cherche specifiquement le trigger onEditValidation
    var hasEditTrigger = false;
    for (var j = 0; j < triggers.length; j++) {
      if (triggers[j].getHandlerFunction() === 'onEditValidation') {
        hasEditTrigger = true;
        Logger.log("TRIGGER onEditValidation TROUVE !");
        break;
      }
    }
    
    if (!hasEditTrigger) {
      Logger.log("PROBLEME: Trigger onEditValidation MANQUANT !");
      Logger.log("Solution: Executez createTriggerManually()");
    }
    
    Logger.log("FIN VERIFICATION");
    
  } catch (error) {
    Logger.log("ERREUR checkTriggers: " + error.message);
  }
}

// Creer le trigger manuellement
function createTriggerManually() {
  try {
    Logger.log("CREATION TRIGGER MANUEL");
    
    // Supprime les anciens triggers onEditValidation
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'onEditValidation') {
        ScriptApp.deleteTrigger(triggers[i]);
        Logger.log("Ancien trigger supprime");
      }
    }
    
    // Cree le nouveau trigger
    var trigger = ScriptApp.newTrigger('onEditValidation').onEdit().create();
    
    Logger.log("TRIGGER CREE AVEC SUCCES !");
    Logger.log("ID du trigger: " + trigger.getUniqueId());
    
    Logger.log("FIN CREATION");
    
  } catch (error) {
    Logger.log("ERREUR createTriggerManually: " + error.message);
  }
}

// Tester la detection des colonnes
function testColumnDetection() {
  try {
    Logger.log("TEST DETECTION COLONNES");
    
    var validationCols = CONFIG.columns.validation;
    Logger.log("Colonnes de validation configurees:");
    Logger.log("  - SUPERIEUR: " + validationCols.SUPERIEUR + " (Colonne " + String.fromCharCode(65 + validationCols.SUPERIEUR) + ")");
    Logger.log("  - RH: " + validationCols.RH + " (Colonne " + String.fromCharCode(65 + validationCols.RH) + ")");
    Logger.log("  - PRESIDENCE: " + validationCols.PRESIDENCE + " (Colonne " + String.fromCharCode(65 + validationCols.PRESIDENCE) + ")");
    Logger.log("  - COMMENTAIRE: " + validationCols.COMMENTAIRE + " (Colonne " + String.fromCharCode(65 + validationCols.COMMENTAIRE) + ")");
    
    // Test des plages
    Logger.log("Plage de detection: Colonnes " + (validationCols.SUPERIEUR + 1) + " a " + (validationCols.COMMENTAIRE + 1));
    Logger.log("Soit: " + String.fromCharCode(65 + validationCols.SUPERIEUR) + " a " + String.fromCharCode(65 + validationCols.COMMENTAIRE));
    
    Logger.log("FIN TEST");
    
  } catch (error) {
    Logger.log("ERREUR testColumnDetection: " + error.message);
  }
}