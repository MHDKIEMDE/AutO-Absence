// CONFIGURATION PRINCIPALE - Modifier selon vos besoins
const CONFIG = {
    emails: {
      rh: "",                    // Email du service RH (ex: "rh@entreprise.com")
      presidence: ["", ""]       // Emails de la présidence - 2 maximum (ex: ["president@entreprise.com", "directeur@entreprise.com"])
    },
    folders: {
      mainId: "180NyuZK0VMhF3_3-S9zxFLya7t-cT5Bb",      // ID du dossier principal Google Drive - À récupérer dans l'URL
      employeeId: "1bw_8yRTU-CHT-S9KcBSP6V0Jmh5BP5dg",  // ID du dossier des employés - Sous-dossier du principal
      templateId: "15fZybSZZ4dxr0aUOnvQZ0kgnzvoWmasAPD5Pd37xdNY" // ID du template de document - Document modèle à dupliquer
    }
  };
  
  // UTILITAIRES - Fonctions réutilisables pour la sécurité et formatage
  
  // Sécurise les valeurs du formulaire (évite null/undefined)
  function getSafeValue(value, defaultValue = "") {
    return value !== null && value !== undefined
      ? value.toString().trim()
      : defaultValue;
  }
  
  // Extrait l'année d'une date pour organiser les dossiers par année
  function extractYear(dateValue) {
    try {
      if (!dateValue) return new Date().getFullYear().toString();
      if (dateValue instanceof Date) return dateValue.getFullYear().toString();
      if (typeof dateValue === "string" && dateValue.includes("/")) {
        const dateParts = dateValue.split("/");
        return dateParts.length >= 3 ? dateParts[2] : new Date().getFullYear().toString();
      }
      return new Date().getFullYear().toString();
    } catch (error) {
      Logger.log(`[ERREUR DATE] ${error.message}`);
      return new Date().getFullYear().toString();
    }
  }
  
  // Crée ou récupère un dossier Google Drive (évite les doublons)
  function getOrCreateFolder(parentFolder, folderName) {
    try {
      // Génère un nom par défaut si le nom est invalide
      if (!folderName || folderName.trim() === "" || folderName.includes("undefined")) {
        folderName = "Dossier_" + new Date().getTime();
      }
      
      // Cherche le dossier existant ou le crée
      var folders = parentFolder.getFoldersByName(folderName);
      var folder = folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
      return folder;
    } catch (error) {
      Logger.log(`[ERREUR DOSSIER] ${error.message}`);
      return null;
    }
  }
  
  // Formate la date et heure actuelles pour les logs et horodatage
  function getFormattedDateTime() {
    var now = new Date();
    var day = ("0" + now.getDate()).slice(-2);        // Force 2 chiffres (01, 02...)
    var month = ("0" + (now.getMonth() + 1)).slice(-2); // +1 car les mois commencent à 0
    var year = now.getFullYear();
    var hours = ("0" + now.getHours()).slice(-2);
    var minutes = ("0" + now.getMinutes()).slice(-2);
    return `${day}/${month}/${year} ${hours}:${minutes}`;
  }
  
  // Détermine l'email de présidence selon le supérieur hiérarchique
  function getPresidenceEmail(superieurEmail) {
    try {
      let presEmails = CONFIG.emails.presidence;
      if (!presEmails || presEmails.length === 0) return null;
      
      // Si le supérieur fait déjà partie de la présidence, utilise l'autre email
      if (presEmails.includes(superieurEmail)) {
        const otherEmail = presEmails.filter(email => email !== superieurEmail)[0];
        return otherEmail || presEmails[0];
      }
      // Sinon, utilise le premier email de présidence
      return presEmails[0];
    } catch (error) {
      Logger.log(`[ERREUR getPresidenceEmail] ${error.message}`);
      return CONFIG.emails.presidence && CONFIG.emails.presidence.length > 0 
        ? CONFIG.emails.presidence[0] : null;
    }
  }
  
  // PROTECTION - Sécurise la feuille de calcul contre les modifications non autorisées
  
  // Protège la feuille sauf les colonnes Q, R, S (colonnes de validation)
  function setupSheetProtection() {
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      
      // Supprime toutes les protections existantes pour éviter les conflits
      var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      protections.forEach(function(protection) {
        protection.remove();
      });
  
      // Protège toute la feuille avec une description
      var protection = sheet.protect().setDescription('Protection générale');
      
      // Définit les colonnes non protégées (seules modifiables)
      var unprotectedRanges = [
        sheet.getRange('Q:Q'),  // Colonne Q = Supérieur Hiérarchique (colonne 17)
        sheet.getRange('R:R'),  // Colonne R = RH (colonne 18)
        sheet.getRange('S:S')   // Colonne S = Présidence (colonne 19)
      ];
      protection.setUnprotectedRanges(unprotectedRanges);
      
      // Gère les permissions utilisateur (optionnel selon votre configuration)
      try {
        var me = Session.getEffectiveUser();
        if (me) {
          protection.removeEditors(protection.getEditors());
          protection.addEditor(me);
        }
      } catch (permError) {
        Logger.log(`[AVERTISSEMENT] Gestion permissions : ${permError.message}`);
      }
  
      return true;
    } catch (error) {
      Logger.log(`[ERREUR PROTECTION] ${error.message}`);
      return false;
    }
  }
  
  // Protège une cellule individuelle après validation (empêche modification)
  function protectValidationCell(sheet, row, colonne) {
    try {
      var protection = sheet.getRange(row, colonne).protect();
      protection.setDescription(`Décision validée - L${row}C${colonne}`);
      return true;
    } catch (error) {
      Logger.log(`[ERREUR PROTECTION CELLULE] ${error.message}`);
      return false;
    }
  }
  
  // GESTION DOCUMENTS - Organise les fichiers selon leur statut de validation
  
  // Déplace un document vers le dossier approprié (Validée/Rejetée)
  function moveDocumentToFolder(doc, targetFolderName) {
    try {
      var parentFolders = doc.getParents();
      if (parentFolders.hasNext()) {
        var currentFolder = parentFolders.next();        // Dossier actuel du document
        var parentFolder = currentFolder.getParents().next(); // Dossier parent (contient Validée/Rejetée)
  
        // Cherche le dossier cible dans le même niveau que le dossier actuel
        var folders = parentFolder.getFolders();
        while (folders.hasNext()) {
          var folder = folders.next();
          if (folder.getName() === targetFolderName) {
            currentFolder.removeFile(doc);  // Retire du dossier "En Attente"
            folder.addFile(doc);             // Ajoute au dossier final (Validée/Rejetée)
            Logger.log(`[DOCUMENT DÉPLACÉ] Vers ${targetFolderName}`);
            break;
          }
        }
      }
    } catch (error) {
      Logger.log(`[ERREUR DÉPLACEMENT] ${error.message}`);
    }
  }