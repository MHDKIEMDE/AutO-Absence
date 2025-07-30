/**
 * Configuration centrale pour le système de gestion des demandes d'absence
 * Formulaire Massaka - Version PE-A/PE-B/PO
 */

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
  },
  
  // MAPPINGS COLONNES RÉELLES DE VOTRE SHEET
  columns: {
    // Colonnes communes (A-G)
    common: {
      HORODATEUR: 0,        // A - Horodateur
      EMAIL: 1,             // B - Adresse e-mail  
      MATRICULE: 2,         // C - Matricule
      NOM: 3,               // D - Nom
      PRENOM: 4,            // E - Prénom
      SERVICE: 5,           // F - Service / Fonction
      TYPE_PERMISSION: 6    // G - Type de permission
    },
    
    // Permission Exceptionnelle A (PE-A) - Colonnes H-N
    peA: {
      MOTIF_SELECTION: 7,     // H - Motif d'absence
      MOTIF_DETAIL: 8,        // I - Motif d'absence (PE-A)
      NB_JOURS: 9,            // J - Nombres de jours sollicités (PE-A)
      DATE_DEBUT: 10,         // K - Date du début (PE-A)
      HEURE_DEBUT: 11,        // L - Heurs du début (PE-A)
      DATE_FIN: 12,           // M - Date de fin (PE-A)
      HEURE_FIN: 13,          // N - Heurs du fin (PE-A)
      EMAIL_SUPERIEUR: 14     // O - Adresse e-mail de votre supérieur hiérarchique (PE-A)
    },
    
    // Permission Exceptionnelle B (PE-B) - Colonnes O-T
    peB: {
      MOTIF_SELECTION: 7,     // H - Motif d'absence (même colonne)
      NB_JOURS: 15,           // P - nombres de jours sollicités (PE-B)
      DATE_DEBUT: 16,         // Q - Date du début (PE-B)
      HEURE_DEBUT: 17,        // R - Heurs du début (PE-B)
      DATE_FIN: 18,           // S - Date de fin (PE-B)
      HEURE_FIN: 19,          // T - Heurs du fin (PE-B)
      EMAIL_SUPERIEUR: 20     // U - Adresse e-mail de votre supérieur hiérarchique (PE-B)
    },
    
    // Permission Ordinaire (PO) - Colonnes U-AA
    po: {
      MOTIF: 21,              // V - Motif d'absence (PO)
      NB_JOURS: 22,           // W - Nombres de jours sollicités (PO)
      DATE_DEBUT: 23,         // X - Date du début (PO)
      HEURE_DEBUT: 24,        // Y - Heurs du début (PO)
      DATE_FIN: 25,           // Z - Date de fin (PO)
      HEURE_FIN: 26,          // AA - Heurs du fin (PO)
      EMAIL_SUPERIEUR: 27     // AB - Adresse e-mail de votre supérieur hiérarchique (PO)
    },
    
    // Colonnes de validation
    validation: {
      SUPERIEUR: 28,          // AC - Avis du supérieur hiérarchique
      RH: 29,                 // AD - Avis des ressources humaines RH
      PRESIDENCE: 30,         // AE - Avis de la présidence
      COMMENTAIRE: 31         // AF - Commentaire de la présidence
    },
    
    // Métadonnées cachées (colonnes après validation)
    metadata: {
      DOC_ID: 32,             // AG - ID du document Google Doc
      EMAIL_RH: 33,           // AH - Email RH (calculé)
      EMAIL_PRESIDENCE: 34,   // AI - Email Présidence (calculé)
      TIMESTAMP: 35,          // AJ - Horodatage création
      STATUS: 36              // AK - Statut global
    }
  },
  
  // Motifs PE-A (nécessitent plus de détails)
  motifsPA: [
    "Maladie",
    "Famille", 
    "Administration",
    "Motif syndical"
  ],
  
  // Motifs PE-B (durées prédéfinies)
  motifsPB: [
    "Mariage du travailleur (02 jours)",
    "Naissance d'un enfant (03 jours)",
    "Décès du conjoint ou d'un ascendante en ligne directe (2 jours)",
    "Mariage d'un enfant, dun frère, ou d'une soeur en ligne directe (2 jours)",
    "Décès du conjoint ou d'un ascendante en ligne directe (1 jours)",
    "Décès d'un ascendant, d'un frère, d'une soeur en ligne directe (2 jours)",
    "Décès d'un beau père ou d'une belle-mère (2 jours)"
  ]
};

// DÉTECTION TYPE DE DEMANDE - Nouvelle fonction clé
function detectDemandType(data) {
  try {
    const typePermission = getSafeValue(data[CONFIG.columns.common.TYPE_PERMISSION]);
    const motifSelection = getSafeValue(data[7]); // Colonne H - Question 7
    
    if (typePermission.toLowerCase().includes("exceptionnelle")) {
      // Détermine PE-A ou PE-B selon le motif choisi
      if (CONFIG.motifsPA.includes(motifSelection) || motifSelection.startsWith("Autre")) {
        return {
          type: "PE-A",
          columns: CONFIG.columns.peA,
          description: "Permission Exceptionnelle - Type A (Motifs généraux)"
        };
      } else if (CONFIG.motifsPB.some(motif => motifSelection.includes(motif.split('(')[0].trim()))) {
        return {
          type: "PE-B", 
          columns: CONFIG.columns.peB,
          description: "Permission Exceptionnelle - Type B (Motifs familiaux)"
        };
      }
    } else if (typePermission.toLowerCase().includes("ordinaire")) {
      return {
        type: "PO",
        columns: CONFIG.columns.po,
        description: "Permission Ordinaire"
      };
    }
    
    // Valeur par défaut si détection échoue
    Logger.log(`[AVERTISSEMENT] Type non détecté, utilisation PE-A par défaut`);
    return {
      type: "PE-A",
      columns: CONFIG.columns.peA,
      description: "Permission Exceptionnelle - Type A (par défaut)"
    };
    
  } catch (error) {
    Logger.log(`[ERREUR detectDemandType] ${error.message}`);
    return {
      type: "PE-A",
      columns: CONFIG.columns.peA,
      description: "Permission Exceptionnelle - Type A (erreur)"
    };
  }
}

// EXTRACTION DONNÉES - Récupère les données selon le type détecté
function extractDemandData(data, demandType) {
  try {
    const cols = demandType.columns;
    const common = CONFIG.columns.common;
    
    // Données communes à tous les types
    const baseData = {
      horodateur: getSafeValue(data[common.HORODATEUR]),
      email: getSafeValue(data[common.EMAIL]),
      matricule: getSafeValue(data[common.MATRICULE]),
      nom: getSafeValue(data[common.NOM]),
      prenom: getSafeValue(data[common.PRENOM]),
      serviceFonction: getSafeValue(data[common.SERVICE]),
      typePermission: getSafeValue(data[common.TYPE_PERMISSION]),
      motifSelection: getSafeValue(data[7]), // Question 7
      demandType: demandType.type
    };
    
    // Données spécifiques selon le type
    if (demandType.type === "PE-A") {
      return {
        ...baseData,
        motifDetail: getSafeValue(data[cols.MOTIF_DETAIL]),
        nbreJours: getSafeValue(data[cols.NB_JOURS]),
        dateDebut: data[cols.DATE_DEBUT],
        heureDebut: getSafeValue(data[cols.HEURE_DEBUT]),
        dateFin: data[cols.DATE_FIN],
        heureFin: getSafeValue(data[cols.HEURE_FIN]),
        emailSuperieur: getSafeValue(data[cols.EMAIL_SUPERIEUR])
      };
    } else if (demandType.type === "PE-B") {
      return {
        ...baseData,
        dateDebut: data[cols.DATE_DEBUT],
        dateFin: data[cols.DATE_FIN],
        emailSuperieur: getSafeValue(data[cols.EMAIL_SUPERIEUR]),
        // PE-B n'a pas d'heures ni de motif détaillé
        heureDebut: "",
        heureFin: "",
        motifDetail: baseData.motifSelection,
        // Calcul automatique du nombre de jours pour PE-B
        nbreJours: calculateDaysBetween(data[cols.DATE_DEBUT], data[cols.DATE_FIN])
      };
    } else { // PO
      return {
        ...baseData,
        motifDetail: getSafeValue(data[cols.MOTIF]),
        nbreJours: getSafeValue(data[cols.NB_JOURS]),
        dateDebut: data[cols.DATE_DEBUT],
        heureDebut: getSafeValue(data[cols.HEURE_DEBUT]),
        dateFin: data[cols.DATE_FIN],
        heureFin: getSafeValue(data[cols.HEURE_FIN]),
        emailSuperieur: getSafeValue(data[cols.EMAIL_SUPERIEUR])
      };
    }
    
  } catch (error) {
    Logger.log(`[ERREUR extractDemandData] ${error.message}`);
    return null;
  }
}

// CALCUL JOURS - Pour les demandes PE-B
function calculateDaysBetween(dateDebut, dateFin) {
  try {
    if (!dateDebut || !dateFin) return "1";
    
    const debut = new Date(dateDebut);
    const fin = new Date(dateFin);
    const diffTime = Math.abs(fin - debut);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
    
    return diffDays.toString();
  } catch (error) {
    Logger.log(`[ERREUR calculateDaysBetween] ${error.message}`);
    return "1";
  }
}

// UTILITAIRES - Fonctions réutilisables (gardées de l'ancien système)

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

// Protège la feuille sauf les colonnes de validation AC, AD, AE, AF
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
      sheet.getRange('AC:AC'),  // Avis du supérieur hiérarchique
      sheet.getRange('AD:AD'),  // Avis des ressources humaines RH
      sheet.getRange('AE:AE'),  // Avis de la présidence
      sheet.getRange('AF:AF')   // Commentaire de la présidence
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