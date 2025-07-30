// SERVICE EMAIL - Gestion complète des notifications PE-A/PE-B/PO

class EmailService {
  
  // SIGNATURES - Modifiez ici pour personnaliser vos signatures
  static getSignatures() {
    return {
      // Signature pour notifications aux validateurs
      validation: `
        <div style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #e0e0e0;">
          <p style="margin: 5px 0; color: #666; font-size: 13px;">
            <strong>Service Ressources Humaines</strong><br>
            Département Gestion du Personnel<br>
            📧 rh@entreprise.com | 📞 +221 XX XXX XX XX<br>
            🏢 Adresse de votre entreprise
          </p>
          <p style="margin: 8px 0 0 0; color: #999; font-size: 11px;">
            <em>Message automatique - Système de gestion des absences.</em>
          </p>
        </div>
      `,
      
      // Signature pour emails aux employés
      employee: `
        <div style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #e0e0e0;">
          <p style="margin: 5px 0; color: #666; font-size: 13px;">
            <strong>Équipe RH</strong><br>
            Service Ressources Humaines<br>
            📧 rh@entreprise.com<br>
            💡 <em>Pour toute question, n'hésitez pas à nous contacter</em>
          </p>
        </div>
      `,
      
      // Signature pour rejets (plus formelle)
      rejection: `
        <div style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #e0e0e0;">
          <p style="margin: 5px 0; color: #666; font-size: 13px;">
            <strong>Service Ressources Humaines</strong><br>
            📧 rh@entreprise.com | 📞 +221 XX XXX XX XX<br>
            🕒 Horaires : Lundi - Vendredi, 8h - 17h
          </p>
          <p style="margin: 8px 0 0 0; color: #999; font-size: 11px;">
            <em>Délai de recours selon règlement intérieur.</em>
          </p>
        </div>
      `,
      
      // Signature pour approbations (chaleureuse)
      approval: `
        <div style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #e0e0e0;">
          <p style="margin: 5px 0; color: #666; font-size: 13px;">
            <strong>Équipe RH</strong><br>
            📧 rh@entreprise.com<br>
            ✨ <em>Nous vous souhaitons un excellent repos !</em>
          </p>
        </div>
      `
    };
  }

  // ENVOI EMAIL - Fonction principale d'envoi
  static sendEmail(recipient, subject, body) {
    try {
      // Validation de l'email destinataire
      if (!recipient || !recipient.includes("@")) {
        Logger.log(`[ERREUR EMAIL] Email invalide : ${recipient}`);
        return false;
      }

      // Envoi via Gmail API
      GmailApp.sendEmail(recipient, subject, "", { htmlBody: body });
      Logger.log(`[EMAIL ENVOYÉ] À : ${recipient} | Sujet : ${subject}`);
      return true;
    } catch (error) {
      Logger.log(`[ERREUR EMAIL] ${error.message}`);
      return false;
    }
  }

  // FORMATAGE DONNÉES - Prépare les informations pour affichage dans l'email
  static formatDemandData(demandData) {
    if (!demandData) return "";

    // Gère l'affichage des heures selon le type
    let heureDebutText = "";
    let heureFinText = "";
    
    if (demandData.heureDebut && demandData.heureDebut.trim() !== "") {
      heureDebutText = ` à ${demandData.heureDebut}`;
    }
    if (demandData.heureFin && demandData.heureFin.trim() !== "") {
      heureFinText = ` à ${demandData.heureFin}`;
    }

    // Formate les dates
    let dateDebutFormatted = demandData.dateDebut;
    let dateFinFormatted = demandData.dateFin;
    
    if (demandData.dateDebut instanceof Date) {
      dateDebutFormatted = demandData.dateDebut.toLocaleDateString("fr-FR");
    }
    if (demandData.dateFin instanceof Date) {
      dateFinFormatted = demandData.dateFin.toLocaleDateString("fr-FR");
    }

    // Badge de couleur selon le type de demande
    let typeBadge = "";
    let typeColor = "#007bff"; // Bleu par défaut
    
    if (demandData.demandType === "PE-A") {
      typeColor = "#28a745"; // Vert pour PE-A
      typeBadge = `<span style="background-color: ${typeColor}; color: white; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 500;">PE-A</span>`;
    } else if (demandData.demandType === "PE-B") {
      typeColor = "#17a2b8"; // Turquoise pour PE-B
      typeBadge = `<span style="background-color: ${typeColor}; color: white; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 500;">PE-B</span>`;
    } else if (demandData.demandType === "PO") {
      typeColor = "#ffc107"; // Jaune pour PO
      typeBadge = `<span style="background-color: ${typeColor}; color: black; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 500;">PO</span>`;
    }

    return `
      <div style="background-color: #f8f9fa; padding: 15px; border-radius: 6px; margin: 15px 0; border-left: 4px solid ${typeColor};">
        <h4 style="margin: 0 0 10px 0; color: #333;">📋 Résumé de la demande ${typeBadge}</h4>
        <p style="margin: 5px 0;"><strong>Type :</strong> ${demandData.typePermission || 'Non spécifié'}</p>
        <p style="margin: 5px 0;"><strong>Durée :</strong> ${demandData.nbreJours || 'Non spécifié'} jour(s)</p>
        <p style="margin: 5px 0;"><strong>Du :</strong> ${dateDebutFormatted || 'Non spécifiée'}${heureDebutText}</p>
        <p style="margin: 5px 0;"><strong>Au :</strong> ${dateFinFormatted || 'Non spécifiée'}${heureFinText}</p>
        <p style="margin: 5px 0;"><strong>Motif :</strong> ${demandData.motifDetail || demandData.motifSelection || 'Non spécifié'}</p>
        ${demandData.serviceFonction ? `<p style="margin: 5px 0;"><strong>Service :</strong> ${demandData.serviceFonction}</p>` : ''}
        ${demandData.matricule ? `<p style="margin: 5px 0; color: #666; font-size: 12px;"><strong>Matricule :</strong> ${demandData.matricule}</p>` : ''}
      </div>
    `;
  }

  // CRÉATION EMAILS - Génère les templates avec informations complètes
  static createProfessionalEmail(role, nomDemandeur, docUrl, sheetUrl, demandData = null) {
    const signatures = this.getSignatures();
    
    // Crée le résumé de la demande si les données sont fournies
    let demandeResume = "";
    if (demandData) {
      demandeResume = this.formatDemandData(demandData);
    }
    
    return {
      // Email de demande de validation (aux validateurs)
      validation: {
        subject: `[${demandData?.demandType || 'DEMANDE'}] Validation requise - ${nomDemandeur}`,
        message: `
          <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <p>Bonjour,</p>
            <p>Une demande d'absence nécessite votre validation en tant que <strong>${role}</strong>.<br>
            Demandeur : <strong>${nomDemandeur}</strong></p>
            
            ${demandeResume}
            
            <p style="margin-top: 20px;">Vous pouvez prendre votre décision directement dans la feuille de calcul ou consulter le document complet pour plus de détails.</p>
            
            <div style="margin: 25px 0;">
              <a href="${sheetUrl}" style="display: inline-block; background-color: #34a853; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: 500; margin-right: 10px;">✅ Décider maintenant</a>
              <a href="${docUrl}" style="display: inline-block; background-color: #1a73e8; color: white; padding: 8px 16px; text-decoration: none; border-radius: 4px; font-weight: 500;">📄 Voir document complet</a>
            </div>
            
            <p style="color: #666; font-size: 13px;">
              <strong>Instructions :</strong> Dans la feuille, cherchez la ligne correspondante et indiquez "Favorable" ou "Non Favorable" dans votre colonne de validation.
            </p>
            
            ${signatures.validation}
          </div>
        `
      },
      
      // Email de confirmation (au demandeur)
      confirmation: {
        subject: `Confirmation - Votre demande d'absence [${demandData?.demandType || 'REF'}]`,
        message: `
          <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <p>Bonjour ${nomDemandeur.split(' ')[0]},</p>
            <p>Votre demande d'absence a été soumise avec succès.<br>
            Elle est actuellement en cours de traitement par votre hiérarchie.</p>
            
            ${demandeResume}
            
            <p><strong>Étapes de validation :</strong></p>
            <div style="background-color: #e9ecef; padding: 12px; border-radius: 4px; margin: 10px 0;">
              <p style="margin: 3px 0;">1️⃣ Supérieur Hiérarchique ⏳ <em>En cours...</em></p>
              <p style="margin: 3px 0;">2️⃣ Service RH ⏸️ <em>En attente</em></p>
              <p style="margin: 3px 0;">3️⃣ Présidence ⏸️ <em>En attente</em></p>
            </div>
            
            <p>Vous recevrez une notification dès qu'une décision sera prise.<br>
            Merci pour votre patience.</p>
            
            <div style="margin: 25px 0;">
              <a href="${docUrl}" style="display: inline-block; background-color: #34a853; color: white; padding: 8px 16px; text-decoration: none; border-radius: 4px; font-weight: 500;">📋 Suivre ma demande</a>
            </div>
            
            ${signatures.employee}
          </div>
        `
      },
      
      // Email d'approbation (au demandeur)
      approval: {
        subject: `✅ Demande d'absence approuvée [${demandData?.demandType || 'REF'}]`,
        message: `
          <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <p>Bonjour ${nomDemandeur.split(' ')[0]},</p>
            <p style="color: #28a745; font-weight: 500;">🎉 Excellente nouvelle ! Votre demande d'absence a été entièrement approuvée.<br>
            Toutes les validations nécessaires ont été obtenues.</p>
            
            ${demandeResume}
            
            <div style="background-color: #d4edda; border: 1px solid #c3e6cb; padding: 12px; border-radius: 4px; margin: 15px 0;">
              <p style="margin: 3px 0; color: #155724;"><strong>✅ Statut final :</strong> APPROUVÉ</p>
              <p style="margin: 3px 0; color: #155724;">✅ Supérieur Hiérarchique : Favorable</p>
              <p style="margin: 3px 0; color: #155724;">✅ Service RH : Favorable</p>
              <p style="margin: 3px 0; color: #155724;">✅ Présidence : Favorable</p>
            </div>
            
            <p>Vous pouvez maintenant organiser votre absence selon les dates demandées.<br>
            Le document final est disponible ci-dessous pour vos archives.</p>
            
            <div style="margin: 25px 0;">
              <a href="${docUrl}" style="display: inline-block; background-color: #28a745; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: 500;">📄 Télécharger l'autorisation</a>
            </div>
            
            ${signatures.approval}
          </div>
        `
      },
      
      // Email de rejet (au demandeur)
      rejection: {
        subject: `❌ Demande d'absence refusée [${demandData?.demandType || 'REF'}]`,
        message: `
          <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <p>Bonjour ${nomDemandeur.split(' ')[0]},</p>
            <p>Nous vous informons que votre demande d'absence a été refusée par <strong>${role}</strong>.<br>
            Cette décision a été prise après examen de votre dossier.</p>
            
            ${demandeResume}
            
            <div style="background-color: #f8d7da; border: 1px solid #f5c6cb; padding: 12px; border-radius: 4px; margin: 15px 0;">
              <p style="margin: 3px 0; color: #721c24;"><strong>❌ Statut :</strong> REFUSÉ par ${role}</p>
              <p style="margin: 3px 0; color: #721c24;">📝 Les motifs du refus sont détaillés dans le document joint</p>
            </div>
            
            <p>Si vous souhaitez contester cette décision ou soumettre une nouvelle demande avec des modifications, vous pouvez nous contacter directement.</p>
            
            <div style="margin: 25px 0;">
              <a href="${docUrl}" style="display: inline-block; background-color: #dc3545; color: white; padding: 8px 16px; text-decoration: none; border-radius: 4px; font-weight: 500;">📄 Voir les détails</a>
              <a href="mailto:${CONFIG.emails.rh}" style="display: inline-block; background-color: #6c757d; color: white; padding: 8px 16px; text-decoration: none; border-radius: 4px; font-weight: 500; margin-left: 10px;">✉️ Nous contacter</a>
            </div>
            
            ${signatures.rejection}
          </div>
        `
      }
    };
  }

  // NOTIFICATIONS SPÉCIALISÉES - Fonctions avec données enrichies

  // Notifie un validateur avec informations de la demande
  static notifyValidator(validatorEmail, role, nomDemandeur, docUrl, sheetUrl, demandData = null) {
    const emails = this.createProfessionalEmail(role, nomDemandeur, docUrl, sheetUrl, demandData);
    return this.sendEmail(validatorEmail, emails.validation.subject, emails.validation.message);
  }

  // Confirme la soumission au demandeur avec récapitulatif
  static confirmSubmission(demandeurEmail, nomDemandeur, docUrl, demandData = null) {
    const emails = this.createProfessionalEmail("", nomDemandeur, docUrl, "", demandData);
    return this.sendEmail(demandeurEmail, emails.confirmation.subject, emails.confirmation.message);
  }

  // Notifie l'approbation finale avec récapitulatif
  static notifyApproval(demandeurEmail, nomDemandeur, docUrl, role, demandData = null) {
    const emails = this.createProfessionalEmail(role, nomDemandeur, docUrl, "", demandData);
    return this.sendEmail(demandeurEmail, emails.approval.subject, emails.approval.message);
  }

  // Notifie le rejet avec récapitulatif
  static notifyRejection(demandeurEmail, nomDemandeur, docUrl, role, demandData = null) {
    const emails = this.createProfessionalEmail(role, nomDemandeur, docUrl, "", demandData);
    return this.sendEmail(demandeurEmail, emails.rejection.subject, emails.rejection.message);
  }

  // CHAÎNE DE VALIDATION PE-A/PE-B/PO - Version corrigée avec vraies colonnes
  static notifyNextValidator(sheet, row, colonne, docId) {
    try {
      var nextRole = "";
      var nextEmail = "";

      // Détermine le prochain validateur selon la colonne modifiée
      if (colonne == CONFIG.columns.validation.SUPERIEUR) {          // Supérieur → RH
        nextRole = "RH";
        nextEmail = CONFIG.emails.rh;
      } else if (colonne == CONFIG.columns.validation.RH) {   // RH → Présidence
        nextRole = "Présidence";  
        nextEmail = sheet.getRange(row, CONFIG.columns.metadata.EMAIL_PRESIDENCE + 1).getValue();
      }

      // Envoie la notification si email valide
      if (nextEmail && nextEmail.includes("@")) {
        var doc = DriveApp.getFileById(docId);
        var demandeurNom = `${sheet.getRange(row, CONFIG.columns.common.NOM + 1).getValue()} ${sheet.getRange(row, CONFIG.columns.common.PRENOM + 1).getValue()}`;
        var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
        
        // Récupère les données de la demande depuis la feuille (version PE-A/PE-B/PO)
        var demandData = this.getDemandDataFromSheet(sheet, row);
        
        const success = this.notifyValidator(nextEmail, nextRole, demandeurNom, doc.getUrl(), sheetUrl, demandData);
        
        if (success) {
          Logger.log(`[NOTIFICATION] ${nextRole} notifié : ${nextEmail}`);
        }
        return success;
      }
      return false;
    } catch (error) {
      Logger.log(`[ERREUR notifyNextValidator] ${error.message}`);
      return false;
    }
  }

  // RÉCUPÉRATION DONNÉES - Version corrigée avec vraies colonnes du sheet
  static getDemandDataFromSheet(sheet, row) {
    try {
      // Récupère le type de permission pour déterminer quelles colonnes lire
      var typePermission = sheet.getRange(row, CONFIG.columns.common.TYPE_PERMISSION + 1).getValue() || "";
      var motifSelection = sheet.getRange(row, CONFIG.columns.peA.MOTIF_SELECTION + 1).getValue() || ""; // Colonne H
      
      // Données communes
      var baseData = {
        matricule: sheet.getRange(row, CONFIG.columns.common.MATRICULE + 1).getValue() || "",
        nom: sheet.getRange(row, CONFIG.columns.common.NOM + 1).getValue() || "",
        prenom: sheet.getRange(row, CONFIG.columns.common.PRENOM + 1).getValue() || "",
        serviceFonction: sheet.getRange(row, CONFIG.columns.common.SERVICE + 1).getValue() || "",
        typePermission: typePermission,
        motifSelection: motifSelection
      };
      
      // Détecte le type de demande (PE-A, PE-B, PO)
      var demandType = this.detectTypeFromSheet(typePermission, motifSelection);
      
      if (demandType === "PE-A") {
        return {
          ...baseData,
          demandType: "PE-A",
          motifDetail: sheet.getRange(row, CONFIG.columns.peA.MOTIF_DETAIL + 1).getValue() || motifSelection,
          nbreJours: sheet.getRange(row, CONFIG.columns.peA.NB_JOURS + 1).getValue() || "Non spécifié",
          dateDebut: sheet.getRange(row, CONFIG.columns.peA.DATE_DEBUT + 1).getValue() || "Non spécifiée",
          heureDebut: sheet.getRange(row, CONFIG.columns.peA.HEURE_DEBUT + 1).getValue() || "",
          dateFin: sheet.getRange(row, CONFIG.columns.peA.DATE_FIN + 1).getValue() || "Non spécifiée",
          heureFin: sheet.getRange(row, CONFIG.columns.peA.HEURE_FIN + 1).getValue() || ""
        };
      } else if (demandType === "PE-B") {
        var dateDebut = sheet.getRange(row, CONFIG.columns.peB.DATE_DEBUT + 1).getValue();
        var dateFin = sheet.getRange(row, CONFIG.columns.peB.DATE_FIN + 1).getValue();
        
        return {
          ...baseData,
          demandType: "PE-B",
          motifDetail: motifSelection,
          dateDebut: dateDebut || "Non spécifiée",
          dateFin: dateFin || "Non spécifiée",
          heureDebut: sheet.getRange(row, CONFIG.columns.peB.HEURE_DEBUT + 1).getValue() || "",
          heureFin: sheet.getRange(row, CONFIG.columns.peB.HEURE_FIN + 1).getValue() || "",
          nbreJours: sheet.getRange(row, CONFIG.columns.peB.NB_JOURS + 1).getValue() || this.calculateDaysFromSheet(dateDebut, dateFin)
        };
      } else { // PO
        return {
          ...baseData,
          demandType: "PO",
          motifDetail: sheet.getRange(row, CONFIG.columns.po.MOTIF + 1).getValue() || "Non spécifié",
          nbreJours: sheet.getRange(row, CONFIG.columns.po.NB_JOURS + 1).getValue() || "Non spécifié",
          dateDebut: sheet.getRange(row, CONFIG.columns.po.DATE_DEBUT + 1).getValue() || "Non spécifiée",
          heureDebut: sheet.getRange(row, CONFIG.columns.po.HEURE_DEBUT + 1).getValue() || "",
          dateFin: sheet.getRange(row, CONFIG.columns.po.DATE_FIN + 1).getValue() || "Non spécifiée",
          heureFin: sheet.getRange(row, CONFIG.columns.po.HEURE_FIN + 1).getValue() || ""
        };
      }
      
    } catch (error) {
      Logger.log(`[ERREUR getDemandDataFromSheet] ${error.message}`);
      return {
        demandType: "Inconnu",
        matricule: "",
        nom: "Inconnu",
        prenom: "",
        serviceFonction: "",
        typePermission: "Non spécifié",
        motifDetail: "Non spécifié",
        nbreJours: "Non spécifié",
        dateDebut: "Non spécifiée",
        heureDebut: "",
        dateFin: "Non spécifiée",
        heureFin: ""
      };
    }
  }

  // DÉTECTION TYPE - Depuis les données de la feuille
  static detectTypeFromSheet(typePermission, motifSelection) {
    try {
      if (typePermission.toLowerCase().includes("exceptionnelle")) {
        if (CONFIG.motifsPA.includes(motifSelection) || motifSelection.startsWith("Autre")) {
          return "PE-A";
        } else {
          return "PE-B";
        }
      } else {
        return "PO";
      }
    } catch (error) {
      Logger.log(`[ERREUR detectTypeFromSheet] ${error.message}`);
      return "PE-A"; // Par défaut
    }
  }

  // CALCUL JOURS - Pour PE-B depuis la feuille
  static calculateDaysFromSheet(dateDebut, dateFin) {
    try {
      if (!dateDebut || !dateFin) return "1";
      
      const debut = new Date(dateDebut);
      const fin = new Date(dateFin);
      const diffTime = Math.abs(fin - debut);
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
      
      return diffDays.toString();
    } catch (error) {
      Logger.log(`[ERREUR calculateDaysFromSheet] ${error.message}`);
      return "1";
    }
  }
}

// COMPATIBILITÉ - Fonctions pour l'ancien code
function sendEmail(recipient, subject, body) {
  return EmailService.sendEmail(recipient, subject, body);
}

function createProfessionalEmail(role, nomDemandeur, docUrl, sheetUrl) {
  return EmailService.createProfessionalEmail(role, nomDemandeur, docUrl, sheetUrl);
}

function notifyNextValidator(sheet, row, colonne, docId) {
  return EmailService.notifyNextValidator(sheet, row, colonne, docId);
}