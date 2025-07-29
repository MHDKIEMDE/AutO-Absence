// SERVICE EMAIL - Gestion complète des notifications

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

  // CRÉATION EMAILS - Génère les templates selon le contexte
  static createProfessionalEmail(role, nomDemandeur, docUrl, sheetUrl) {
    const signatures = this.getSignatures();
    
    return {
      // Email de demande de validation (aux validateurs)
      validation: {
        subject: `Validation requise - Demande d'absence de ${nomDemandeur}`,
        message: `
          <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <p>Bonjour,</p>
            <p>Une demande d'absence nécessite votre validation en tant que <strong>${role}</strong>.<br>
            Demandeur : <strong>${nomDemandeur}</strong>.</p>
            <p>Merci de consulter le document et de prendre votre décision dans la feuille de calcul.<br>
            Votre validation est importante pour le traitement de cette demande.</p>
            
            <div style="margin: 25px 0;">
              <a href="${docUrl}" style="display: inline-block; background-color: #1a73e8; color: white; padding: 8px 16px; text-decoration: none; border-radius: 4px; font-weight: 500; margin-right: 10px;">Voir document</a>
              <a href="${sheetUrl}" style="display: inline-block; background-color: #34a853; color: white; padding: 8px 16px; text-decoration: none; border-radius: 4px; font-weight: 500;">Décider</a>
            </div>
            
            ${signatures.validation}
          </div>
        `
      },
      
      // Email de confirmation (au demandeur)
      confirmation: {
        subject: `Confirmation - Votre demande d'absence`,
        message: `
          <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <p>Bonjour ${nomDemandeur.split(' ')[0]},</p>
            <p>Votre demande d'absence a été soumise avec succès.<br>
            Elle est actuellement en cours de traitement par votre hiérarchie.</p>
            <p>Vous recevrez une notification dès qu'une décision sera prise.<br>
            Merci pour votre patience.</p>
            
            <div style="margin: 25px 0;">
              <a href="${docUrl}" style="display: inline-block; background-color: #34a853; color: white; padding: 8px 16px; text-decoration: none; border-radius: 4px; font-weight: 500;">Suivre</a>
            </div>
            
            ${signatures.employee}
          </div>
        `
      },
      
      // Email d'approbation (au demandeur)
      approval: {
        subject: `Demande d'absence approuvée`,
        message: `
          <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <p>Bonjour ${nomDemandeur.split(' ')[0]},</p>
            <p>Nous avons le plaisir de vous informer que votre demande d'absence a été entièrement approuvée.<br>
            Toutes les validations nécessaires ont été obtenues.</p>
            <p>Vous pouvez maintenant organiser votre absence selon les dates demandées.<br>
            Le document final est disponible ci-dessous.</p>
            
            <div style="margin: 25px 0;">
              <a href="${docUrl}" style="display: inline-block; background-color: #34a853; color: white; padding: 8px 16px; text-decoration: none; border-radius: 4px; font-weight: 500;">Voir document</a>
            </div>
            
            ${signatures.approval}
          </div>
        `
      },
      
      // Email de rejet (au demandeur)
      rejection: {
        subject: `Demande d'absence refusée`,
        message: `
          <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <p>Bonjour ${nomDemandeur.split(' ')[0]},</p>
            <p>Nous vous informons que votre demande d'absence a été refusée par <strong>${role}</strong>.<br>
            Cette décision a été prise après examen de votre dossier.</p>
            <p>Les motifs du refus sont détaillés dans le document joint.<br>
            Vous pouvez soumettre une nouvelle demande si nécessaire.</p>
            
            <div style="margin: 25px 0;">
              <a href="${docUrl}" style="display: inline-block; background-color: #ea4335; color: white; padding: 8px 16px; text-decoration: none; border-radius: 4px; font-weight: 500;">Voir document</a>
            </div>
            
            ${signatures.rejection}
          </div>
        `
      }
    };
  }

  // NOTIFICATIONS SPÉCIALISÉES - Fonctions prêtes à l'emploi

  // Notifie un validateur
  static notifyValidator(validatorEmail, role, nomDemandeur, docUrl, sheetUrl) {
    const emails = this.createProfessionalEmail(role, nomDemandeur, docUrl, sheetUrl);
    return this.sendEmail(validatorEmail, emails.validation.subject, emails.validation.message);
  }

  // Confirme la soumission au demandeur
  static confirmSubmission(demandeurEmail, nomDemandeur, docUrl) {
    const emails = this.createProfessionalEmail("", nomDemandeur, docUrl, "");
    return this.sendEmail(demandeurEmail, emails.confirmation.subject, emails.confirmation.message);
  }

  // Notifie l'approbation finale
  static notifyApproval(demandeurEmail, nomDemandeur, docUrl, role) {
    const emails = this.createProfessionalEmail(role, nomDemandeur, docUrl, "");
    return this.sendEmail(demandeurEmail, emails.approval.subject, emails.approval.message);
  }

  // Notifie le rejet
  static notifyRejection(demandeurEmail, nomDemandeur, docUrl, role) {
    const emails = this.createProfessionalEmail(role, nomDemandeur, docUrl, "");
    return this.sendEmail(demandeurEmail, emails.rejection.subject, emails.rejection.message);
  }

  // CHAÎNE DE VALIDATION - Notifie le prochain validateur
  static notifyNextValidator(sheet, row, colonne, docId) {
    try {
      var nextRole = "";
      var nextEmail = "";

      // Détermine le prochain validateur selon la colonne modifiée
      if (colonne == 17) {          // Supérieur → RH
        nextRole = "RH";
        nextEmail = sheet.getRange(row, 22).getValue();
      } else if (colonne == 18) {   // RH → Présidence
        nextRole = "Présidence";  
        nextEmail = sheet.getRange(row, 23).getValue();
      }

      // Envoie la notification si email valide
      if (nextEmail && nextEmail.includes("@")) {
        var doc = DriveApp.getFileById(docId);
        var demandeurNom = `${sheet.getRange(row, 4).getValue()} ${sheet.getRange(row, 5).getValue()}`;
        var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
        
        const success = this.notifyValidator(nextEmail, nextRole, demandeurNom, doc.getUrl(), sheetUrl);
        
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