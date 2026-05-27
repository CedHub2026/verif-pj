/*
 * Vérification "pièce jointe oubliée" — Smart Alerts (OnMessageSend)
 * Fonctionne sur Outlook Web, nouveau Outlook Windows, Outlook classique Windows (>= 2412), Mac (préversion).
 *
 * Logique :
 *   1. Lit l'objet + le corps (texte brut).
 *   2. Cherche une mention de pièce jointe (ci-joint, PJ, pièce jointe, attachment, attachement, devis, etc.).
 *   3. Si mention trouvée -> vérifie qu'au moins une pièce jointe NON inline est présente.
 *   4. Sinon -> affiche un avertissement avec option "Envoyer quand même" (SendMode=PromptUser dans le manifeste).
 */

// Expression de détection. Insensible à la casse. \b = limite de mot.
// "devis" est volontairement inclus (cf. demande) : il déclenchera aussi sur des phrases sans PJ
// (ex. "je prépare votre devis"). Comme le mode est "PromptUser", ce n'est jamais bloquant : un clic sur
// "Envoyer quand même" suffit. Retire "devis" du motif si les faux positifs te gênent.
var MOTIF_PJ = /(\bci-?joints?\b|\bci-?jointe?s?\b|\bpi[eè]ces?\s+jointes?\b|\bp\.?\s?j\.?\b|\ben\s+pi[eè]ce\s+jointe\b|\bjoint\s+[aà]\s+ce\b|\battach(?:ment|ed|e)?\b|\battachement\b|\bveuillez\s+trouver\b|\bvous\s+trouverez\b|\bdevis\b)/i;

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;

  item.subject.getAsync(function (resObjet) {
    var objet = (resObjet.status === Office.AsyncResultStatus.Succeeded && resObjet.value) ? resObjet.value : "";

    item.body.getAsync(Office.CoercionType.Text, function (resCorps) {
      var corps = (resCorps.status === Office.AsyncResultStatus.Succeeded && resCorps.value) ? resCorps.value : "";
      var texte = objet + "\n" + corps;

      // Pas de mention de PJ -> on laisse partir.
      if (!MOTIF_PJ.test(texte)) {
        event.completed({ allowEvent: true });
        return;
      }

      // Mention détectée -> on vérifie les pièces jointes réelles.
      item.getAttachmentsAsync(function (resPJ) {
        var aPieceJointe = false;
        if (resPJ.status === Office.AsyncResultStatus.Succeeded && resPJ.value) {
          // On ne compte que les PJ NON inline (les images de signature/corps sont inline -> ignorées).
          aPieceJointe = resPJ.value.some(function (pj) { return pj.isInline === false; });
        }

        if (aPieceJointe) {
          event.completed({ allowEvent: true });
        } else {
          event.completed({
            allowEvent: false,
            // Limite ~150 caractères pour le message court.
            errorMessage: "Le message mentionne une pièce jointe (ci-joint, PJ, devis…) mais aucune n'est attachée. Vérifie avant d'envoyer."
          });
        }
      });
    });
  });
}

// Enregistrement du handler auprès du runtime événementiel.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
