/*
 * Vérification "pièce jointe oubliée" — Smart Alerts (OnMessageSend)
 * Fonctionne sur Outlook Web, nouveau Outlook Windows, Outlook classique Windows (>= 2412), Mac (préversion).
 *
 * Lecture de l'objet, du corps et des pièces jointes EN PARALLÈLE (3 appels simultanés),
 * puis décision une fois les trois revenus -> temps = lecture la plus lente, pas la somme.
 */

// Motif de détection. Insensible à la casse. \b = limite de mot.
// "devis" déclenche seul (cf. demande) ; non bloquant grâce au mode "PromptUser".
var MOTIF_PJ = /(\bci-?joints?\b|\bci-?jointe?s?\b|\bpi[eè]ces?\s+jointes?\b|\bp\.?\s?j\.?\b|\ben\s+pi[eè]ce\s+jointe\b|\bjoint\s+[aà]\s+ce\b|\battach(?:ment|ed|e)?\b|\battachement\b|\bveuillez\s+trouver\b|\bvous\s+trouverez\b|\bdevis\b)/i;

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;

  var restants = 3;
  var objet = "";
  var corps = "";
  var aPieceJointe = false;

  function decider() {
    restants--;
    if (restants > 0) return; // on attend que les 3 lectures soient revenues

    var texte = objet + "\n" + corps;
    if (aPieceJointe || !MOTIF_PJ.test(texte)) {
      event.completed({ allowEvent: true });
    } else {
      event.completed({
        allowEvent: false,
        errorMessage: "Le message mentionne une pièce jointe (ci-joint, PJ, devis…) mais aucune n'est attachée. Vérifie avant d'envoyer."
      });
    }
  }

  // Les trois lectures partent en même temps.
  item.subject.getAsync(function (r) {
    if (r.status === Office.AsyncResultStatus.Succeeded && r.value) objet = r.value;
    decider();
  });

  item.body.getAsync(Office.CoercionType.Text, function (r) {
    if (r.status === Office.AsyncResultStatus.Succeeded && r.value) corps = r.value;
    decider();
  });

  item.getAttachmentsAsync(function (r) {
    if (r.status === Office.AsyncResultStatus.Succeeded && r.value) {
      aPieceJointe = r.value.some(function (pj) { return pj.isInline === false; });
    }
    decider();
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
