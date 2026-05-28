/*
 * Vérification "pièce jointe oubliée" — Smart Alerts (OnMessageSend)
 * Outlook Web, nouveau Outlook Windows, Outlook classique Windows (>= 2412), Mac (préversion).
 *
 * Lectures objet + corps + pièces jointes EN PARALLÈLE.
 * La recherche des formules ne porte QUE sur le texte saisi : le corps est coupé
 * au premier séparateur de message cité / transféré (l'historique du fil est ignoré).
 */

// Formules déclenchantes. Insensible à la casse. \b = limite de mot.
var MOTIF_PJ = /(\bci-?joints?\b|\bci-?jointe?s?\b|\bpi[eè]ces?\s+jointes?\b|\bp\.?\s?j\.?\b|\ben\s+pi[eè]ce\s+jointe\b|\bjoint\s+[aà]\s+ce\b|\battach(?:ment|ed|e)?\b|\battachement\b|\bveuillez\s+trouver\b|\bvous\s+trouverez\b|\bdevis\b)/i;

// Séparateurs marquant le DÉBUT du message cité / transféré (Outlook FR/EN, Web et Desktop).
// Tout ce qui suit le premier séparateur trouvé est ignoré.
var SEPARATEURS = [
  /\n\s*_{10,}\s*\n/,                                  // ligne de soulignés insérée par Outlook
  /-{3,}\s*(?:Original Message|Message d.origine)\s*-{3,}/i,
  /^\s*(?:De|From|Exp[eé]diteur)\s*:.*$/m,             // en-tête du message cité (De :/From:)
  /^\s*Le\s.+\s[aà]\s[eé]crit\s*:/m,                   // "Le <date>, X a écrit :"
  /^\s*On\s.+\swrote\s*:/m                             // "On <date>, X wrote:"
];

// Renvoie uniquement le texte saisi (avant le premier séparateur).
function texteSaisi(corps) {
  if (!corps) return "";
  var coupe = corps.length;
  for (var i = 0; i < SEPARATEURS.length; i++) {
    var m = corps.match(SEPARATEURS[i]);
    if (m && m.index < coupe) coupe = m.index;
  }
  return corps.slice(0, coupe);
}

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;

  var restants = 3;
  var objet = "";
  var corps = "";
  var aPieceJointe = false;

  function decider() {
    restants--;
    if (restants > 0) return;

    var texte = objet + "\n" + texteSaisi(corps);
    if (aPieceJointe || !MOTIF_PJ.test(texte)) {
      event.completed({ allowEvent: true });
    } else {
      event.completed({
        allowEvent: false,
        errorMessage: "Le message mentionne une pièce jointe (ci-joint, PJ, devis…) mais aucune n'est attachée. Vérifie avant d'envoyer."
      });
    }
  }

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
