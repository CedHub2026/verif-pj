/*
 * Vérification "pièce jointe oubliée" — Smart Alerts (OnMessageSend)
 * Outlook Web, nouveau Outlook Windows, Outlook classique Windows (>= 2412), Mac (préversion).
 *
 * Optimisation anti-délai :
 *   - dès qu'une pièce jointe non inline est trouvée, on AUTORISE sans attendre la lecture
 *     du corps (cas le plus courant et le plus lent sur les fils longs) ;
 *   - on ne lit/analyse le texte que s'il n'y a AUCUNE pièce jointe.
 *   - la recherche ne porte que sur le texte saisi (corps tronqué au 1er séparateur de citation).
 */

var MOTIF_PJ = /(\bci-?joints?\b|\bci-?jointe?s?\b|\bpi[eè]ces?\s+jointes?\b|\bp\.?\s?j\.?\b|\ben\s+pi[eè]ce\s+jointe\b|\bjoint\s+[aà]\s+ce\b|\battach(?:ment|ed|e)?\b|\battachement\b|\bveuillez\s+trouver\b|\bvous\s+trouverez\b|\bdevis\b)/i;

var SEPARATEURS = [
  /\n\s*_{10,}\s*\n/,
  /-{3,}\s*(?:Original Message|Message d.origine)\s*-{3,}/i,
  /^\s*(?:De|From|Exp[eé]diteur)\s*:.*$/m,
  /^\s*Le\s.+\s[aà]\s[eé]crit\s*:/m,
  /^\s*On\s.+\swrote\s*:/m
];

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

  var fini = false;
  var objet = null;   // null = pas encore lu
  var corps = null;
  var attLu = false;
  var aPieceJointe = false;

  function finir(autoriser) {
    if (fini) return;
    fini = true;
    if (autoriser) {
      event.completed({ allowEvent: true });
    } else {
      event.completed({
        allowEvent: false,
        errorMessage: "Le message mentionne une pièce jointe (ci-joint, PJ, devis…) mais aucune n'est attachée. Vérifie avant d'envoyer."
      });
    }
  }

  function decider() {
    if (fini) return;
    // PJ présente -> on autorise tout de suite, sans attendre le corps.
    if (attLu && aPieceJointe) { finir(true); return; }
    // Pas de PJ -> il faut l'objet ET le corps pour analyser le texte saisi.
    if (attLu && !aPieceJointe && objet !== null && corps !== null) {
      var texte = objet + "\n" + texteSaisi(corps);
      finir(!MOTIF_PJ.test(texte)); // aucune formule -> autoriser ; formule trouvée -> avertir
    }
  }

  // Les pièces jointes en priorité : c'est ce qui permet de court-circuiter.
  item.getAttachmentsAsync(function (r) {
    if (r.status === Office.AsyncResultStatus.Succeeded && r.value) {
      aPieceJointe = r.value.some(function (pj) { return pj.isInline === false; });
    }
    attLu = true;
    decider();
  });

  // Objet + corps lancés en parallèle ; seulement utiles s'il n'y a pas de PJ.
  item.subject.getAsync(function (r) {
    objet = (r.status === Office.AsyncResultStatus.Succeeded && r.value) ? r.value : "";
    decider();
  });

  item.body.getAsync(Office.CoercionType.Text, function (r) {
    corps = (r.status === Office.AsyncResultStatus.Succeeded && r.value) ? r.value : "";
    decider();
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
