# Vérification pièce jointe — add-in Outlook

Garde-fou anti-oubli de pièce jointe. À l'envoi d'un message, si le texte évoque une
pièce jointe sans qu'aucune ne soit attachée, Outlook affiche un avertissement avec
l'option « Envoyer quand même ».

Fonctionne sur Outlook Web, le nouveau Outlook pour Windows, Outlook classique Windows
(version 2412 ou ultérieure) et macOS (préversion). Installé en sideload, pour un seul
utilisateur, sans droits administrateur.

## Besoin couvert

- Exécution automatique à l'envoi, sur Outlook Web et Outlook Desktop.
- Détection dans l'objet et le corps des formules : ci-joint, PJ, pièce jointe,
  attachment, attachement, devis (ainsi que « veuillez trouver », « vous trouverez »).
- Vérification réelle d'au moins une pièce jointe **non inline** (les images de signature
  ou de corps sont ignorées).
- En l'absence de pièce jointe : avertissement non bloquant avec « Envoyer quand même ».
- Recherche limitée au **texte saisi** : l'historique cité d'une réponse ou d'un transfert
  n'est pas analysé.

## Approche technique

Add-in de type **Smart Alerts**, branché sur l'événement `OnMessageSend`, en mode
`PromptUser` (avertissement avec choix « Envoyer quand même », plutôt que blocage strict
qui imposerait un déploiement par l'administrateur). Cible le jeu d'exigences Mailbox 1.12,
requis par `OnMessageSend`.

C'est la seule des trois approches envisagées à couvrir Web et bureau : une macro VBA
`Application_ItemSend` ne vaut que pour Outlook classique Windows ; le rappel de pièce
jointe natif d'Outlook n'est pas personnalisable.

## Fichiers

| Fichier | Rôle |
| --- | --- |
| `manifest.xml` | Déclare l'add-in, l'événement `OnMessageSend`, le mode `PromptUser` et les URL des ressources. C'est le fichier chargé dans Outlook ; il reste sur le poste. |
| `commands.html` | Page de runtime sans interface ; charge `office.js` et le script. |
| `commands.js` | Logique de détection et de décision. |
| `icon-64.png`, `icon-128.png` | Icônes requises par le manifeste. |
| `README.md` | Ce document. |

`commands.html`, `commands.js` et les deux icônes sont servis en HTTPS depuis GitHub Pages :
`https://cedhub2026.github.io/verif-pj/`. Le `manifest.xml` n'a pas besoin d'être hébergé.

## Logique de traitement

À chaque envoi, le handler `onMessageSendHandler` :

1. Lance **en parallèle** trois lectures : objet (`subject.getAsync`), corps en texte brut
   (`body.getAsync`), pièces jointes (`getAttachmentsAsync`). Le temps de traitement égale
   la lecture la plus lente, et non leur somme.
2. **Isole le texte saisi** : le corps est tronqué au premier séparateur de citation détecté
   (ligne de soulignés d'Outlook, en-tête `De :` / `From:`, « Message d'origine »,
   « Le … a écrit : », « On … wrote: »). Seul ce qui précède est analysé.
3. **Détecte les formules** via une expression régulière insensible à la casse, sur l'objet
   et le texte saisi.
4. **Vérifie réellement** : une pièce jointe ne compte que si `isInline === false`.
5. **Décide** : envoi autorisé s'il existe une pièce jointe ou si aucune formule n'est
   détectée ; sinon avertissement avec « Envoyer quand même ».

Les mots-clés sont définis dans la constante `MOTIF_PJ`, les séparateurs de citation dans
`SEPARATEURS`, en tête de `commands.js`.

## Installation (sideload, sans admin)

Prérequis : le tenant doit autoriser les compléments personnalisés (vérifiable dans
`https://aka.ms/olksideload` — la section « Compléments personnalisés / À partir d'un
fichier » doit être présente).

1. Héberger `commands.html`, `commands.js`, `icon-64.png`, `icon-128.png` en HTTPS
   (ici GitHub Pages, dépôt public). Vérifier que
   `https://cedhub2026.github.io/verif-pj/commands.js` affiche bien le code.
2. Conserver `manifest.xml` sur le poste (URL et GUID déjà renseignés).
3. Ouvrir `https://aka.ms/olksideload` → **Mes compléments › Compléments personnalisés ›
   Ajouter un complément personnalisé › À partir d'un fichier** → sélectionner `manifest.xml`.
4. Test : nouveau message, « ci-joint le devis » dans le corps sans pièce jointe → l'avertissement
   doit apparaître. Refaire avec une pièce jointe → l'envoi doit passer sans alerte.

Effet immédiat sur Outlook Web ; jusqu'à 24 h de cache sur Outlook classique Windows.

## Maintenance

Modifier les mots-clés (`MOTIF_PJ`) ou les séparateurs (`SEPARATEURS`) dans `commands.js`,
recommiter sur GitHub, attendre la coche verte dans l'onglet **Actions**, puis recharger
Outlook (Ctrl+F5). Aucune réinstallation tant que le `manifest.xml` ne change pas.

Validation du manifeste (recommandée avant tout chargement) :

```
npm i -g office-addin-manifest
office-addin-manifest validate manifest.xml
```

## Limites connues

- **Dépendance d'hébergement** : GitHub Pages doit rester en ligne et le dépôt public.
  Dépôt privé ou supprimé = add-in inopérant. Aucune donnée sensible dans le dépôt.
- **Coupure de l'historique** : heuristique. Fiable sur les séparateurs usuels d'Outlook
  en français et anglais ; un format de citation exotique peut passer au travers.
- **Mot « devis »** : déclenche seul, même sans intention de pièce jointe. Volontaire ;
  non bloquant grâce au mode `PromptUser`.
- **Objet d'une réponse** : « RE : Devis … » peut hériter d'un mot-clé du message d'origine.
  Point non tranché (options : laisser, ne plus analyser l'objet, ou ne l'analyser que sur
  un message neuf).
- **Latence du premier envoi** : plus lent après ouverture d'Outlook (chargement initial du
  runtime). Le bandeau « traitement en cours » est le comportement normal d'Outlook dès qu'un
  add-in est branché sur l'envoi.
- **Format du manifeste** : XML add-in-only, pleinement supporté et suffisant ici ; Microsoft
  pousse à terme le manifeste unifié JSON.
