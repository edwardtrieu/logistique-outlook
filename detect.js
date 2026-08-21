/* ============================================================
   SOEXIMEX - Evenement logistique
   detect.js : moteur de detection (aucune dependance Office)
   Testable en Node : require('./detect.js')
   ============================================================ */

(function (global) {
  "use strict";

  /* ---------- Configuration ---------- */

  var TYPES = [
    { libelle: "Expedition",      motifs: ["expedition", "expédition", "expedie", "expédié", "expediee", "expédiée", "shipment", "shipping", "depart usine", "départ usine", "etd"] },
    { libelle: "Livraison",       motifs: ["livraison", "livre le", "livré le", "livrer le", "delivery", "eta"] },
    { libelle: "Reception",       motifs: ["reception", "réception", "recu le", "reçu le", "receipt"] },
    { libelle: "Enlevement",      motifs: ["enlevement", "enlèvement", "pickup", "pick-up", "collecte", "ramassage"] },
    { libelle: "Dedouanement",    motifs: ["dedouanement", "dédouanement", "douane", "customs", "declaration en douane", "déclaration en douane"] },
    { libelle: "RDV transitaire", motifs: ["transitaire", "rdv transitaire", "forwarder", "commissionnaire"] },
    { libelle: "Chargement",      motifs: ["chargement", "mise a quai", "mise à quai", "loading", "empotage"] },
    { libelle: "Dechargement",    motifs: ["dechargement", "déchargement", "unloading", "depotage", "dépotage"] },
    { libelle: "Booking",         motifs: ["booking", "reservation navire", "réservation navire", "closing date", "cut-off", "cut off"] }
  ];

  var MOIS = {
    "janvier": 0, "janv": 0, "jan": 0,
    "fevrier": 1, "février": 1, "fevr": 1, "fev": 1, "fév": 1,
    "mars": 2, "mar": 2,
    "avril": 3, "avr": 3,
    "mai": 4,
    "juin": 5,
    "juillet": 6, "juil": 6,
    "aout": 7, "août": 7,
    "septembre": 8, "sept": 8, "sep": 8,
    "octobre": 9, "oct": 9,
    "novembre": 10, "nov": 10,
    "decembre": 11, "décembre": 11, "dec": 11, "déc": 11
  };

  var JOURS_SEMAINE = "lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche|lun|mar|mer|jeu|ven|sam|dim";

  // Moments de journee : convention interne, signalee a l'utilisateur.
  var MOMENTS = [
    /* Ordre imperatif : du plus long au plus court. "midi" est une sous-chaine de
       "apres-midi" et le tiret fait frontiere de mot, donc les composes passent d'abord. */
    { motifs: ["fin de matinee", "fin de matinée"], h: 11, libelle: "fin de matinee" },
    { motifs: ["debut d'apres-midi", "début d'après-midi", "debut d apres-midi"], h: 14, libelle: "debut d'apres-midi" },
    { motifs: ["apres-midi", "après-midi", "apres midi", "après midi"], h: 14, libelle: "apres-midi" },
    { motifs: ["fin de journee", "fin de journée", "soir"], h: 16, libelle: "fin de journee" },
    { motifs: ["matin"],                            h: 9,  libelle: "matin" },
    { motifs: ["midi"],                             h: 12, libelle: "midi" }
  ];

  var HEURE_DEFAUT = 9;
  var DUREE_DEFAUT_MIN = 60;
  var AMPLITUDE_PLAGE_MAX_H = 4;   // au-dela, c'est un horaire d'ouverture, pas un RDV
  var FENETRE_AVANT = 45;          // caracteres avant la date ou chercher l'heure
  var FENETRE_APRES = 90;          // caracteres apres la date ou chercher l'heure

  /* ---------- 1. Nettoyage et isolement du message du dessus ---------- */

  function normaliser(texte) {
    return (texte || "")
      .replace(/ /g, " ")
      .replace(/ /g, " ")
      .replace(/[‐-―]/g, "-")
      .replace(/\r\n?/g, "\n");
  }

  // Coupures de fil de discussion, FR et EN, Outlook classique / nouveau / Web / mobile
  var COUPURES = [
    /\n\s*_{10,}/,
    /\n\s*-{3,}\s*Message d'origine\s*-{0,3}/i,
    /\n\s*-{3,}\s*Original Message\s*-{0,3}/i,
    /\n\s*-{3,}\s*Transf[eé]r[eé] par/i,
    /\n\s*-{3,}\s*Forwarded message/i,
    /\n\s*De\s*:\s/i,
    /\n\s*From\s*:\s/i,
    /\n\s*Exp[eé]diteur\s*:\s/i,
    // "Le mer. 12 aout 2026 a 09:14, X a ecrit :" et variantes
    /\n?\s*Le\s+.{0,60}?\s+a\s+[eé]crit\s*:/i,
    /\n?\s*On\s+.{0,60}?\s+wrote\s*:/i,
    /\n\s*Envoy[eé]\s*(?:de|depuis|à partir de|a partir de)\s/i,
    /\n\s*Sent from\s/i,
    /\n\s*Obtenir Outlook pour/i,
    /\n\s*>{1,}\s/                 // citation par chevrons
  ];

  function messageDuDessus(texte) {
    var fin = texte.length;
    for (var i = 0; i < COUPURES.length; i++) {
      var m = texte.search(COUPURES[i]);
      if (m > -1 && m < fin) { fin = m; }
    }
    return texte.slice(0, fin);
  }

  /* Mail transfere : le contenu utile est SOUS le bloc d'en-tete
     (De: / Envoye: / A: / Objet:). Couper avant ce bloc ne laisse rien a dater,
     et la ligne "Envoye :" porte une date d'expedition qui n'est pas l'evenement.
     On saute donc l'en-tete et on date ce qui suit. */
  function zoneDatable(complet) {
    /* Discriminant : un mail TRANSFERE commence par le bloc d'en-tete, alors qu'un
       fil de REPONSE commence par le texte de la reponse et cite l'en-tete plus bas.
       Dans le second cas les dates citees sont perimees et ne doivent pas etre datees
       (l'utilisateur garde le bouton d'elargissement au fil). */
    var debutEnTete = /^\s*(?:De|From|Exp[eé]diteur)[ \t]*:/i.test(complet.slice(0, 200));
    if (!debutEnTete) {
      return { texte: messageDuDessus(complet), transfert: false };
    }

    // Dernier "Objet :" / "Subject:" de la zone d'en-tete
    var enTete = complet.slice(0, 2000);
    var rx = /(?:^|\n)[ \t]*(?:Objet|Subject)[ \t]*:[^\n]*\n/gi;
    var m, dernier = null;
    while ((m = rx.exec(enTete)) !== null) { dernier = m; }
    if (!dernier) {
      return { texte: complet, transfert: true };
    }

    var apres = complet.slice(dernier.index + dernier[0].length);
    return { texte: messageDuDessus("\n" + apres).slice(1), transfert: true };
  }

  /* ---------- 2. Masquage des faux positifs ----------
     On remplace les zones piegeuses par des '#' de meme longueur :
     les positions restent valides pour l'affichage de l'extrait. */

  var PIEGES = [
    // Ceinture de securite : une ligne d'en-tete "Envoye :" porte la date d'expedition
    // du mail, jamais celle de l'evenement. Neutralisee meme si le decoupage a echoue.
    /(?:^|\n)[ \t]*(?:Envoy[eé]|Sent)[ \t]*:[^\n]*/gi,
    // Telephone / fax annonce par un libelle
    /(?:t[eé]l|tel|mob|portable|fax|gsm|phone|whatsapp)\s*\.?\s*:?\s*\+?[\d\s.\-()]{8,}/gi,
    // Numero francais 0X XX XX XX XX (points, tirets, espaces)
    /\b0\d(?:[\s.\-]?\d{2}){4}\b/g,
    // International +33 ..., +213 ...
    /\+\d{1,3}(?:[\s.\-]?\d{2,3}){3,}/g,
    // Montants
    /\d[\d\s.,]*\s*(?:EUR|USD|DZD|EGP|TRY|€|\$)\b/gi,
    /(?:EUR|USD|DZD|EGP|TRY|€|\$)\s*\d[\d\s.,]*/gi,
    // Numeros de version 1.5.2 / v2.0.1
    /\bv?\d+\.\d+(?:\.\d+)+\b/gi,
    // Reference SOEXIMEX client/fournisseur (traitee a part comme reference)
    /\b[A-Z]{2}\d{5}\b/g,
    // Numero de conteneur ISO : 4 lettres + 7 chiffres
    /\b[A-Z]{4}\d{7}\b/g,
    // Reference dossier type PL-12-08-26 / BL/12/08/2026
    /\b[A-Z]{2,5}\s*[-_\/]\s*\d{1,4}\s*[-_\/]\s*\d{1,4}\s*[-_\/]\s*\d{2,4}\b/g,
    // Numeros de documents : BL n°123456, INV 2026-0045, AWB 020-12345678
    /\b(?:BL|B\/L|AWB|CMR|INV|FAC|FA|PL|PO|CMD|DO|HS)\s*(?:n\s*[°ºo]?)?\s*[:.]?\s*[A-Z0-9][A-Z0-9\-\/]{3,}/gi,
    // Poids / dimensions / quantites suivies d'unite
    /\b\d+[\d.,]*\s*(?:kg|kgs|t|tonnes?|cbm|m3|m²|palettes?|colis|cartons?|pcs|pi[eè]ces?)\b/gi,
    // Codes postaux + ville non geres, mais on masque les longues suites de chiffres
    /\b\d{6,}\b/g
  ];

  function masquer(texte) {
    var t = texte;
    for (var i = 0; i < PIEGES.length; i++) {
      t = t.replace(PIEGES[i], function (m) {
        return new Array(m.length + 1).join("#");
      });
    }
    return t;
  }

  /* ---------- 3. Dates candidates ---------- */

  function dateValide(annee, mois, jour) {
    if (annee < 2020 || annee > 2100) { return null; }
    if (mois < 0 || mois > 11) { return null; }
    if (jour < 1 || jour > 31) { return null; }
    var d = new Date(annee, mois, jour, 12, 0, 0, 0);
    // round-trip : rejette 31/02, 31/04, etc.
    if (d.getFullYear() !== annee || d.getMonth() !== mois || d.getDate() !== jour) { return null; }
    return d;
  }

  function trouverDates(masque, aujourdhui) {
    var out = [];
    var m, d;

    // a) ISO : 2026-08-24
    var rxIso = /\b(20\d{2})-(\d{1,2})-(\d{1,2})\b/g;
    while ((m = rxIso.exec(masque)) !== null) {
      d = dateValide(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
      if (d) { out.push({ date: d, index: m.index, longueur: m[0].length, format: "iso", anneeExplicite: true }); }
    }

    // b) Numerique FR : 24/08/2026, 24-08-26, 24.08.2026
    var rxNum = /\b(\d{1,2})\s*[\/\.\-]\s*(\d{1,2})\s*[\/\.\-]\s*(\d{2,4})\b/g;
    while ((m = rxNum.exec(masque)) !== null) {
      var an = parseInt(m[3], 10);
      if (an < 100) { an += 2000; }
      d = dateValide(an, parseInt(m[2], 10) - 1, parseInt(m[1], 10));
      if (d) { out.push({ date: d, index: m.index, longueur: m[0].length, format: "numerique", anneeExplicite: true }); }
    }

    // c) Numerique court sans annee : 24/08, precede d'un declencheur ou d'un jour
    //    de semaine ("livre lundi 24/08 matin" est la formulation la plus courante
    //    des transporteurs).
    var rxCourt = new RegExp("\\b(?:le|du|au|d[eè]s|avant|apr[eè]s|pour|" + JOURS_SEMAINE +
                             ")\\.?\\s+(\\d{1,2})\\s*[\\/\\.]\\s*(\\d{1,2})(?!\\s*[\\/\\.\\-]\\s*\\d)", "gi");
    while ((m = rxCourt.exec(masque)) !== null) {
      var res = resoudreAnnee(parseInt(m[2], 10) - 1, parseInt(m[1], 10), aujourdhui);
      if (res) { out.push({ date: res, index: m.index, longueur: m[0].length, format: "numerique-court", anneeExplicite: false }); }
    }

    // d) Textuel : 24 aout 2026 / 1er septembre / 24 aout
    var noms = Object.keys(MOIS).sort(function (a, b) { return b.length - a.length; }).join("|");
    var rxTxt = new RegExp("\\b(\\d{1,2})\\s*(?:er)?\\s+(" + noms + ")\\.?(?:\\s+(20\\d{2}))?", "gi");
    while ((m = rxTxt.exec(masque)) !== null) {
      var mois = MOIS[m[2].toLowerCase()];
      var jour = parseInt(m[1], 10);
      if (m[3]) {
        d = dateValide(parseInt(m[3], 10), mois, jour);
        if (d) { out.push({ date: d, index: m.index, longueur: m[0].length, format: "textuel", anneeExplicite: true }); }
      } else {
        d = resoudreAnnee(mois, jour, aujourdhui);
        if (d) { out.push({ date: d, index: m.index, longueur: m[0].length, format: "textuel", anneeExplicite: false }); }
      }
    }

    return dedoublonner(out);
  }

  // Annee absente : on ne prend l'annee suivante que si la date est deja passee
  function resoudreAnnee(mois, jour, aujourdhui) {
    var d = dateValide(aujourdhui.getFullYear(), mois, jour);
    if (!d) { return null; }
    var minuitAuj = new Date(aujourdhui.getFullYear(), aujourdhui.getMonth(), aujourdhui.getDate(), 12, 0, 0, 0);
    if (d < minuitAuj) {
      d = dateValide(aujourdhui.getFullYear() + 1, mois, jour);
    }
    return d;
  }

  function dedoublonner(liste) {
    var vus = {}, out = [];
    liste.sort(function (a, b) { return a.index - b.index; });
    for (var i = 0; i < liste.length; i++) {
      var cle = liste[i].date.getTime();
      if (!vus[cle]) { vus[cle] = true; out.push(liste[i]); }
      // chevauchement d'index : on garde le premier trouve
    }
    return out;
  }

  /* ---------- 4. Heure par proximite ---------- */

  function trouverHeure(masque, indexDate, longueurDate) {
    var debut = Math.max(0, indexDate - FENETRE_AVANT);
    var fin = Math.min(masque.length, indexDate + longueurDate + FENETRE_APRES);
    var zone = masque.slice(debut, fin);

    /* Les plages sont retirees de la zone AVANT toute autre recherche, pour ne pas
       masquer une heure annoncee situee plus loin ("... a 14h30 ... joignables 8h-18h").
       Une plage courte reste exploitable en repli ; une plage large est un horaire
       d'ouverture et n'est jamais retenue. */
    var plagesCourtes = [];
    var notePlage = "";
    var rxPlage = /\b(\d{1,2})\s*(?:h|:)\s*(\d{2})?\s*(?:-|a|à|et|jusqu'?a|jusqu'?à)\s*(\d{1,2})\s*(?:h|:)\s*(\d{2})?/gi;
    zone = zone.replace(rxPlage, function (texte, h1, min1, h2) {
      var a = parseInt(h1, 10), b = parseInt(h2, 10);
      if (isNaN(a) || isNaN(b) || a < 0 || a > 23 || b < 0 || b > 23) { return texte; }
      if (b - a > 0 && b - a <= AMPLITUDE_PLAGE_MAX_H) {
        plagesCourtes.push({ h: a, min: min1 ? parseInt(min1, 10) : 0, borne: a + "h-" + b + "h" });
      } else {
        notePlage = "Amplitude " + a + "h-" + b + "h ignoree (horaire d'ouverture).";
      }
      return new Array(texte.length + 1).join("#");
    });

    // Heure precise annoncee par un declencheur : a 14h30, vers 15h, avant 10h
    var rxAnnoncee = /(?:\b(?:a|à|vers|avant|d[eè]s|apr[eè]s|au plus tard(?: a| à)?|rdv|rendez-vous)\s+)(\d{1,2})\s*(?:h|:)\s*(\d{2})?/i;
    var ma = zone.match(rxAnnoncee);
    if (ma) {
      var ha = parseInt(ma[1], 10);
      if (ha >= 0 && ha <= 23) {
        var mina = ma[2] ? parseInt(ma[2], 10) : 0;
        if (mina >= 0 && mina <= 59) { return { h: ha, min: mina, source: "annoncee", note: notePlage }; }
      }
    }

    // Repli : plage courte, on retient la borne basse
    if (plagesCourtes.length) {
      return { h: plagesCourtes[0].h, min: plagesCourtes[0].min, source: "plage",
               note: "Plage " + plagesCourtes[0].borne + " detectee, debut retenu." };
    }

    // Heure nue dans la fenetre : 14h30 / 14:30 (minutes obligatoires pour ':')
    var rxNue = /\b(\d{1,2})\s*(?:h\s*(\d{2})?|:\s*(\d{2}))(?!\s*(?:-|a\s|à\s)\s*\d)/i;
    var mn = zone.match(rxNue);
    if (mn) {
      var hn = parseInt(mn[1], 10);
      var minn = mn[2] !== undefined ? parseInt(mn[2], 10) : (mn[3] !== undefined ? parseInt(mn[3], 10) : 0);
      if (hn >= 0 && hn <= 23 && minn >= 0 && minn <= 59) {
        return { h: hn, min: minn, source: "proximite", note: "Heure lue a proximite de la date, a verifier." };
      }
    }

    // Moment de journee : "lundi 24/08 matin", "livraison apres-midi"
    var zoneMin = zone.toLowerCase();
    for (var i = 0; i < MOMENTS.length; i++) {
      for (var j = 0; j < MOMENTS[i].motifs.length; j++) {
        var mot = MOMENTS[i].motifs[j];
        var rxMot = new RegExp("\\b" + mot.replace(/[-']/g, "\\$&") + "\\b", "i");
        if (rxMot.test(zoneMin)) {
          return { h: MOMENTS[i].h, min: 0, source: "moment",
                   note: "Mail indique \"" + MOMENTS[i].libelle + "\", " + MOMENTS[i].h + "h retenu par convention." };
        }
      }
    }

    return { h: HEURE_DEFAUT, min: 0, source: "defaut", note: "Heure absente du mail, " + HEURE_DEFAUT + "h par defaut." };
  }

  /* ---------- 5. Type d'evenement ---------- */

  function trouverTypes(objet, corps) {
    var res = [];
    var zones = [{ t: (objet || "").toLowerCase(), poids: 100 }, { t: (corps || "").toLowerCase(), poids: 10 }];
    for (var i = 0; i < TYPES.length; i++) {
      var meilleur = null;
      for (var z = 0; z < zones.length; z++) {
        for (var j = 0; j < TYPES[i].motifs.length; j++) {
          var p = zones[z].t.indexOf(TYPES[i].motifs[j]);
          if (p > -1) {
            var score = zones[z].poids - Math.min(p / 100, 9);
            if (!meilleur || score > meilleur.score) { meilleur = { score: score, position: p, zone: z }; }
          }
        }
      }
      if (meilleur) { res.push({ libelle: TYPES[i].libelle, score: meilleur.score, position: meilleur.position, zone: meilleur.zone }); }
    }
    res.sort(function (a, b) { return b.score - a.score; });
    return res;
  }

  /* ---------- 6. Reference et lieu ---------- */

  function trouverReferences(objet, corps) {
    var out = [];
    var rxRef = /\b[A-Z]{2}\d{5}\b/g;
    var rxCont = /\b[A-Z]{4}\d{7}\b/g;
    [objet || "", corps || ""].forEach(function (src) {
      var m;
      while ((m = rxRef.exec(src)) !== null) { if (out.indexOf(m[0]) === -1) { out.push(m[0]); } }
      while ((m = rxCont.exec(src)) !== null) { if (out.indexOf(m[0]) === -1) { out.push(m[0]); } }
    });
    return out;
  }

  function trouverLieu(corps) {
    var rx = /\b(?:port de|port d'|terminal|quai|entrep[oô]t|plateforme|depot|dépôt|agence|usine|magasin)\s+[A-Za-zÀ-ÿ0-9'\-\s]{2,40}/i;
    var m = (corps || "").match(rx);
    return m ? m[0].replace(/\s+/g, " ").trim() : "";
  }

  /* ---------- 7. Scoring des candidats ---------- */

  var MOTS_DECLENCHEURS = ["prevu", "prévu", "prevue", "prévue", "planifie", "planifié", "confirme", "confirmé",
                           "le ", "du ", "date de", "eta", "etd", "au plus tard", "livraison", "expedition",
                           "expédition", "enlevement", "enlèvement", "chargement", "reception", "réception",
                           "rdv", "rendez-vous", "cut-off", "closing"];

  // Une date qui SUIT un mot de report est la nouvelle date : elle gagne.
  var MOTS_REPORT_AVANT = ["report", "decal", "décal", "avanc", "nouvelle date", "au lieu du",
                           "remplace", "modifie", "modifié", "finalement", "desormais", "désormais"];
  // Une date SUIVIE d'un mot de report ou d'annulation est l'ancienne : elle perd.
  var MOTS_REPORT_APRES = ["report", "decal", "décal", "annul", "n'est plus", "caduc", "initialement"];

  function scorer(candidat, masque) {
    var s = 0;
    var avant = masque.slice(Math.max(0, candidat.index - 60), candidat.index).toLowerCase();
    var apres = masque.slice(candidat.index + candidat.longueur, candidat.index + candidat.longueur + 45).toLowerCase();
    var contexte = avant + apres;

    for (var i = 0; i < MOTS_DECLENCHEURS.length; i++) {
      if (contexte.indexOf(MOTS_DECLENCHEURS[i]) > -1) { s += 8; }
    }
    for (var j = 0; j < MOTS_REPORT_AVANT.length; j++) {
      if (avant.indexOf(MOTS_REPORT_AVANT[j]) > -1) { s += 35; break; }
    }
    for (var k = 0; k < MOTS_REPORT_APRES.length; k++) {
      if (apres.indexOf(MOTS_REPORT_APRES[k]) > -1) { s -= 30; break; }
    }

    if (candidat.anneeExplicite) { s += 6; }
    if (candidat.format === "textuel" || candidat.format === "numerique") { s += 4; }
    // a defaut d'autre signal, une date plus haut dans le message compte davantage
    s += Math.max(0, 20 - Math.floor(candidat.index / 40));
    return s;
  }

  function extrait(texteOriginal, candidat) {
    var debut = Math.max(0, candidat.index - 60);
    var fin = Math.min(texteOriginal.length, candidat.index + candidat.longueur + 60);
    // couper aux frontieres de mots pour ne pas laisser de fragment isole
    if (debut > 0) {
      var esp = texteOriginal.slice(debut, candidat.index).search(/[\s\n]/);
      if (esp > -1) { debut += esp + 1; }
    }
    if (fin < texteOriginal.length) {
      var reste = texteOriginal.slice(candidat.index + candidat.longueur, fin);
      var dernier = Math.max(reste.lastIndexOf(" "), reste.lastIndexOf("\n"));
      if (dernier > 0) { fin = candidat.index + candidat.longueur + dernier; }
    }
    var s = texteOriginal.slice(debut, fin).replace(/\s+/g, " ").trim();
    return (debut > 0 ? "... " : "") + s + (fin < texteOriginal.length ? " ..." : "");
  }

  /* ---------- 8. Point d'entree ---------- */

  function analyser(objet, corpsComplet, maintenant, options) {
    var aujourdhui = maintenant || new Date();
    var opt = options || {};
    objet = normaliser(objet);
    var complet = normaliser(corpsComplet);
    var zone = zoneDatable(complet);
    var haut = zone.texte;

    // Le type peut venir de tout le fil, la date seulement du message du dessus + objet.
    var types = trouverTypes(objet, complet);

    // Par defaut on ne date que le message du dessus : les dates du fil cite
    // sont presque toujours perimees. inclureFil = recherche elargie sur demande.
    var texteDate = objet + "\n" + (opt.inclureFil ? complet : haut);
    var masque = masquer(texteDate);
    var brutes = trouverDates(masque, aujourdhui);

    var candidats = brutes.map(function (c) {
      var heure = trouverHeure(masque, c.index, c.longueur);
      var d = new Date(c.date.getTime());
      d.setHours(heure.h, heure.min, 0, 0);
      return {
        debut: d,
        fin: new Date(d.getTime() + DUREE_DEFAUT_MIN * 60000),
        format: c.format,
        anneeExplicite: c.anneeExplicite,
        heureSource: heure.source,
        note: heure.note,
        extrait: extrait(texteDate, c),
        score: scorer(c, masque)
      };
    });

    candidats.sort(function (a, b) { return b.score - a.score; });

    return {
      types: types.map(function (t) { return t.libelle; }),
      typePrincipal: types.length ? types[0].libelle : "",
      references: trouverReferences(objet, haut),
      lieu: trouverLieu(haut),
      candidats: candidats,
      filTronque: haut.length < complet.length,
      filInclus: !!opt.inclureFil,
      dureeDefautMin: DUREE_DEFAUT_MIN
    };
  }

  var api = {
    analyser: analyser,
    // exposes pour les tests unitaires
    _messageDuDessus: messageDuDessus,
    _masquer: masquer,
    _normaliser: normaliser,
    TYPES: TYPES,
    HEURE_DEFAUT: HEURE_DEFAUT,
    DUREE_DEFAUT_MIN: DUREE_DEFAUT_MIN
  };

  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  global.LOGI = api;

})(typeof window !== "undefined" ? window : globalThis);
