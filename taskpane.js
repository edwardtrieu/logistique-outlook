/* ============================================================
   SOEXIMEX - Evenement logistique
   taskpane.js : liaison Office <-> moteur de detection <-> formulaire
   ============================================================ */

(function () {
  "use strict";

  var etat = { objet: "", corps: "", analyse: null, filInclus: false };

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) { return; }
    lireEtAnalyser(false);

    document.getElementById("btnCreer").addEventListener("click", creerRdv);
    document.getElementById("btnFil").addEventListener("click", function () {
      etat.filInclus = true;
      analyserEtAfficher();
    });
  });

  /* ---------- Lecture du mail ---------- */

  function lireEtAnalyser() {
    var item = Office.context.mailbox.item;
    if (!item) { erreurFatale("Aucun mail ouvert."); return; }

    etat.objet = item.subject || "";
    item.body.getAsync(Office.CoercionType.Text, function (res) {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        erreurFatale("Lecture du corps du mail impossible : " + (res.error ? res.error.message : "erreur inconnue"));
        return;
      }
      etat.corps = res.value || "";
      analyserEtAfficher();
      document.getElementById("chargement").style.display = "none";
      document.getElementById("contenu").style.display = "block";
    });
  }

  function erreurFatale(msg) {
    var c = document.getElementById("chargement");
    c.innerHTML = '<div class="alerte">' + echapper(msg) + "</div>";
  }

  /* ---------- Analyse et rendu ---------- */

  function analyserEtAfficher() {
    var a = LOGI.analyser(etat.objet, etat.corps, new Date(), { inclureFil: etat.filInclus });
    etat.analyse = a;

    remplirMessages(a);
    remplirCandidats(a);
    remplirTypes(a);

    document.getElementById("lieu").value = a.lieu || "";
    document.getElementById("ref").value = a.references.length ? a.references[0] : "";

    if (a.candidats.length) {
      appliquerCandidat(0);
    } else {
      var d = new Date();
      d.setDate(d.getDate() + 1);
      document.getElementById("date").value = isoDate(d);
      document.getElementById("heure").value = pad(LOGI.HEURE_DEFAUT) + ":00";
      majTitre();
    }
  }

  function remplirMessages(a) {
    var out = [];

    if (!a.typePrincipal) {
      out.push('<div class="alerte">Aucun mot-cle logistique reconnu dans ce mail. Choisissez le type a la main si besoin.</div>');
    }

    if (!a.candidats.length) {
      if (a.filTronque && !a.filInclus) {
        out.push('<div class="alerte">Aucune date exploitable dans le message du dessus. Le mail contient un fil cite : vous pouvez elargir la recherche.</div>');
      } else {
        out.push('<div class="alerte">Aucune date exploitable trouvee. Saisissez-la a la main.</div>');
      }
    } else if (a.candidats.length > 1) {
      out.push('<div class="alerte">' + a.candidats.length + ' dates possibles dans ce mail. Verifiez celle qui correspond a l\'evenement.</div>');
    }

    if (a.filInclus) {
      out.push('<div class="alerte">Recherche elargie au fil cite : les dates ci-dessous peuvent etre perimees.</div>');
    }

    if (a.references.length > 1) {
      out.push('<div class="alerte">Plusieurs references detectees : ' + a.references.map(echapper).join(", ") + ". Corrigez si besoin.</div>");
    }

    document.getElementById("messages").innerHTML = out.join("");

    var btn = document.getElementById("btnFil");
    btn.style.display = (a.filTronque && !a.filInclus) ? "block" : "none";
  }

  function remplirCandidats(a) {
    var box = document.getElementById("candidats");
    if (!a.candidats.length) {
      box.innerHTML = '<div class="info">Rien de detecte. Renseignez la date dans le bloc ci-dessous.</div>';
      return;
    }
    var html = a.candidats.map(function (c, i) {
      var notes = [];
      if (c.note) { notes.push(c.note); }
      if (!c.anneeExplicite) { notes.push("Annee absente du mail, deduite."); }
      return '<label class="cand' + (i === 0 ? " sel" : "") + '" data-i="' + i + '">' +
        '<input type="radio" name="cand" value="' + i + '"' + (i === 0 ? " checked" : "") + ' />' +
        '<span class="d">' + libelleDate(c.debut) + "</span>" +
        '<span class="x">' + echapper(c.extrait) + "</span>" +
        (notes.length ? '<span class="n">' + echapper(notes.join(" ")) + "</span>" : "") +
        "</label>";
    }).join("");
    box.innerHTML = html;

    Array.prototype.forEach.call(box.querySelectorAll('input[name="cand"]'), function (r) {
      r.addEventListener("change", function () {
        Array.prototype.forEach.call(box.querySelectorAll(".cand"), function (l) { l.classList.remove("sel"); });
        r.parentNode.classList.add("sel");
        appliquerCandidat(parseInt(r.value, 10));
      });
    });
  }

  function remplirTypes(a) {
    var sel = document.getElementById("type");
    sel.innerHTML = "";
    var liste = LOGI.TYPES.map(function (t) { return t.libelle; });
    liste.forEach(function (lib) {
      var o = document.createElement("option");
      o.value = lib; o.textContent = lib;
      if (lib === a.typePrincipal) { o.selected = true; }
      sel.appendChild(o);
    });
    var autre = document.createElement("option");
    autre.value = "Evenement logistique"; autre.textContent = "Autre / non precise";
    if (!a.typePrincipal) { autre.selected = true; }
    sel.appendChild(autre);
    sel.addEventListener("change", majTitre);
    document.getElementById("ref").addEventListener("input", majTitre);
  }

  function appliquerCandidat(i) {
    var c = etat.analyse.candidats[i];
    if (!c) { return; }
    document.getElementById("date").value = isoDate(c.debut);
    document.getElementById("heure").value = pad(c.debut.getHours()) + ":" + pad(c.debut.getMinutes());
    majTitre();
  }

  function majTitre() {
    var type = document.getElementById("type").value;
    var ref = document.getElementById("ref").value.trim();
    document.getElementById("titre").value = ref ? type + " " + ref : type;
  }

  /* ---------- Creation du rendez-vous ---------- */

  function creerRdv() {
    var dateStr = document.getElementById("date").value;
    var heureStr = document.getElementById("heure").value || "09:00";
    if (!dateStr) {
      alerter("Renseignez une date avant de creer le rendez-vous.");
      return;
    }

    var p = dateStr.split("-");
    var hm = heureStr.split(":");
    var debut = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10),
                         parseInt(hm[0], 10), parseInt(hm[1], 10), 0, 0);
    if (isNaN(debut.getTime())) { alerter("Date ou heure invalide."); return; }

    var duree = parseInt(document.getElementById("duree").value, 10);
    var journee = duree === 1440;
    var fin = new Date(debut.getTime() + (journee ? 1440 : duree) * 60000);
    if (journee) { debut.setHours(0, 0, 0, 0); fin = new Date(debut.getTime() + 1440 * 60000); }

    var titre = document.getElementById("titre").value.trim() || "Evenement logistique";
    var lieu = document.getElementById("lieu").value.trim();
    var ref = document.getElementById("ref").value.trim();

    var corpsRdv = [
      "Cree depuis le mail : " + etat.objet,
      "",
      "Reference : " + (ref || "non renseignee"),
      "Type : " + document.getElementById("type").value,
      lieu ? "Lieu : " + lieu : null,
      "",
      "-- Complement Evenement logistique / SOEXIMEX --"
    ].filter(function (l) { return l !== null; }).join("\n");

    var params = {
      requiredAttendees: [],
      subject: titre.slice(0, 255),
      start: debut,
      end: fin,
      body: corpsRdv
    };
    if (lieu) { params.location = lieu.slice(0, 255); }

    var mb = Office.context.mailbox;

    /* La variante asynchrone (Mailbox 1.9) est la seule fiable sur le nouveau Outlook
       et sur Outlook Web ; la version synchrone peut echouer sans lever d'erreur.
       Dans tous les cas on propose un repli .ics, car l'ouverture du formulaire
       passe par une fenetre que le navigateur peut bloquer. */
    var asyncDispo = false;
    try {
      asyncDispo = typeof mb.displayNewAppointmentFormAsync === "function" &&
                   Office.context.requirements &&
                   Office.context.requirements.isSetSupported("Mailbox", "1.9");
    } catch (e) { asyncDispo = false; }

    if (asyncDispo) {
      mb.displayNewAppointmentFormAsync(params, function (res) {
        if (res && res.status === Office.AsyncResultStatus.Succeeded) {
          succesAvecIcs("Formulaire de rendez-vous ouvert. Enregistrez-le dans Outlook pour le confirmer.", params);
        } else {
          alerterAvecIcs("Outlook a refuse l'ouverture du formulaire" +
            (res && res.error && res.error.message ? " (" + res.error.message + ")" : "") + ".", params);
        }
      });
      return;
    }

    try {
      mb.displayNewAppointmentForm(params);
      succesAvecIcs("Formulaire de rendez-vous ouvert. Enregistrez-le dans Outlook pour le confirmer.", params);
    } catch (e) {
      alerterAvecIcs("Ouverture du formulaire impossible : " +
        (e && e.message ? e.message : "erreur inconnue") + ".", params);
    }
  }

  /* ---------- Repli : invitation .ics telechargeable ---------- */

  function horodatageIcs(d, utc) {
    function p(n) { return (n < 10 ? "0" : "") + n; }
    if (utc) {
      return d.getUTCFullYear() + p(d.getUTCMonth() + 1) + p(d.getUTCDate()) + "T" +
             p(d.getUTCHours()) + p(d.getUTCMinutes()) + "00Z";
    }
    // heure locale flottante : Outlook l'interprete dans le fuseau du poste
    return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + "T" +
           p(d.getHours()) + p(d.getMinutes()) + "00";
  }

  function echapperIcs(s) {
    return String(s || "").replace(/\\/g, "\\\\").replace(/;/g, "\\;")
                          .replace(/,/g, "\\,").replace(/\r?\n/g, "\\n");
  }

  function construireIcs(params) {
    var uid = "logi-" + horodatageIcs(params.start, true) + "-" +
              Math.abs(String(params.subject).split("").reduce(function (a, c) {
                return ((a << 5) - a + c.charCodeAt(0)) | 0;
              }, 0)) + "@soeximex";
    var lignes = [
      "BEGIN:VCALENDAR",
      "VERSION:2.0",
      "PRODID:-//SOEXIMEX//Evenement logistique//FR",
      "CALSCALE:GREGORIAN",
      "METHOD:PUBLISH",
      "BEGIN:VEVENT",
      "UID:" + uid,
      "DTSTAMP:" + horodatageIcs(new Date(), true),
      "DTSTART:" + horodatageIcs(params.start, false),
      "DTEND:" + horodatageIcs(params.end, false),
      "SUMMARY:" + echapperIcs(params.subject)
    ];
    if (params.location) { lignes.push("LOCATION:" + echapperIcs(params.location)); }
    lignes.push("DESCRIPTION:" + echapperIcs(params.body));
    lignes.push("END:VEVENT", "END:VCALENDAR");
    return lignes.join("\r\n") + "\r\n";
  }

  function lienIcs(params) {
    try {
      var blob = new Blob([construireIcs(params)], { type: "text/calendar;charset=utf-8" });
      var url = URL.createObjectURL(blob);
      var nom = String(params.subject).replace(/[^A-Za-z0-9 _-]/g, "").trim().replace(/\s+/g, "_") || "evenement";
      return '<a class="lien-ics" href="' + url + '" download="' + nom + '.ics">' +
             "Telecharger l'invitation (.ics)</a>";
    } catch (e) {
      return "";
    }
  }

  function succesAvecIcs(msg, params) {
    document.getElementById("messages").innerHTML =
      '<div class="ok">' + echapper(msg) +
      '<div class="repli">Aucune fenetre ne s\'est ouverte ? Votre navigateur bloque les fenetres surgissantes. ' +
      lienIcs(params) + " puis ouvrez le fichier.</div></div>";
  }

  function alerterAvecIcs(msg, params) {
    document.getElementById("messages").innerHTML =
      '<div class="alerte">' + echapper(msg) +
      '<div class="repli">Solution de contournement : ' + lienIcs(params) +
      " puis ouvrez le fichier, Outlook creera le rendez-vous.</div></div>";
  }

  /* ---------- Utilitaires ---------- */

  function alerter(msg) {
    document.getElementById("messages").innerHTML = '<div class="alerte">' + echapper(msg) + "</div>";
  }
  function succes(msg) {
    document.getElementById("messages").innerHTML = '<div class="ok">' + echapper(msg) + "</div>";
  }
  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function isoDate(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }

  var JOURS = ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"];
  var MOIS_L = ["janvier", "fevrier", "mars", "avril", "mai", "juin", "juillet", "aout", "septembre", "octobre", "novembre", "decembre"];
  function libelleDate(d) {
    return JOURS[d.getDay()] + " " + d.getDate() + " " + MOIS_L[d.getMonth()] + " " + d.getFullYear() +
           " a " + pad(d.getHours()) + "h" + pad(d.getMinutes());
  }
  function echapper(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

})();
