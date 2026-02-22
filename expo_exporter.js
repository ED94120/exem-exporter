(async () => {

  const SCRIPT_VERSION = "EXPO_CAPTEUR_V1_2026_02_22";
  const SEUIL_EXPO_MAX = 10.0;        // si E >= 10 V/m => Exposition vidée
  const SEUIL_DELTA_MINUTES = 30;     // si delta <= 30 min (entre 2 mesures valides) => Exposition vidée

  // --------------------------
  // Utils dates / nombres
  // --------------------------

  function parseFRDate(s) {
    // Accepte : "dd/mm/yyyy hh:mm" ou "dd/mm/yyyy hh:mm:ss"
    const m = String(s || "").trim().match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})(?::(\d{2}))?$/);
    if (!m) return null;
    const sec = m[6] ? +m[6] : 0;
    return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], sec, 0);
  }

  function fmtFRDate(d) {
    const p = n => String(n).padStart(2, "0");
    return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
  }

  function fmtFRNumber(n) {
    return Number(n).toFixed(2).replace(".", ",");
  }

  function fmtCompactLocal(d) {
    const p = n => String(n).padStart(2, "0");
    return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
  }

  function sanitizeFileName(s) {
    return String(s || "")
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-zA-Z0-9._-]+/g, "_")
      .replace(/^_+|_+$/g, "")
      .slice(0, 120);
  }

  // --------------------------
  // Saisie
  // --------------------------

  function askText(label, current) {
    const v = prompt(label, current == null ? "" : String(current));
    if (v === null) return null;      // cancel = inconnu
    return String(v);                 // vide autorisé
  }

  function askDate(label, currentStr) {
    const v = prompt(label, currentStr || "");
    if (v === null) return { str: null, date: null };     // cancel = inconnu
    const s = String(v).trim();
    if (s === "") return { str: "", date: null };         // vide = inconnu
    const d = parseFRDate(s);
    if (!d) {
      alert("Date invalide. Format attendu : dd/mm/yyyy hh:mm");
      return askDate(label, currentStr);
    }
    return { str: s, date: d };
  }

  function isMissingText(v) {
    return (v == null) || (String(v).trim() === "");
  }

  function isMissingDate(d) {
    return !(d instanceof Date) || isNaN(d.getTime());
  }

  function getMissingFields(P) {
    const miss = [];
    if (isMissingText(P.reference)) miss.push("Référence");
    if (isMissingText(P.adresse)) miss.push("Adresse");
    if (isMissingDate(P.DateDeb)) miss.push("Date début");
    if (isMissingDate(P.DateFin)) miss.push("Date fin");
    return miss;
  }

  function buildRecap(P) {
    const show = v => (isMissingText(v) ? "NON RENSEIGNÉ" : String(v));
    const showDate = (s, d) => (isMissingDate(d) ? "NON RENSEIGNÉ" : s);
    const yesno = b => (b ? "OUI" : "NON");

    const missing = getMissingFields(P);

    return [
      "Récapitulatif",
      "--------------------------------",
      `Référence : ${show(P.reference)}`,
      `Adresse   : ${show(P.adresse)}`,
      `Date début: ${showDate(P.sDateDeb, P.DateDeb)}`,
      `Date fin  : ${showDate(P.sDateFin, P.DateFin)}`,
      `Pixels bruts : ${yesno(P.archiverPixels)}`,
      "",
      missing.length ? ("Champs manquants : " + missing.join(", ")) : "Tous les champs sont renseignés.",
      "",
      "OK = Lancer (si complet) / Annuler = Modifier"
    ].join("\n");
  }

  function editOneField(P) {
    const choice = prompt(
      "Modifier quel champ ?\n" +
      "1 Référence\n" +
      "2 Adresse\n" +
      "3 Date début\n" +
      "4 Date fin\n" +
      "8 Pixels bruts\n" +
      "0 Retour\n" +
      "00 EXIT (abandonner)",
      "0"
    );

    if (choice === null) return true; // Cancel = retour synthèse

    const c = String(choice).trim();

    if (c === "00") return false;   // EXIT complet
    if (c === "0") return true;     // retour synthèse

    if (c === "1") { P.reference = askText("Référence Capteur (ex: Site #Nantes_46) :", P.reference); return true; }
    if (c === "2") { P.adresse = askText("Adresse Capteur :", P.adresse); return true; }

    if (c === "3") {
      const r = askDate("Date début (dd/mm/yyyy hh:mm) :", P.sDateDeb);
      P.sDateDeb = r.str; P.DateDeb = r.date;
      return true;
    }

    if (c === "4") {
      const r = askDate("Date fin (dd/mm/yyyy hh:mm) :", P.sDateFin);
      P.sDateFin = r.str; P.DateFin = r.date;
      return true;
    }

    if (c === "8") { P.archiverPixels = confirm("Archiver pixels bruts ?"); return true; }

    alert("Choix invalide.");
    return true;
  }

  function collectAndConfirmUserInputs() {

    const P = {
      reference: null,
      adresse: null,
      sDateDeb: null,
      sDateFin: null,
      DateDeb: null,
      DateFin: null,
      archiverPixels: false
    };

    P.reference = askText("Référence Capteur (ex: Site #Nantes_46) :", "");
    P.adresse = askText("Adresse Capteur :", "");

    {
      const rDeb = askDate("Date début (dd/mm/yyyy hh:mm) :", "");
      P.sDateDeb = rDeb.str; P.DateDeb = rDeb.date;
    }
    {
      const rFin = askDate("Date fin (dd/mm/yyyy hh:mm) :", "");
      P.sDateFin = rFin.str; P.DateFin = rFin.date;
    }

    P.archiverPixels = confirm("Archiver pixels bruts ?");

    while (true) {

      if (!isMissingDate(P.DateDeb) && !isMissingDate(P.DateFin)) {
        if (P.DateFin.getTime() <= P.DateDeb.getTime()) {
          alert("Erreur : Date fin doit être strictement après Date début.");
          const rFin = askDate("Date fin (dd/mm/yyyy hh:mm) :", P.sDateFin);
          P.sDateFin = rFin.str; P.DateFin = rFin.date;
        }
      }

      const missing = getMissingFields(P);
      const ok = confirm(buildRecap(P));

      if (ok) {
        if (missing.length) {
          alert("Impossible de lancer : complète les champs manquants.\n" + missing.join(", "));
          const keep = editOneField(P);
          if (!keep) return null;
          continue;
        }
        return P;
      }

      const keep = editOneField(P);
      if (!keep) return null;
    }
  }

  // ============================================================
  // UI EXPORT
  // ============================================================

  function getOrCreateExportBox() {
    let box = document.getElementById("EXEM_EXPORT_BOX");
    if (!box) {
      box = document.createElement("div");
      box.id = "EXEM_EXPORT_BOX";
      box.style.cssText =
        "position:fixed;right:12px;bottom:12px;z-index:999999;" +
        "background:#fff;border:1px solid #999;border-radius:6px;" +
        "padding:10px;max-width:55vw;max-height:45vh;overflow:auto;" +
        "font:12px/1.3 Arial;box-shadow:0 2px 10px rgba(0,0,0,0.2)";

      const closeBtn = document.createElement("button");
      closeBtn.textContent = "✕";
      closeBtn.style.cssText =
        "position:absolute;top:6px;right:8px;border:none;background:none;" +
        "cursor:pointer;font-weight:bold;font-size:14px";
      closeBtn.onclick = () => box.remove();
      box.appendChild(closeBtn);

      const title = document.createElement("div");
      title.textContent = "Exports";
      title.style.cssText = "font-weight:bold;margin-bottom:8px";
      box.appendChild(title);

      const btns = document.createElement("div");
      btns.id = "EXEM_EXPORT_BTNS";
      btns.style.cssText = "display:flex;flex-direction:column;gap:6px";
      box.appendChild(btns);

      const hint = document.createElement("div");
      hint.textContent = "Clique sur un bouton pour enregistrer le fichier.";
      hint.style.cssText = "margin-top:8px;color:#444";
      box.appendChild(hint);

      document.body.appendChild(box);
    }

    let btnWrap = box.querySelector("#EXEM_EXPORT_BTNS");
    if (!btnWrap) {
      btnWrap = document.createElement("div");
      btnWrap.id = "EXEM_EXPORT_BTNS";
      btnWrap.style.cssText = "display:flex;flex-direction:column;gap:6px";
      box.appendChild(btnWrap);
    }

    return box;
  }

  async function downloadFileUserClick(name, content, label = "Télécharger") {

    const box = getOrCreateExportBox();
    const btnWrap = box.querySelector("#EXEM_EXPORT_BTNS");

    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = `${label} : ${name}`;
    btn.style.cssText =
      "cursor:pointer;padding:6px 10px;text-align:left;border:1px solid #777;" +
      "border-radius:4px;background:#f7f7f7";

    btn.onclick = async () => {

      if (window.showSaveFilePicker) {
        try {
          const handle = await window.showSaveFilePicker({
            suggestedName: name,
            types: [{
              description: "CSV",
              accept: { "text/csv": [".csv"] }
            }]
          });

          const writable = await handle.createWritable();
          await writable.write(content);
          await writable.close();

          btn.disabled = true;
          btn.textContent = `Enregistré : ${name}`;
          btn.style.opacity = "0.7";
          return;

        } catch (err) {
          if (err && err.name === "AbortError") return;
          alert("Erreur sauvegarde : " + (err && err.message ? err.message : err));
          console.error(err);
          return;
        }
      }

      try {
        const blob = new Blob([content], { type: "text/csv;charset=utf-8" });
        const url = URL.createObjectURL(blob);

        const a = document.createElement("a");
        a.href = url;
        a.download = name;
        document.body.appendChild(a);
        a.click();
        a.remove();

        setTimeout(() => URL.revokeObjectURL(url), 1000);

        btn.disabled = true;
        btn.textContent = `Téléchargé : ${name}`;
        btn.style.opacity = "0.7";

      } catch (e) {
        alert("Erreur téléchargement : " + (e && e.message ? e.message : e));
        console.error(e);
      }
    };

    btnWrap.appendChild(btn);
  }

  // ============================================================
  // 1) Saisie en premier (référence/adresse + dates de contrôle)
  // ============================================================

  const P = collectAndConfirmUserInputs();
  if (!P) { alert("Abandon."); return; }

  const reference = P.reference;
  const adresse = P.adresse;

  const sDateDeb = P.sDateDeb;
  const sDateFin = P.sDateFin;

  const DateDeb = P.DateDeb;
  const DateFin = P.DateFin;

  const archiverPixels = P.archiverPixels;

  // ============================================================
  // 2) Extraction mesures via les POP-UPS (tooltips HTML)
  // ============================================================

  const sleep = (ms) => new Promise(r => setTimeout(r, ms));

  function getTooltipTextHTML() {
    const el = document.querySelector("div.highcharts-label, div.highcharts-tooltip, span.highcharts-tooltip");
    if (!el) return null;
    const t = (el.innerText || el.textContent || "").replace(/\s+/g, " ").trim();
    return t || null;
  }

  function parseTooltip(t) {
    // Exemples possibles :
    // "02/09/2022 21:06 3.55 V/m"
    // ou tooltip sur 2 lignes -> innerText remet un espace
    const s = String(t || "").replace(/\s+/g, " ").trim();
    const m = s.match(/(\d{2}\/\d{2}\/\d{4})\s+(\d{2}:\d{2})(?::(\d{2}))?.*?(-?\d+(?:[.,]\d+)?)/);
    if (!m) return null;

    const dt = m[1] + " " + m[2] + (m[3] ? (":" + m[3]) : "");
    const d = parseFRDate(dt);
    if (!d) return null;

    const E = parseFloat(m[4].replace(",", "."));
    if (!Number.isFinite(E)) return null;

    return [d.getTime(), E];
  }

  const ptEls = Array.from(document.querySelectorAll("g.highcharts-markers .highcharts-point"));
  if (!ptEls.length) { alert("Points (markers) introuvables"); return; }

  // Tri gauche->droite : on utilise le centre écran (plus robuste que getBBox seul)
  const pts = ptEls
    .map(el => {
      const r = el.getBoundingClientRect();
      return { el, cx: r.left + r.width / 2, cy: r.top + r.height / 2, x: r.left + r.width / 2 };
    })
    .sort((a, b) => a.x - b.x)
    .map(o => o.el);

  const decoded = [];
  const audit = [];
  const seen = new Set();

  for (let i = 0; i < pts.length; i++) {
    const el = pts[i];
    const r = el.getBoundingClientRect();
    const cx = r.left + r.width / 2;
    const cy = r.top + r.height / 2;

    el.dispatchEvent(new MouseEvent("mouseover", { bubbles: true, clientX: cx, clientY: cy }));
    el.dispatchEvent(new MouseEvent("mousemove", { bubbles: true, clientX: cx, clientY: cy }));

    await sleep(350);

    const txt = getTooltipTextHTML();
    if (!txt) {
      audit.push(`AUDIT;TOOLTIP_INTROUVABLE;${i}`);
      continue;
    }

    const rParsed = parseTooltip(txt);
    if (!rParsed) {
      audit.push(`AUDIT;TOOLTIP_NON_PARSE;${i};${txt}`);
      continue;
    }

    const key = rParsed[0] + "|" + rParsed[1].toFixed(6);
    if (!seen.has(key)) {
      seen.add(key);
      decoded.push(rParsed);
    }
  }

  if (decoded.length < 2) { alert("Pas assez de mesures extraites depuis les pop-ups."); return; }

  decoded.sort((a, b) => a[0] - b[0]);

  // ============================================================
  // 3) Contrôle cohérence dates saisies vs dates extraites
  // ============================================================

  const tFirst = decoded[0][0];
  const tLast  = decoded[decoded.length - 1][0];

  const diffFirstMin = Math.round((tFirst - DateDeb.getTime()) / 60000);
  const diffLastMin  = Math.round((tLast  - DateFin.getTime()) / 60000);

  audit.push(`AUDIT;DATE_FIRST_EXTRACTED;${fmtFRDate(new Date(tFirst))};DiffMin=${diffFirstMin}`);
  audit.push(`AUDIT;DATE_LAST_EXTRACTED;${fmtFRDate(new Date(tLast))};DiffMin=${diffLastMin}`);

  if (Math.abs(diffFirstMin) > 5 || Math.abs(diffLastMin) > 5) {
    alert(
      "Alerte cohérence dates\n" +
      "Date début saisie : " + fmtFRDate(DateDeb) + "\n" +
      "Date début extraite : " + fmtFRDate(new Date(tFirst)) + " (Delta " + diffFirstMin + " min)\n\n" +
      "Date fin saisie : " + fmtFRDate(DateFin) + "\n" +
      "Date fin extraite : " + fmtFRDate(new Date(tLast)) + " (Delta " + diffLastMin + " min)\n\n" +
      "Vérifie les dates saisies et celles affichées dans les pop-ups."
    );
  }

  // ============================================================
  // 4) Filtrages : delta (uniquement entre mesures valides) + seuil E
  // ============================================================

  let deltaIssues = 0;
  let expoIssues = 0;

  let prevValidIdx = null;

  for (let k = 0; k < decoded.length; k++) {

    // Filtre expo max
    if (decoded[k][1] !== null && decoded[k][1] >= SEUIL_EXPO_MAX) {
      expoIssues++;
      audit.push(`AUDIT;EXPO_SUP_10;${k};E=${decoded[k][1]}`);
      decoded[k][1] = null;
      continue;
    }

    // Filtre delta : comparer uniquement à la précédente mesure valide (E non vide)
    if (decoded[k][1] !== null) {
      if (prevValidIdx !== null) {
        const dtMin = (decoded[k][0] - decoded[prevValidIdx][0]) / 60000;
        if (dtMin <= SEUIL_DELTA_MINUTES) {
          deltaIssues++;
          audit.push(`AUDIT;DELTA_TROP_PETIT;${k};DeltaMin=${dtMin}`);
          decoded[k][1] = null;
          continue;
        }
      }
      prevValidIdx = k;
    }
  }

  const nbMesures = decoded.length;
  const nbMesuresValides = decoded.reduce((acc, d) => acc + (d[1] === null ? 0 : 1), 0);

  // ============================================================
  // 5) Stats a posteriori : Min / Moy / Max
  // ============================================================

  const vals = decoded.map(d => d[1]).filter(v => v !== null && Number.isFinite(v));
  let Emin = NaN, Emoy = NaN, Emax = NaN;

  if (vals.length) {
    Emin = Math.min(...vals);
    Emax = Math.max(...vals);
    Emoy = vals.reduce((a, b) => a + b, 0) / vals.length;
  }

  audit.push(`AUDIT;STATS;Emin=${Number.isFinite(Emin) ? fmtFRNumber(Emin) : ""};Emoy=${Number.isFinite(Emoy) ? fmtFRNumber(Emoy) : ""};Emax=${Number.isFinite(Emax) ? fmtFRNumber(Emax) : ""}`);
  audit.push(`AUDIT;FILTERS;NbDeltaTropPetit=${deltaIssues};NbExpoSup10=${expoIssues}`);

  alert(
    "Stats calculées (à comparer avec EXEM) :\n" +
    "Min : " + (Number.isFinite(Emin) ? fmtFRNumber(Emin) : "NA") + " V/m\n" +
    "Moy : " + (Number.isFinite(Emoy) ? fmtFRNumber(Emoy) : "NA") + " V/m\n" +
    "Max : " + (Number.isFinite(Emax) ? fmtFRNumber(Emax) : "NA") + " V/m"
  );

  // ============================================================
  // 6) Construction CSV
  // ============================================================

  const now = new Date();
  const refSafe = sanitizeFileName(reference);
  const baseName = `${refSafe || "Capteur"}__${fmtCompactLocal(DateDeb)}__${fmtCompactLocal(DateFin)}`;

  const lines = [];

  lines.push(`META;Format;EXPO_CAPTEUR_V1`);
  lines.push(`META;ScriptVersion;${SCRIPT_VERSION}`);
  lines.push(`META;DateCreationExport;${fmtFRDate(now)}`);
  lines.push(`META;Reference_Capteur;${reference}`);
  lines.push(`META;Adresse_Capteur;${adresse}`);
  lines.push(`META;DateDebut_Saisie;${sDateDeb}`);
  lines.push(`META;DateFin_Saisie;${sDateFin}`);
  lines.push(`META;DateDebut_Extraite;${fmtFRDate(new Date(tFirst))}`);
  lines.push(`META;DateFin_Extraite;${fmtFRDate(new Date(tLast))}`);
  lines.push(`META;DeltaDebut_Min;${diffFirstMin}`);
  lines.push(`META;DeltaFin_Min;${diffLastMin}`);

  lines.push(`META;ExpoMin_Decodee_Vm;${Number.isFinite(Emin) ? fmtFRNumber(Emin) : ""}`);
  lines.push(`META;ExpoMoy_Decodee_Vm;${Number.isFinite(Emoy) ? fmtFRNumber(Emoy) : ""}`);
  lines.push(`META;ExpoMax_Decodee_Vm;${Number.isFinite(Emax) ? fmtFRNumber(Emax) : ""}`);

  lines.push(`META;NbDeltaTropPetit;${deltaIssues}`);
  lines.push(`META;NbExpoSup10;${expoIssues}`);

  lines.push(`META;Pixels_Archive;${archiverPixels ? "OUI" : "NON"}`);
  lines.push(`META;NbMesures;${nbMesures}`);
  lines.push(`META;NbMesuresValides;${nbMesuresValides}`);
  lines.push(`META;SeuilExpoMax_Vm;${fmtFRNumber(SEUIL_EXPO_MAX)}`);
  lines.push(`META;SeuilDeltaMinutes;${SEUIL_DELTA_MINUTES}`);
  lines.push(`META;RegleFiltrage;Delta<=${SEUIL_DELTA_MINUTES}min_exclu_sur_mesures_valides;Expo>=${fmtFRNumber(SEUIL_EXPO_MAX)}Vm_exclu`);

  lines.push(`DATA;DateHeure;Exposition_Vm`);

  decoded.forEach(d => {
    lines.push(`DATA;${fmtFRDate(new Date(d[0]))};${d[1] === null ? "" : fmtFRNumber(d[1])}`);
  });

  audit.forEach(a => lines.push(a));

  // ============================================================
  // 7) Export
  // ============================================================

  try {
    downloadFileUserClick(baseName + ".csv", lines.join("\n"), "Télécharger CSV");
    alert("Export prêt. Les boutons sont en bas à droite : " + baseName);
  } catch (e) {
    alert("Erreur export UI : " + (e && e.message ? e.message : e));
    console.error(e);
  }

})();
