(() => {

  const SCRIPT_VERSION = "EXPO_CAPTEUR_V1_2026_02_22";
  const SEUIL_EXPO_MAX = 10.0;

  // --------------------------
  // Utils dates / nombres
  // --------------------------

  function parseFRDate(s) {
    const m = String(s || "").trim().match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})$/);
    if (!m) return null;
    return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], 0, 0);
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
  // Saisie : Cancel = "je ne sais pas encore"
  // --------------------------

  function askText(label, current) {
    const v = prompt(label, current == null ? "" : String(current));
    if (v === null) return null;          // cancel = inconnu
    const s = String(v);
    return s;                              // vide autorisé (bloqué à la synthèse si obligatoire)
  }

  function askDate(label, currentStr) {
    const v = prompt(label, currentStr || "");
    if (v === null) return { str: null, date: null };   // cancel = inconnu
    const s = String(v).trim();
    if (s === "") return { str: "", date: null };       // vide = inconnu
    const d = parseFRDate(s);
    if (!d) {
      alert("Date invalide. Format attendu : dd/mm/yyyy hh:mm");
      return askDate(label, currentStr);
    }
    return { str: s, date: d };
  }

  function askNumberFR(label, currentNum) {
    const cur = (currentNum == null || !Number.isFinite(currentNum)) ? "" : String(currentNum).replace(".", ",");
    const v = prompt(label, cur);
    if (v === null) return NaN;            // cancel = inconnu
    const s = String(v).trim();
    if (s === "") return NaN;              // vide = inconnu
    const n = parseFloat(s.replace(",", "."));
    if (!Number.isFinite(n)) {
      alert("Nombre invalide.");
      return askNumberFR(label, currentNum);
    }
    return n;
  }

  function isMissingText(v) {
    return (v == null) || (String(v).trim() === "");
  }
  function isMissingDate(d) {
    return !(d instanceof Date) || isNaN(d.getTime());
  }
  function isMissingNumber(n) {
    return !Number.isFinite(n);
  }

  function getMissingFields(P) {
    const miss = [];
    if (isMissingText(P.reference)) miss.push("Référence");
    if (isMissingText(P.adresse)) miss.push("Adresse");
    if (isMissingDate(P.DateDeb)) miss.push("Date début");
    if (isMissingDate(P.DateFin)) miss.push("Date fin");
    if (isMissingNumber(P.ExpoDeb)) miss.push("Expo début");
    if (isMissingNumber(P.ExpoFin)) miss.push("Expo fin");
    if (isMissingNumber(P.ExpoMax)) miss.push("Expo max");
    return miss;
  }

  function buildRecap(P) {
    const show = v => (isMissingText(v) ? "NON RENSEIGNÉ" : String(v));
    const showDate = (s, d) => (isMissingDate(d) ? "NON RENSEIGNÉ" : s);
    const showNum = n => (isMissingNumber(n) ? "NON RENSEIGNÉ" : (fmtFRNumber(n) + " V/m"));
    const yesno = b => (b ? "OUI" : "NON");

    const missing = getMissingFields(P);

    return [
      "Récapitulatif",
      "--------------------------------",
      `Référence : ${show(P.reference)}`,
      `Adresse   : ${show(P.adresse)}`,
      `Date début: ${showDate(P.sDateDeb, P.DateDeb)}`,
      `Date fin  : ${showDate(P.sDateFin, P.DateFin)}`,
      `Expo début: ${showNum(P.ExpoDeb)}`,
      `Expo fin  : ${showNum(P.ExpoFin)}`,
      `Expo max  : ${showNum(P.ExpoMax)}`,
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
      "5 Expo début\n" +
      "6 Expo fin\n" +
      "7 Expo max\n" +
      "8 Pixels bruts\n" +
      "0 Retour",
      "0"
    );
    if (choice === null) return true; // cancel -> retour synthèse

    const c = String(choice).trim();
    if (c === "0") return true;

    if (c === "1") { P.reference = askText("Référence Capteur :", P.reference); return true; }
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

    if (c === "5") { P.ExpoDeb = askNumberFR("Expo début (V/m) :", P.ExpoDeb); return true; }
    if (c === "6") { P.ExpoFin = askNumberFR("Expo fin (V/m) :", P.ExpoFin); return true; }
    if (c === "7") { P.ExpoMax = askNumberFR("Expo max (V/m) :", P.ExpoMax); return true; }

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
      ExpoDeb: NaN,
      ExpoFin: NaN,
      ExpoMax: NaN,
      archiverPixels: false
    };

    // Saisie initiale : Cancel => valeur manquante, on continue
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

    P.ExpoDeb = askNumberFR("Expo début (V/m) :", NaN);
    P.ExpoFin = askNumberFR("Expo fin (V/m) :", NaN);
    P.ExpoMax = askNumberFR("Expo max (V/m) :", NaN);

    P.archiverPixels = confirm("Archiver pixels bruts ?");

    // Boucle synthèse / correction : OK ne marche que si complet
    while (true) {

      // Contrôle DateFin > DateDeb si les 2 existent
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
          editOneField(P);
          continue;
        }
        return P;
      }

      editOneField(P);
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

      // Fallback universel : Blob download
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

  // --------------------------
  // Extraction pixels
  // --------------------------

  const graph = document.querySelector("path.highcharts-graph");
  if (!graph) { alert("Graph not found"); return; }

  const d = graph.getAttribute("d") || "";
  const tokens = d.match(/[A-Za-z]|-?\d*\.?\d+(?:e[+-]?\d+)?/g) || [];

  const pts = [];
  let i = 0;

  while (i < tokens.length) {
    const t = tokens[i++];
    if (t === "M" || t === "L") {
      const x = parseFloat(tokens[i++]);
      const y = parseFloat(tokens[i++]);
      if (Number.isFinite(x) && Number.isFinite(y)) pts.push([x, y]);
    } else if (t === "C") {
      i += 4;
      const x = parseFloat(tokens[i++]);
      const y = parseFloat(tokens[i++]);
      if (Number.isFinite(x) && Number.isFinite(y)) pts.push([x, y]);
    }
  }

  if (pts.length < 2) { alert("Pas assez de points."); return; }

  pts.sort((a, b) => a[0] - b[0]);

  // --------------------------
  // Saisie + synthèse/corrections
  // --------------------------

  const P = collectAndConfirmUserInputs();

  const reference = P.reference;
  const adresse = P.adresse;

  const sDateDeb = P.sDateDeb;
  const sDateFin = P.sDateFin;

  const DateDeb = P.DateDeb;
  const DateFin = P.DateFin;

  const ExpoDeb = P.ExpoDeb;
  const ExpoFin = P.ExpoFin;
  const ExpoMaxUser = P.ExpoMax;

  const archiverPixels = P.archiverPixels;

  // --------------------------
  // Calibration
  // --------------------------

  const xDeb = pts[0][0];
  const yDeb = pts[0][1];
  const xFin = pts[pts.length - 1][0];
  const yFin = pts[pts.length - 1][1];

  const tDeb = DateDeb.getTime();
  const tFin = DateFin.getTime();

  const penteTemps = (tFin - tDeb) / (xFin - xDeb);
  const penteExpo = (ExpoFin - ExpoDeb) / (yFin - yDeb);

  // --------------------------
  // Décodage
  // --------------------------

  const decoded = [];
  const audit = [];
  let inversions = 0;
  let prevTcode = -Infinity;

  for (let k = 0; k < pts.length; k++) {

    const x = pts[k][0];
    const y = pts[k][1];

    if (x <= prevTcode) {
      inversions++;
      audit.push(`AUDIT;INVERSION_TEMPS_CODE;${k};x=${x}`);
    }
    prevTcode = x;

    const t_ms = tDeb + (x - xDeb) * penteTemps;
    const E = ExpoDeb + (y - yDeb) * penteExpo;

    decoded.push([t_ms, E]);
  }

  // --------------------------
  // Vérifications finales
  // --------------------------

  let deltaIssues = 0;

  for (let k = 1; k < decoded.length; k++) {
    const dtMin = (decoded[k][0] - decoded[k - 1][0]) / 60000;

    if (dtMin <= 30) {
      deltaIssues++;
      audit.push(`AUDIT;DELTA_TROP_PETIT;${k};DeltaMin=${dtMin}`);
      decoded[k][1] = null; // expo vide
    }

    if (decoded[k][1] !== null && decoded[k][1] >= SEUIL_EXPO_MAX) {
      audit.push(`AUDIT;EXPO_SUP_10;${k};E=${decoded[k][1]}`);
      decoded[k][1] = null; // expo vide
    }
  }

  const nbMesures = decoded.length;
  const nbMesuresValides = decoded.reduce((acc, d) => acc + (d[1] === null ? 0 : 1), 0);

  // --------------------------
  // Stats a posteriori : Min / Moy (moyenne simple EXEM) / Max
  // --------------------------

  const vals = decoded.map(d => d[1]).filter(v => v !== null && Number.isFinite(v));
  let Emin = NaN, Emoy = NaN, Emax = NaN;

  if (vals.length) {
    Emin = Math.min(...vals);
    Emax = Math.max(...vals);
    Emoy = vals.reduce((a, b) => a + b, 0) / vals.length;
  }

  // Contrôle visuel : comparaison max saisi vs max décodé (informatif)
  if (Number.isFinite(Emax) && Number.isFinite(ExpoMaxUser)) {
    const diff = Math.abs(Emax - ExpoMaxUser);
    if (diff > 0.15 && diff > 0.05 * ExpoMaxUser) {
      audit.push(`AUDIT;MAX_INCOHERENT;EmaxDecoded=${fmtFRNumber(Emax)};EmaxSaisi=${fmtFRNumber(ExpoMaxUser)};Diff=${fmtFRNumber(diff)}`);
    }
  }

  audit.push(`AUDIT;STATS;Emin=${Number.isFinite(Emin) ? fmtFRNumber(Emin) : ""};Emoy=${Number.isFinite(Emoy) ? fmtFRNumber(Emoy) : ""};Emax=${Number.isFinite(Emax) ? fmtFRNumber(Emax) : ""}`);

  alert(
    "Stats calculées (à comparer avec EXEM) :\n" +
    "Min : " + (Number.isFinite(Emin) ? fmtFRNumber(Emin) : "NA") + " V/m\n" +
    "Moy : " + (Number.isFinite(Emoy) ? fmtFRNumber(Emoy) : "NA") + " V/m\n" +
    "Max : " + (Number.isFinite(Emax) ? fmtFRNumber(Emax) : "NA") + " V/m"
  );

  // --------------------------
  // Construction CSV
  // --------------------------

  const now = new Date();
  const refSafe = sanitizeFileName(reference);
  const baseName = `${refSafe || "Capteur"}__${fmtCompactLocal(DateDeb)}__${fmtCompactLocal(DateFin)}`;

  const lines = [];

  lines.push(`META;Format;EXPO_CAPTEUR_V1`);
  lines.push(`META;ScriptVersion;${SCRIPT_VERSION}`);
  lines.push(`META;DateCreationExport;${fmtFRDate(now)}`);
  lines.push(`META;Reference_Capteur;${reference}`);
  lines.push(`META;Adresse_Capteur;${adresse}`);
  lines.push(`META;DateDebut;${sDateDeb}`);
  lines.push(`META;DateFin;${sDateFin}`);
  lines.push(`META;ExpoDebut_Vm;${fmtFRNumber(ExpoDeb)}`);
  lines.push(`META;ExpoFin_Vm;${fmtFRNumber(ExpoFin)}`);
  lines.push(`META;ExpoMax_Saisie_Vm;${Number.isFinite(ExpoMaxUser) ? fmtFRNumber(ExpoMaxUser) : ""}`);
  lines.push(`META;ExpoMin_Decodee_Vm;${Number.isFinite(Emin) ? fmtFRNumber(Emin) : ""}`);
  lines.push(`META;ExpoMoy_Decodee_Vm;${Number.isFinite(Emoy) ? fmtFRNumber(Emoy) : ""}`);
  lines.push(`META;ExpoMax_Decodee_Vm;${Number.isFinite(Emax) ? fmtFRNumber(Emax) : ""}`);
  lines.push(`META;InversionDetectee;${inversions > 0 ? "OUI" : "NON"}`);
  lines.push(`META;NbCouplesInverses;${inversions}`);
  lines.push(`META;NbDeltaTropPetit;${deltaIssues}`);
  lines.push(`META;Pixels_Archive;${archiverPixels ? "OUI" : "NON"}`);
  lines.push(`META;NbMesures;${nbMesures}`);
  lines.push(`META;NbMesuresValides;${nbMesuresValides}`);

  lines.push(`DATA;DateHeure;Exposition_Vm`);

  decoded.forEach(d => {
    lines.push(`DATA;${fmtFRDate(new Date(d[0]))};${d[1] === null ? "" : fmtFRNumber(d[1])}`);
  });

  audit.forEach(a => lines.push(a));

  // --------------------------
  // Export
  // --------------------------

  try {
    downloadFileUserClick(baseName + ".csv", lines.join("\n"), "Télécharger CSV");

    if (archiverPixels) {
      const pix = ["PIXELS;xp;yp"];
      pts.forEach(p => pix.push(`PIXELS;${p[0]};${p[1]}`));
      downloadFileUserClick(baseName + "_pixels.csv", pix.join("\n"), "Télécharger PIXELS");
    }

    alert("Export prêt. Les boutons sont en bas à droite : " + baseName);

  } catch (e) {
    alert("Erreur export UI : " + (e && e.message ? e.message : e));
    console.error(e);
  }

})();
