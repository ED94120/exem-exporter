(() => {

  const SCRIPT_VERSION = "EXPO_CAPTEUR_V1_2026_02_21";
  const SEUIL_EXPO_MAX = 10.0;

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

function askText(label, current) {
  const v = prompt(label, current == null ? "" : String(current));
  if (v === null) return null; // annulation
  return v;
}

function askDate(label, currentStr) {
  const v = prompt(label, currentStr || "");
  if (v === null) return null;
  const d = parseFRDate(v);
  if (!d) { alert("Date invalide. Format attendu : dd/mm/yyyy hh:mm"); return askDate(label, currentStr); }
  return { str: v, date: d };
}

function askNumberFR(label, currentNum) {
  const cur = (currentNum == null || !Number.isFinite(currentNum)) ? "" : String(currentNum).replace(".", ",");
  const v = prompt(label, cur);
  if (v === null) return null;
  const n = parseFloat(String(v).replace(",", "."));
  if (!Number.isFinite(n)) { alert("Nombre invalide."); return askNumberFR(label, currentNum); }
  return n;
}

function buildRecap(P) {
  const yesno = b => (b ? "OUI" : "NON");
  return [
    "Récapitulatif (valider = OK)",
    "--------------------------------",
    `Référence : ${P.reference || ""}`,
    `Adresse   : ${P.adresse || ""}`,
    `Date début: ${P.sDateDeb || ""}`,
    `Date fin  : ${P.sDateFin || ""}`,
    `Expo début: ${Number.isFinite(P.ExpoDeb) ? fmtFRNumber(P.ExpoDeb) : ""} V/m`,
    `Expo fin  : ${Number.isFinite(P.ExpoFin) ? fmtFRNumber(P.ExpoFin) : ""} V/m`,
    `ExpoMax ? : ${yesno(P.hasMax)}`,
    `Date max  : ${P.sDateMax || ""}`,
    `Expo max  : ${Number.isFinite(P.ExpoMax) ? fmtFRNumber(P.ExpoMax) : ""} V/m`,
    `Pixels bruts : ${yesno(P.archiverPixels)}`,
    "",
    "OK = Continuer / Annuler = Modifier"
  ].join("\n");
}

function validateExpoMaxRule(P) {
  // règle ExpoMax obligatoire
  if (P.ExpoDeb <= 0 || P.ExpoFin === P.ExpoDeb ||
      Math.abs(P.ExpoFin - P.ExpoDeb) / P.ExpoDeb < 0.20) {
    if (!Number.isFinite(P.ExpoMax)) {
      return false;
    }
  }
  return true;
}

function editOneField(P) {
  const choice = prompt(
    "Quel champ modifier ?\n" +
    "1 Référence\n" +
    "2 Adresse\n" +
    "3 Date début\n" +
    "4 Date fin\n" +
    "5 Expo début\n" +
    "6 Expo fin\n" +
    "7 Activer/Désactiver ExpoMax\n" +
    "8 Date max\n" +
    "9 Expo max\n" +
    "10 Pixels bruts\n" +
    "0 Annuler",
    "0"
  );
  if (choice === null) return false;

  const c = String(choice).trim();

  if (c === "0") return false;

  if (c === "1") {
    const v = askText("Référence Capteur :", P.reference);
    if (v === null) return false;
    P.reference = v;
    return true;
  }

  if (c === "2") {
    const v = askText("Adresse Capteur :", P.adresse);
    if (v === null) return false;
    P.adresse = v;
    return true;
  }

  if (c === "3") {
    const r = askDate("Date début (dd/mm/yyyy hh:mm) :", P.sDateDeb);
    if (r === null) return false;
    P.sDateDeb = r.str; P.DateDeb = r.date;
    return true;
  }

  if (c === "4") {
    const r = askDate("Date fin (dd/mm/yyyy hh:mm) :", P.sDateFin);
    if (r === null) return false;
    P.sDateFin = r.str; P.DateFin = r.date;
    return true;
  }

  if (c === "5") {
    const n = askNumberFR("Expo début (V/m) :", P.ExpoDeb);
    if (n === null) return false;
    P.ExpoDeb = n;
    return true;
  }

  if (c === "6") {
    const n = askNumberFR("Expo fin (V/m) :", P.ExpoFin);
    if (n === null) return false;
    P.ExpoFin = n;
    return true;
  }

  if (c === "7") {
    P.hasMax = confirm("Fournir ExpoMax ?");
    if (!P.hasMax) { P.sDateMax = ""; P.DateMax = null; P.ExpoMax = NaN; }
    return true;
  }

  if (c === "8") {
    if (!P.hasMax) { alert("ExpoMax n'est pas activé."); return true; }
    const r = askDate("Date MAX (dd/mm/yyyy hh:mm) :", P.sDateMax);
    if (r === null) return false;
    P.sDateMax = r.str; P.DateMax = r.date;
    return true;
  }

  if (c === "9") {
    if (!P.hasMax) { alert("ExpoMax n'est pas activé."); return true; }
    const n = askNumberFR("Expo MAX (V/m) :", P.ExpoMax);
    if (n === null) return false;
    P.ExpoMax = n;
    return true;
  }

  if (c === "10") {
    P.archiverPixels = confirm("Archiver pixels bruts ?");
    return true;
  }

  alert("Choix invalide.");
  return true;
}

function collectAndConfirmUserInputs() {
  const P = {
    reference: "",
    adresse: "",
    sDateDeb: "",
    sDateFin: "",
    DateDeb: null,
    DateFin: null,
    ExpoDeb: NaN,
    ExpoFin: NaN,
    hasMax: false,
    sDateMax: "",
    DateMax: null,
    ExpoMax: NaN,
    archiverPixels: false
  };

  // Saisie initiale
  {
    const v1 = askText("Référence Capteur (ex: Site #Nantes_46) :", "");
    if (v1 === null) return null;
    P.reference = v1;

    const v2 = askText("Adresse Capteur :", "");
    if (v2 === null) return null;
    P.adresse = v2;

    const rDeb = askDate("Date début (dd/mm/yyyy hh:mm) :", "");
    if (rDeb === null) return null;
    P.sDateDeb = rDeb.str; P.DateDeb = rDeb.date;

    const rFin = askDate("Date fin (dd/mm/yyyy hh:mm) :", "");
    if (rFin === null) return null;
    P.sDateFin = rFin.str; P.DateFin = rFin.date;

    const eDeb = askNumberFR("Expo début (V/m) :", NaN);
    if (eDeb === null) return null;
    P.ExpoDeb = eDeb;

    const eFin = askNumberFR("Expo fin (V/m) :", NaN);
    if (eFin === null) return null;
    P.ExpoFin = eFin;

    P.hasMax = confirm("Fournir ExpoMax ?");
    if (P.hasMax) {
      const rMax = askDate("Date MAX (dd/mm/yyyy hh:mm) :", "");
      if (rMax === null) return null;
      P.sDateMax = rMax.str; P.DateMax = rMax.date;

      const eMax = askNumberFR("Expo MAX (V/m) :", NaN);
      if (eMax === null) return null;
      P.ExpoMax = eMax;
    }

    P.archiverPixels = confirm("Archiver pixels bruts ?");
  }

  // Boucle récap + correction
  while (true) {

    if (!validateExpoMaxRule(P)) {
      alert("ExpoMax obligatoire selon règle (Expo début <=0, Expo fin = Expo début, ou variation < 20%).");
      P.hasMax = true;
      const rMax = askDate("Date MAX (dd/mm/yyyy hh:mm) :", P.sDateMax);
      if (rMax === null) return null;
      P.sDateMax = rMax.str; P.DateMax = rMax.date;

      const eMax = askNumberFR("Expo MAX (V/m) :", P.ExpoMax);
      if (eMax === null) return null;
      P.ExpoMax = eMax;
      continue;
    }

    const ok = confirm(buildRecap(P));
    if (ok) return P;

    const keepGoing = editOneField(P);
    if (!keepGoing) return null; // annulation
  }
}
  
  // ============================================================
  // UI EXPORT : crée une boîte + zone de boutons (corrige ton bug)
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
      
      box.style.position = "fixed"; // assure position absolue interne
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

    // garantit que la zone boutons existe
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
    const btnWrap = box.querySelector("#EXEM_EXPORT_BTNS"); // non-null garanti

    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = `${label} : ${name}`;
    btn.style.cssText =
      "cursor:pointer;padding:6px 10px;text-align:left;border:1px solid #777;" +
      "border-radius:4px;background:#f7f7f7";

    btn.onclick = async () => {

      // 1) mode moderne (Save File Picker)
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

      // 2) fallback universel : téléchargement par Blob (si showSaveFilePicker absent)
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
  // Infos utilisateur + récapitulatif + corrections
  // --------------------------
  
  const P = collectAndConfirmUserInputs();
  if (!P) { alert("Annulé par l'utilisateur."); return; }
  
  const reference = P.reference;
  const adresse = P.adresse;
  
  const sDateDeb = P.sDateDeb;
  const sDateFin = P.sDateFin;
  
  const DateDeb = P.DateDeb;
  const DateFin = P.DateFin;
  
  const ExpoDeb = P.ExpoDeb;
  const ExpoFin = P.ExpoFin;
  
  const DateMax = P.DateMax;
  const ExpoMax = P.ExpoMax;
  
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
  // Décodage + contrôle
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
  lines.push(`META;DateMax;${DateMax ? fmtFRDate(DateMax) : ""}`);
  lines.push(`META;ExpoMax_Vm;${Number.isFinite(ExpoMax) ? fmtFRNumber(ExpoMax) : ""}`);
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
  // Export (avec protection erreurs)
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
