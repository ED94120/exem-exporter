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
  // Infos utilisateur
  // --------------------------

  const reference = prompt("Référence Capteur (ex: Site #Nantes_46) :", "");
  const adresse = prompt("Adresse Capteur :", "");

  const sDateDeb = prompt("Date début (dd/mm/yyyy hh:mm) :", "");
  const sDateFin = prompt("Date fin (dd/mm/yyyy hh:mm) :", "");

  const DateDeb = parseFRDate(sDateDeb);
  const DateFin = parseFRDate(sDateFin);

  if (!DateDeb || !DateFin) { alert("Dates invalides"); return; }

  const ExpoDeb = parseFloat(String(prompt("Expo début (V/m) :", "") || "").replace(",", "."));
  const ExpoFin = parseFloat(String(prompt("Expo fin (V/m) :", "") || "").replace(",", "."));

  if (!Number.isFinite(ExpoDeb) || !Number.isFinite(ExpoFin)) {
    alert("Expositions invalides"); return;
  }

  let DateMax = null, ExpoMax = null;
  const hasMax = confirm("Fournir ExpoMax ?");
  if (hasMax) {
    const sDateMax = prompt("Date MAX (dd/mm/yyyy hh:mm) :", "");
    DateMax = parseFRDate(sDateMax);
    ExpoMax = parseFloat(String(prompt("Expo MAX (V/m) :", "") || "").replace(",", "."));
  }

  // règle ExpoMax obligatoire
  if (ExpoDeb <= 0 || ExpoFin === ExpoDeb ||
    Math.abs(ExpoFin - ExpoDeb) / ExpoDeb < 0.20) {
    if (!Number.isFinite(ExpoMax)) {
      alert("ExpoMax obligatoire selon règle.");
      return;
    }
  }

  const archiverPixels = confirm("Archiver pixels bruts ?");

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
