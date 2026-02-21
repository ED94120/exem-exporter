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
    return `${p(d.getDate())}/${p(d.getMonth()+1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
  }

  function fmtFRNumber(n) {
    return Number(n).toFixed(2).replace(".", ",");
  }

  function sanitizeFileName(s) {
    return s.replace(/[\\\/:*?"<>|]/g, "").replace(/\s+/g, "_");
  }

  function downloadFile(name, content) {
    const blob = new Blob([content], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
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

  const reference = prompt("Référence Capteur (ex: Site #Nantes_46) :","");
  const adresse = prompt("Adresse Capteur :","");

  const sDateDeb = prompt("Date début (dd/mm/yyyy hh:mm) :","");
  const sDateFin = prompt("Date fin (dd/mm/yyyy hh:mm) :","");

  const DateDeb = parseFRDate(sDateDeb);
  const DateFin = parseFRDate(sDateFin);

  if (!DateDeb || !DateFin) { alert("Dates invalides"); return; }

  const ExpoDeb = parseFloat(prompt("Expo début (V/m) :","").replace(",","."));
  const ExpoFin = parseFloat(prompt("Expo fin (V/m) :","").replace(",","."));

  if (!Number.isFinite(ExpoDeb) || !Number.isFinite(ExpoFin)) {
    alert("Expositions invalides"); return;
  }

  let DateMax = null, ExpoMax = null;
  let hasMax = confirm("Fournir ExpoMax ?");
  if (hasMax) {
    const sDateMax = prompt("Date MAX (optionnel) :","");
    DateMax = parseFRDate(sDateMax);
    ExpoMax = parseFloat(prompt("Expo MAX :","").replace(",","."));
  }

  // règle ExpoMax obligatoire
  if (ExpoDeb <= 0 || ExpoFin === ExpoDeb ||
      Math.abs(ExpoFin-ExpoDeb)/ExpoDeb < 0.20) {
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
  const xFin = pts[pts.length-1][0];
  const yFin = pts[pts.length-1][1];

  const tDeb = DateDeb.getTime();
  const tFin = DateFin.getTime();

  const penteTemps = (tFin - tDeb)/(xFin - xDeb);
  let penteExpo = (ExpoFin - ExpoDeb)/(yFin - yDeb);

  // --------------------------
  // Décodage + contrôle
  // --------------------------

  const decoded = [];
  const audit = [];
  let inversions = 0;
  let prevTcode = -Infinity;

  for (let k=0;k<pts.length;k++) {

    let x = pts[k][0];
    let y = pts[k][1];

    if (x <= prevTcode) {
      inversions++;
      audit.push(`AUDIT;INVERSION_TEMPS_CODE;${k};x=${x}`);
    }
    prevTcode = x;

    const t_ms = tDeb + (x-xDeb)*penteTemps;
    const E = ExpoDeb + (y-yDeb)*penteExpo;

    decoded.push([t_ms, E]);
  }

  // --------------------------
  // Vérifications finales
  // --------------------------

  let deltaIssues = 0;

  for (let k=1;k<decoded.length;k++) {
    const dtMin = (decoded[k][0]-decoded[k-1][0])/60000;
    if (dtMin <= 30) {
      deltaIssues++;
      audit.push(`AUDIT;DELTA_TROP_PETIT;${k};DeltaMin=${dtMin}`);
      decoded[k][1] = null; // expo vide
    }
    if (decoded[k][1] !== null && decoded[k][1] >= SEUIL_EXPO_MAX) {
      audit.push(`AUDIT;EXPO_SUP_10;${k};E=${decoded[k][1]}`);
      decoded[k][1] = null;
    }
  }

  // --------------------------
  // Construction CSV
  // --------------------------

  const now = new Date();
  const baseName =
    sanitizeFileName(reference) + "__" +
    DateDeb.toISOString().slice(0,16).replace(/[-:T]/g,"") +
    "__" +
    DateFin.toISOString().slice(0,16).replace(/[-:T]/g,"");

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
  lines.push(`META;DateMax;${DateMax?fmtFRDate(DateMax):""}`);
  lines.push(`META;ExpoMax_Vm;${ExpoMax?fmtFRNumber(ExpoMax):""}`);
  lines.push(`META;InversionDetectee;${inversions>0?"OUI":"NON"}`);
  lines.push(`META;NbCouplesInverses;${inversions}`);
  lines.push(`META;NbDeltaTropPetit;${deltaIssues}`);
  lines.push(`META;Pixels_Archive;${archiverPixels?"OUI":"NON"}`);

  lines.push(`DATA;DateHeure;Exposition_Vm`);

  decoded.forEach(d => {
    lines.push(`DATA;${fmtFRDate(new Date(d[0]))};${d[1]===null?"":fmtFRNumber(d[1])}`);
  });

  audit.forEach(a => lines.push(a));

  downloadFile(baseName + ".csv", lines.join("\n"));

  if (archiverPixels) {
    const pix = ["PIXELS;xp;yp"];
    pts.forEach(p=>pix.push(`PIXELS;${p[0]};${p[1]}`));
    downloadFile(baseName + "_pixels.csv", pix.join("\n"));
  }

  alert("Export terminé : " + baseName + ".csv");

})();
