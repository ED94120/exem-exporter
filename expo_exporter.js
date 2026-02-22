(async () => {

  const SCRIPT_VERSION = "EXPO_CAPTEUR_V2_HIGHCHARTS_DATA_2026_02_22";
  const SEUIL_EXPO_MAX = 10.0;
  const SEUIL_DELTA_MINUTES = 30;

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
  // Saisie minimale (uniquement contrôle)
  // --------------------------

  function askText(label, current) {
    const v = prompt(label, current == null ? "" : String(current));
    if (v === null) return null;
    return String(v);
  }

  function askDate(label, currentStr) {
    const v = prompt(label, currentStr || "");
    if (v === null) return { str: null, date: null };
    const s = String(v).trim();
    if (s === "") return { str: "", date: null };
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

  function collectInputs() {
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
      const rDeb = askDate("Date début (dd/mm/yyyy hh:mm) (contrôle) :", "");
      P.sDateDeb = rDeb.str; P.DateDeb = rDeb.date;
    }
    {
      const rFin = askDate("Date fin (dd/mm/yyyy hh:mm) (contrôle) :", "");
      P.sDateFin = rFin.str; P.DateFin = rFin.date;
    }

    P.archiverPixels = confirm("Archiver pixels bruts (inutile ici) ?");

    if (isMissingText(P.reference) || isMissingText(P.adresse) || isMissingDate(P.DateDeb) || isMissingDate(P.DateFin)) {
      alert("Champs obligatoires manquants (référence, adresse, date début, date fin).");
      return null;
    }
    if (P.DateFin.getTime() <= P.DateDeb.getTime()) {
      alert("Erreur : Date fin doit être strictement après Date début.");
      return null;
    }
    return P;
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
            types: [{ description: "CSV", accept: { "text/csv": [".csv"] } }]
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
  // 1) Saisie
  // ============================================================

  const P = collectInputs();
  if (!P) { alert("Abandon."); return; }

  const reference = P.reference;
  const adresse = P.adresse;

  const sDateDeb = P.sDateDeb;
  const sDateFin = P.sDateFin;

  const DateDeb = P.DateDeb;
  const DateFin = P.DateFin;

  // ============================================================
  // 2) Récupération du chart Highcharts (sans markers / sans popups)
  // ============================================================

  if (!window.Highcharts || !Array.isArray(window.Highcharts.charts)) {
    alert("Highcharts introuvable (window.Highcharts.charts absent).");
    return;
  }

  const chart = window.Highcharts.charts.find(c => c && c.series && c.series.length);
  if (!chart) {
    alert("Aucun chart Highcharts trouvé.");
    return;
  }

  const series = chart.series.find(s => s && (s.xData || (s.options && s.options.data)));
  if (!series) {
    alert("Aucune série exploitable trouvée dans le chart.");
    return;
  }

  // xData/yData = méthode la plus fiable quand boost/simplification est active
  let xData = series.xData;
  let yData = series.yData;

  // fallback: options.data
  if ((!xData || !yData || xData.length !== yData.length) && series.options && Array.isArray(series.options.data)) {
    const raw = series.options.data;
    // raw peut être [y] (si catégories) ou [ [x,y], ... ]
    if (Array.isArray(raw[0])) {
      xData = raw.map(p => p[0]);
      yData = raw.map(p => p[1]);
    } else {
      // pas d'X explicite -> on utilisera les catégories ou l'index
      yData = raw.slice();
      xData = raw.map((_, i) => i);
    }
  }

  if (!xData || !yData || xData.length < 2 || xData.length !== yData.length) {
    alert("Impossible de récupérer xData/yData de façon cohérente.");
    return;
  }

  // Détection si X est un timestamp (ms) ou juste un index
  const looksLikeMs = (v) => Number.isFinite(v) && v > 1e11; // ~2003 en ms
  const xIsDateMs = looksLikeMs(xData[0]) || looksLikeMs(xData[xData.length - 1]);

  // Si X est un index : on tente de récupérer les catégories (labels) du chart
  let categories = null;
  if (!xIsDateMs && chart.xAxis && chart.xAxis[0] && Array.isArray(chart.xAxis[0].categories)) {
    categories = chart.xAxis[0].categories;
  }

  // ============================================================
  // 3) Construction decoded = [t_ms, E] (ou E sans date si impossible)
  // ============================================================

  const decoded = [];
  const audit = [];

  for (let i = 0; i < xData.length; i++) {
    const x = xData[i];
    const y = yData[i];

    if (!Number.isFinite(y)) continue; // ignore NaN/undefined

    let t_ms = null;

    if (xIsDateMs && Number.isFinite(x)) {
      t_ms = x;
    } else if (categories && categories[i]) {
      // categories souvent sous forme "dd/mm/yyyy hh:mm" ou similaire
      const d = parseFRDate(String(categories[i]).trim());
      if (d) t_ms = d.getTime();
    }

    decoded.push([t_ms, y]);
  }

  if (decoded.length < 2) {
    alert("Pas assez de valeurs exploitables dans la série.");
    return;
  }

  // Si on a des temps, on trie par temps
  const hasTimes = decoded.some(r => r[0] !== null);
  if (hasTimes) decoded.sort((a, b) => (a[0] ?? 0) - (b[0] ?? 0));

  // ============================================================
  // 4) Contrôle cohérence (si temps dispo)
  // ============================================================

  if (hasTimes) {
    const tFirst = decoded.find(r => r[0] !== null)?.[0];
    const tLast = [...decoded].reverse().find(r => r[0] !== null)?.[0];

    if (tFirst != null && tLast != null) {
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
          "=> Si ce n’est pas cohérent, c’est que le graphe n’expose pas les vraies dates en xData (cas ‘index/catégories’)."
        );
      }
    }
  } else {
    audit.push("AUDIT;WARNING;Aucune date récupérable via Highcharts (xData non timestamp et categories absentes).");
    alert("Attention : je récupère bien les expositions, mais je n’arrive pas à récupérer les dates (xData non temporel).");
  }

  // ============================================================
  // 5) Filtrage delta + expo (delta uniquement entre mesures valides)
  // ============================================================

  let deltaIssues = 0;
  let expoIssues = 0;

  // Filtre expo max
  for (let k = 0; k < decoded.length; k++) {
    if (decoded[k][1] !== null && decoded[k][1] >= SEUIL_EXPO_MAX) {
      expoIssues++;
      audit.push(`AUDIT;EXPO_SUP_10;${k};E=${decoded[k][1]}`);
      decoded[k][1] = null;
    }
  }

  // Filtre delta (uniquement si on a des temps)
  if (hasTimes) {
    let prevValidIdx = null;
    for (let k = 0; k < decoded.length; k++) {
      const t = decoded[k][0];
      const E = decoded[k][1];
      if (t == null || E === null) continue;

      if (prevValidIdx !== null) {
        const dtMin = (t - decoded[prevValidIdx][0]) / 60000;
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
  // 6) Stats
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
    "Stats calculées (série Highcharts) :\n" +
    "Min : " + (Number.isFinite(Emin) ? fmtFRNumber(Emin) : "NA") + " V/m\n" +
    "Moy : " + (Number.isFinite(Emoy) ? fmtFRNumber(Emoy) : "NA") + " V/m\n" +
    "Max : " + (Number.isFinite(Emax) ? fmtFRNumber(Emax) : "NA") + " V/m"
  );

  // ============================================================
  // 7) CSV
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
  lines.push(`META;HasTimes;${hasTimes ? "OUI" : "NON"}`);
  lines.push(`META;NbDeltaTropPetit;${deltaIssues}`);
  lines.push(`META;NbExpoSup10;${expoIssues}`);
  lines.push(`META;NbMesures;${nbMesures}`);
  lines.push(`META;NbMesuresValides;${nbMesuresValides}`);
  lines.push(`META;SeuilExpoMax_Vm;${fmtFRNumber(SEUIL_EXPO_MAX)}`);
  lines.push(`META;SeuilDeltaMinutes;${SEUIL_DELTA_MINUTES}`);
  lines.push(`META;ExpoMin_Vm;${Number.isFinite(Emin) ? fmtFRNumber(Emin) : ""}`);
  lines.push(`META;ExpoMoy_Vm;${Number.isFinite(Emoy) ? fmtFRNumber(Emoy) : ""}`);
  lines.push(`META;ExpoMax_Vm;${Number.isFinite(Emax) ? fmtFRNumber(Emax) : ""}`);

  lines.push(`DATA;DateHeure;Exposition_Vm`);

  decoded.forEach(d => {
    const dt = (d[0] == null) ? "" : fmtFRDate(new Date(d[0]));
    const e = (d[1] === null) ? "" : fmtFRNumber(d[1]);
    lines.push(`DATA;${dt};${e}`);
  });

  audit.forEach(a => lines.push(a));

  // ============================================================
  // 8) Export
  // ============================================================

  try {
    await downloadFileUserClick(baseName + ".csv", lines.join("\n"), "Télécharger CSV");
    alert("Export prêt. Les boutons sont en bas à droite : " + baseName);
  } catch (e) {
    alert("Erreur export UI : " + (e && e.message ? e.message : e));
    console.error(e);
  }

})();
