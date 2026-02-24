(async () => {

  const SCRIPT_VERSION = "EXPO_CAPTEUR_POPUP_V2_2026_02_23";
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

  const sleep = (ms) => new Promise(r => setTimeout(r, ms));

  // --------------------------
  // UI EXPORT
  // --------------------------

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

  // --------------------------
  // 0) Vérifier présence des points (markers) visibles dans la zone de tracé
  // --------------------------

  function getVisiblePointElements() {
    // On se limite au SVG Highcharts pour éviter de capturer d'autres SVG de la page
    const svg = document.querySelector("svg.highcharts-root");
    if (!svg) return [];

    // Zone de tracé (plot area)
    const plotBg = svg.querySelector("rect.highcharts-plot-background");
    const plotRect = plotBg ? plotBg.getBoundingClientRect() : null;

    // Points Highcharts (circle ou path) : class = highcharts-point
    const candidates = Array.from(svg.querySelectorAll("g.highcharts-markers .highcharts-point"));

    return candidates.filter(el => {
      const r = el.getBoundingClientRect();
      if (r.width <= 0 || r.height <= 0) return false;

      const cs = getComputedStyle(el);
      if (cs.display === "none" || cs.visibility === "hidden") return false;
      if (parseFloat(cs.opacity || "1") === 0) return false;

      // Vérifie que le point est dans le rectangle du graphe (évite les points hors zone)
      if (plotRect) {
        const cx = r.left + r.width / 2;
        const cy = r.top + r.height / 2;
        if (cx < plotRect.left || cx > plotRect.right) return false;
        if (cy < plotRect.top || cy > plotRect.bottom) return false;
      }
      return true;
    });
  }

  const ptElsCheck = getVisiblePointElements();

  if (ptElsCheck.length < 2) {
    alert(
      "Impossible d'extraire les mesures par pop-up : les points de mesure ne sont pas affichés.\n\n" +
      "Réduis la période (en général ≤ 7 jours) jusqu'à voir les points sur la courbe, puis relance."
    );
    return;
  }
  
 
  // --------------------------
  // 1) Saisie minimale (référence + adresse)
  // --------------------------

  const reference = prompt("Référence Capteur (ex: Site Nantes_01) :", "") || "";
  const adresse = prompt("Adresse Capteur :", "") || "";

  if (!String(reference).trim() || !String(adresse).trim()) {
    alert("Abandon : Référence et Adresse sont obligatoires.");
    return;
  }

  const archiverPixels = confirm("Archiver pixels bruts (positions écran des points) ?");

  // --------------------------
  // 2) Lecture tooltip (SVG Highcharts)
  // --------------------------

  function getTooltipTextAny() {
    // 1) Tooltip SVG Highcharts
    {
      const tip = document.querySelector("g.highcharts-tooltip");
      if (tip) {
        const tspans = Array.from(tip.querySelectorAll("text tspan"));
        const txt = tspans.map(t => (t.textContent || "").trim()).filter(Boolean).join(" ");
        const s = txt.replace(/\s+/g, " ").trim();
        if (s) return s;
      }
    }

    // 2) Fallback : tooltip HTML (selon config Highcharts)
    {
      const el = document.querySelector("div.highcharts-tooltip, div.highcharts-label, span.highcharts-tooltip");
      if (el) {
        const s = String(el.innerText || el.textContent || "").replace(/\s+/g, " ").trim();
        if (s) return s;
      }
    }

    return null;
  }

  function parseTooltip(t) {
    // On cherche : date + heure + nombre (V/m)
    // Exemple : "19/02/2026 00:06 2.78 V/m" (ou virgule)
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

  // Tri gauche->droite à partir des coordonnées écran
  const pts = getVisiblePointElements()
    .map(el => {
      const r = el.getBoundingClientRect();
      return { el, x: r.left + r.width / 2, y: r.top + r.height / 2 };
    })
    .sort((a, b) => a.x - b.x);

  const decoded = [];
  const audit = [];
  const seen = new Set();
  const pix = []; // optionnel

  // Important : sur certains graphiques, survoler le circle ne suffit pas,
  // on force aussi un mousemove sur document.
  for (let i = 0; i < pts.length; i++) {

    const cx = pts[i].x;
    const cy = pts[i].y;

    if (archiverPixels) {
      pix.push([cx, cy]);
    }

    try {
      pts[i].el.dispatchEvent(new MouseEvent("mouseover", { bubbles: true, clientX: cx, clientY: cy }));
      pts[i].el.dispatchEvent(new MouseEvent("mousemove", { bubbles: true, clientX: cx, clientY: cy }));
      document.dispatchEvent(new MouseEvent("mousemove", { bubbles: true, clientX: cx, clientY: cy }));
    } catch (e) {
      audit.push(`AUDIT;MOUSEEVENT_ERROR;${i};${String(e && e.message ? e.message : e)}`);
      continue;
    }

    // Le tooltip peut mettre un peu de temps à se mettre à jour (animation/latence)
    let txt = null;
    for (let attempt = 0; attempt < 3; attempt++) {
      await sleep(180);
      txt = getTooltipTextAny();
      if (txt) break;
    }

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

  if (decoded.length < 2) {
    alert(
      "Pas assez de mesures extraites depuis les pop-ups.\n\n" +
      "Si la période est trop longue, les points peuvent être désactivés : réduis la période (≤ 7 jours en général)."
    );
    return;
  }

  decoded.sort((a, b) => a[0] - b[0]);

  // Dates extraites (pour nom fichier + méta)
  const tFirst = decoded[0][0];
  const tLast = decoded[decoded.length - 1][0];

  audit.push(`AUDIT;DATE_FIRST_EXTRACTED;${fmtFRDate(new Date(tFirst))}`);
  audit.push(`AUDIT;DATE_LAST_EXTRACTED;${fmtFRDate(new Date(tLast))}`);

  // --------------------------
  // 3) Filtrages : seuil E + delta entre mesures valides
  // --------------------------

  let deltaIssues = 0;
  let expoIssues = 0;
  let prevValidIdx = null;

  for (let k = 0; k < decoded.length; k++) {

    if (decoded[k][1] !== null && decoded[k][1] >= SEUIL_EXPO_MAX) {
      expoIssues++;
      audit.push(`AUDIT;EXPO_SUP_10;${k};E=${decoded[k][1]}`);
      decoded[k][1] = null;
      continue;
    }

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

  // --------------------------
  // 4) Stats
  // --------------------------

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
    "Mesures extraites : " + nbMesures + " (valides : " + nbMesuresValides + ")\n\n" +
    "Stats (à comparer EXEM) :\n" +
    "Min : " + (Number.isFinite(Emin) ? fmtFRNumber(Emin) : "NA") + " V/m\n" +
    "Moy : " + (Number.isFinite(Emoy) ? fmtFRNumber(Emoy) : "NA") + " V/m\n" +
    "Max : " + (Number.isFinite(Emax) ? fmtFRNumber(Emax) : "NA") + " V/m"
  );

  // --------------------------
  // 5) CSV
  // --------------------------

  const now = new Date();
  const refSafe = sanitizeFileName(reference);
  const baseName = `${refSafe || "Capteur"}__${fmtCompactLocal(new Date(tFirst))}__${fmtCompactLocal(new Date(tLast))}`;

  const lines = [];
  lines.push(`META;Format;EXPO_CAPTEUR_V1`);
  lines.push(`META;ScriptVersion;${SCRIPT_VERSION}`);
  lines.push(`META;DateCreationExport;${fmtFRDate(now)}`);
  lines.push(`META;Reference_Capteur;${reference}`);
  lines.push(`META;Adresse_Capteur;${adresse}`);

  lines.push(`META;DateDebut_Extraite;${fmtFRDate(new Date(tFirst))}`);
  lines.push(`META;DateFin_Extraite;${fmtFRDate(new Date(tLast))}`);

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

  if (archiverPixels) {
    lines.push(`PIXELS;Index;X_screen;Y_screen`);
    pix.forEach((p, idx) => {
      lines.push(`PIXELS;${idx};${Math.round(p[0])};${Math.round(p[1])}`);
    });
  }

  // --------------------------
  // 6) Export
  // --------------------------

  try {
    downloadFileUserClick(baseName + ".csv", lines.join("\n"), "Télécharger CSV");
    alert("Export prêt. Boutons en bas à droite : " + baseName);
  } catch (e) {
    alert("Erreur export UI : " + (e && e.message ? e.message : e));
    console.error(e);
  }

})();
