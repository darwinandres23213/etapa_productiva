(() => {
  "use strict";

  const SENA_COLORS = ["#39a900", "#0f3d56", "#2d8600", "#b45309", "#5b7c5a", "#1f6f8b", "#7a9e4f", "#355c3a"];

  const DETAIL_FIELDS = [
    ["documento", "Documento"], ["tipoDoc", "Tipo documento"], ["nombres", "Nombres"], ["apellidos", "Apellidos"],
    ["telefono", "Teléfono"], ["correo", "Correo"], ["ficha", "Ficha"], ["instructorAsignado", "Instructor asignado"],
    ["instructor", "Instructor (Integral)"], ["programa", "Programa"], ["codPrograma", "Código programa"],
    ["nivel", "Nivel"], ["estadoAprendiz", "Estado aprendiz"], ["estadoFicha", "Estado ficha"],
    ["jornada", "Jornada"], ["modalidad", "Modalidad"], ["etapa", "Etapa"], ["fase", "Fase"],
    ["alternativa", "Alternativa"], ["empresa", "Empresa"], ["nit", "NIT"], ["ciudadEmpresa", "Ciudad empresa"],
    ["deptoEmpresa", "Depto. empresa"], ["direccionEmpresa", "Dirección"], ["telEmpresa", "Tel. empresa"],
    ["correoEmpresa", "Correo empresa"], ["centro", "Centro"], ["regional", "Regional"],
    ["especialidad", "Especialidad"], ["inicioLectiva", "Inicio lectiva"], ["finLectiva", "Fin lectiva"],
    ["inicioProductiva", "Inicio productiva"], ["fechaFin", "Fecha fin"], ["inicioEP", "Inicio EP (alt)"],
    ["finEP", "Fin EP (alt)"], ["contratoInicio", "Contrato inicio"], ["contratoFin", "Contrato fin"],
    ["calculoFinEP", "Cálculo fin EP"], ["estadoRap", "Estado RAP"], ["estadoCmp", "Estado CMP"],
    ["rapsOk", "RAPS aprobados"], ["rapsNo", "RAPS no aprobados"], ["rapsPend", "RAPS por evaluar"],
    ["rapsSum", "Sumatoria RAPS"], ["fechaRap", "Fecha emisión RAP"], ["juicioEP", "Juicio evaluación EP"],
    ["visitas", "Nº visitas"], ["certificado", "Certificado"], ["autorizacion", "Autorización"],
    ["gfpi", "GFPI-F023"], ["detalleBitacoras", "Detalle bitácoras"],
    ["bit1", "Bitácora 1"], ["bit2", "Bitácora 2"], ["bit3", "Bitácora 3"], ["bit4", "Bitácora 4"],
    ["bit5", "Bitácora 5"], ["bit6", "Bitácora 6"], ["bit7", "Bitácora 7"], ["bit8", "Bitácora 8"],
    ["bit9", "Bitácora 9"], ["bit10", "Bitácora 10"], ["bit11", "Bitácora 11"], ["bit12", "Bitácora 12"],
    ["otrosDocs", "Otros documentos"], ["reconocimiento", "Reconocimiento"], ["comite", "Comité académico"],
    ["acta", "Acta"], ["desercion", "Deserción"], ["obsFicha", "Obs. ficha"], ["observaciones", "Observaciones"],
    ["fuente", "Fuente"], ["fechaActSVP", "Actualización SVP"],
  ];

  const state = {
    data: null,
    planByFicha: new Map(),
    fichaResp: new Map(),
    instructorIndex: [],
    tab: "aprendices",
    mode: "todo",
    query: "",
    filters: {},
    filtered: [],
    filteredFichas: [],
    filteredPlan: [],
    summary: null,
    page: 1,
    pageSize: 25,
    sortKey: "apellidos",
    sortDir: 1,
    charts: {},
    chartsOpen: false,
    selectedId: null,
    pendingUpload: null,
    lastDiff: null,
  };

  const $ = (id) => document.getElementById(id);

  function showBusy(text) {
    const el = $("busy");
    if (!el) return;
    $("busyText").textContent = text || "Procesando…";
    el.hidden = false;
  }

  function hideBusy() {
    const el = $("busy");
    if (el) el.hidden = true;
  }

  function esc(s) {
    return String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function norm(v) {
    return String(v ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .trim();
  }

  function fichaKey(v) {
    if (v == null || v === "") return "";
    return String(v).replace(/\D/g, "") || String(v).trim();
  }

  function cleanInstructor(v) {
    if (v == null) return "";
    return String(v).replace(/\*+$/, "").trim();
  }

  function resolveInstructor(row) {
    const fromRow = cleanInstructor(row.instructor);
    if (fromRow) return fromRow;
    const fk = fichaKey(row.ficha);
    if (fk && state.planByFicha.has(fk)) return state.planByFicha.get(fk);
    if (fk && state.fichaResp.has(fk)) return state.fichaResp.get(fk);
    return "";
  }

  function enrichRows() {
    for (const r of state.data.aprendices) {
      r.instructorAsignado = resolveInstructor(r);
      // rebuild search blob including assigned instructor
      const parts = [];
      for (const [k, v] of Object.entries(r)) {
        if (k === "_q" || v == null) continue;
        parts.push(String(v));
      }
      r._q = parts.join(" ").toLowerCase();
    }
    for (const f of state.data.fichas) {
      const fk = fichaKey(f.ficha);
      f.instructorAsignado = cleanInstructor(f.responsable) || state.planByFicha.get(fk) || "";
      const parts = Object.values(f).filter((v) => v != null && v !== f._q).map(String);
      f._q = parts.join(" ").toLowerCase();
    }
  }

  function buildIndexes() {
    state.planByFicha = new Map();
    for (const p of state.data.plan) {
      const fk = fichaKey(p.ficha);
      const resp = cleanInstructor(p.responsable);
      if (fk && resp && !state.planByFicha.has(fk)) state.planByFicha.set(fk, resp);
    }
    state.fichaResp = new Map();
    for (const f of state.data.fichas) {
      const fk = fichaKey(f.ficha);
      const resp = cleanInstructor(f.responsable);
      if (fk && resp) state.fichaResp.set(fk, resp);
    }

    const names = new Set();
    const addName = (raw) => {
      const n = cleanInstructor(raw);
      if (!n || n.length < 3) return;
      if (/^\d+$/.test(n)) return;
      names.add(n);
    };
    for (const r of state.data.aprendices) addName(r.instructor);
    for (const p of state.data.plan) addName(p.responsable);
    for (const f of state.data.fichas) addName(f.responsable);
    state.instructorIndex = [...names].sort((a, b) => a.localeCompare(b, "es"));
  }

  function fillSelect(id, values, blank) {
    const el = $(id);
    const current = el.value;
    el.innerHTML = `<option value="">${blank}</option>`;
    values.forEach((v) => {
      const opt = document.createElement("option");
      opt.value = v;
      opt.textContent = v;
      el.appendChild(opt);
    });
    if ([...el.options].some((o) => o.value === current)) el.value = current;
  }

  function badgeForEstado(v) {
    const n = norm(v);
    if (!v) return `<span class="badge badge-muted">—</span>`;
    if (n.includes("certific") || n.includes("aprob")) return `<span class="badge badge-ok">${esc(v)}</span>`;
    if (n.includes("cancel") || n.includes("retiro") || n.includes("desert")) return `<span class="badge badge-danger">${esc(v)}</span>`;
    if (n.includes("aplaz") || n.includes("traslad") || n.includes("por_evaluar") || n.includes("pend")) return `<span class="badge badge-warn">${esc(v)}</span>`;
    if (n.includes("formacion") || n.includes("ejecucion") || n.includes("induc")) return `<span class="badge badge-info">${esc(v)}</span>`;
    return `<span class="badge badge-muted">${esc(v)}</span>`;
  }

  function modeMatch(row, kind) {
    const q = state.query;
    if (!q) return true;
    const tokens = q.split(/\s+/).filter(Boolean);

    if (state.mode === "todo") {
      const hay = row._q || "";
      return tokens.every((t) => hay.includes(t));
    }

    if (state.mode === "aprendiz") {
      const blob = norm([row.documento, row.documentoSVP, row.nombres, row.apellidos, row.correo, row.telefono].join(" "));
      return tokens.every((t) => blob.includes(t));
    }

    if (state.mode === "instructor") {
      const blob = norm([row.instructorAsignado, row.instructor, row.responsable].join(" "));
      return tokens.every((t) => blob.includes(t));
    }

    if (state.mode === "ficha") {
      const fk = fichaKey(row.ficha);
      const qDigits = q.replace(/\D/g, "");
      if (qDigits && fk) return fk.includes(qDigits) || fk === qDigits;
      const blob = norm(String(row.ficha ?? ""));
      return tokens.every((t) => blob.includes(t));
    }
    return true;
  }

  function sideFilters(row, type) {
    const f = state.filters;
    if (type === "aprendiz") {
      const checks = [
        ["estadoAprendiz", f.estado],
        ["nivel", f.nivel],
        ["jornada", f.jornada],
        ["modalidad", f.modalidad],
        ["etapa", f.etapa],
        ["estadoFicha", f.estadoFicha],
        ["alternativa", f.alternativa],
        ["estadoRap", f.rap],
        ["estadoCmp", f.cmp],
        ["programa", f.programa],
        ["fuente", f.fuente],
      ];
      for (const [key, val] of checks) {
        if (val && String(row[key] ?? "") !== val) return false;
      }
      if (f.instructor) {
        const blob = norm(row.instructorAsignado || row.instructor || "");
        if (!blob.includes(norm(f.instructor))) return false;
      }
      if (f.ficha && fichaKey(row.ficha) !== fichaKey(f.ficha)) return false;
      if (f.empresa) {
        const blob = norm(`${row.empresa ?? ""} ${row.nit ?? ""}`);
        if (!blob.includes(norm(f.empresa))) return false;
      }
      return true;
    }

    if (f.ficha && fichaKey(row.ficha) !== fichaKey(f.ficha)) return false;
    if (f.modalidad && row.modalidad && String(row.modalidad) !== f.modalidad) return false;
    if (f.jornada && row.jornada && String(row.jornada) !== f.jornada) return false;
    if (f.etapa && row.etapa && String(row.etapa) !== f.etapa) return false;
    if (f.programa && row.programa && String(row.programa) !== f.programa) return false;
    if (f.nivel && row.nivel && String(row.nivel) !== f.nivel) return false;
    if (f.estadoFicha && row.estado && String(row.estado) !== f.estadoFicha) return false;
    if (f.instructor) {
      const blob = norm(row.instructorAsignado || row.responsable || row.instructor || "");
      if (!blob.includes(norm(f.instructor))) return false;
    }
    if (f.empresa) {
      const blob = norm(`${row.empresa ?? ""} ${row.especialidad ?? ""}`);
      if (!blob.includes(norm(f.empresa))) return false;
    }
    return true;
  }

  function applyFilters(opts = {}) {
    const withBusy = opts.busy !== false;
    if (withBusy) showBusy(opts.busyText || "Buscando…");
    const run = () => {
      try {
        state.filtered = state.data.aprendices.filter((r) => modeMatch(r, "aprendiz") && sideFilters(r, "aprendiz"));
        state.filteredFichas = state.data.fichas.filter((r) => modeMatch(r, "ficha") && sideFilters(r, "ficha"));
        state.filteredPlan = state.data.plan.filter((r) => {
          const fake = { ...r, instructorAsignado: cleanInstructor(r.responsable), instructor: r.responsable };
          return modeMatch(fake, "plan") && sideFilters(fake, "plan");
        });
        state.page = 1;
        sortFiltered();
        state.summary = buildSummary();
        renderAll();
      } finally {
        if (withBusy) hideBusy();
      }
    };
    // Deja pintar el loader antes del filtro pesado
    if (withBusy) setTimeout(run, 20);
    else run();
  }

  function countMap(rows, getter) {
    const map = new Map();
    for (const r of rows) {
      const v = getter(r);
      const key = v ? String(v) : "(sin dato)";
      map.set(key, (map.get(key) || 0) + 1);
    }
    return [...map.entries()].sort((a, b) => b[1] - a[1]);
  }

  function buildSummary() {
    const rows = state.filtered;
    const q = state.query || state.filters.ficha || state.filters.instructor || "";
    const fichas = new Set(rows.map((r) => fichaKey(r.ficha)).filter(Boolean));
    const instructors = countMap(rows, (r) => r.instructorAsignado || "");
    const programas = countMap(rows, (r) => r.programa || "").slice(0, 5);
    const estados = countMap(rows, (r) => r.estadoAprendiz || "").slice(0, 5);

    let title = "Resultados de la consulta";
    if (state.mode === "aprendiz" && q) title = `Búsqueda por aprendiz: “${q}”`;
    if (state.mode === "instructor" && q) title = `Búsqueda por instructor: “${q}”`;
    if (state.mode === "ficha" && q) title = `Búsqueda por ficha: “${q}”`;
    if (state.mode === "todo" && q) title = `Búsqueda general: “${q}”`;
    if (!q && state.filters.instructor) title = `Filtro instructor: ${state.filters.instructor}`;
    if (!q && state.filters.ficha) title = `Filtro ficha: ${state.filters.ficha}`;

    // If searching a single ficha-like query, surface plan instructor explicitly
    const focusFicha = state.mode === "ficha" || state.filters.ficha ? fichaKey(state.filters.ficha || state.query) : "";
    let planHit = null;
    if (focusFicha) {
      planHit = state.data.plan.find((p) => fichaKey(p.ficha) === focusFicha) || null;
    }

    return {
      title,
      aprendices: rows.length,
      fichas: fichas.size,
      plan: state.filteredPlan.length,
      instructors,
      programas,
      estados,
      planHit,
      topInstructor: instructors.find((x) => x[0] !== "(sin dato)") || null,
      hasQuery: Boolean(q || state.filters.instructor || state.filters.ficha || state.filters.estado),
    };
  }

  function sortFiltered() {
    const k = state.sortKey;
    const d = state.sortDir;
    state.filtered.sort((a, b) => {
      const av = a[k];
      const bv = b[k];
      if (av == null && bv == null) return 0;
      if (av == null) return 1;
      if (bv == null) return -1;
      const an = Number(av);
      const bn = Number(bv);
      if (!Number.isNaN(an) && !Number.isNaN(bn) && String(av).trim() !== "" && String(bv).trim() !== "") {
        return (an - bn) * d;
      }
      return String(av).localeCompare(String(bv), "es", { sensitivity: "base" }) * d;
    });
  }

  function renderBanner() {
    const s = state.summary;
    const el = $("resultBanner");
    if (!s || !s.hasQuery) {
      el.hidden = true;
      el.innerHTML = "";
      return;
    }
    el.hidden = false;
    const instrChips = s.instructors.slice(0, 6).map(([name, n]) =>
      `<span class="chip"><strong>${esc(name)}</strong> · ${n} apr.</span>`
    ).join("");

    const planLine = s.planHit
      ? `<div><b>Instructor en plan (ficha ${esc(s.planHit.ficha)})</b>${esc(cleanInstructor(s.planHit.responsable) || "Sin responsable")}</div>
         <div><b>Especialidad plan</b>${esc(s.planHit.especialidad || "—")}</div>`
      : s.topInstructor
        ? `<div><b>Instructor más frecuente</b>${esc(s.topInstructor[0])} (${s.topInstructor[1]} aprendices)</div>`
        : `<div><b>Instructor</b>Sin dato en la selección</div>`;

    el.innerHTML = `
      <div class="rb-title">${esc(s.title)}</div>
      <div class="rb-grid">
        <div><b>Aprendices encontrados</b>${s.aprendices.toLocaleString("es-CO")}</div>
        <div><b>Fichas distintas</b>${s.fichas.toLocaleString("es-CO")}</div>
        <div><b>En plan instructores</b>${s.plan.toLocaleString("es-CO")}</div>
        ${planLine}
      </div>
      <div class="chips">${instrChips || '<span class="chip">Sin instructores identificados</span>'}</div>`;
  }

  function renderKpis() {
    const rows = state.filtered;
    const fichas = new Set(rows.map((r) => r.ficha).filter(Boolean));
    const conEmpresa = rows.filter((r) => r.empresa).length;
    const enFormacion = rows.filter((r) => /formacion|induccion/i.test(String(r.estadoAprendiz || ""))).length;
    const certificados = rows.filter((r) => /certific/i.test(String(r.estadoAprendiz || ""))).length;
    const conInst = rows.filter((r) => r.instructorAsignado).length;

    $("kpis").innerHTML = [
      ["Aprendices", rows.length, `de ${state.data.meta.totales.aprendices}`],
      ["Fichas", fichas.size, "en selección"],
      ["Con instructor", conInst, "asignado / plan"],
      ["En formación", enFormacion, "estado"],
      ["Certificados", certificados, "estado"],
      ["Con empresa", conEmpresa, "razón social"],
    ].map(([label, value, hint]) => `
      <article class="kpi">
        <div class="label">${esc(label)}</div>
        <div class="value">${Number(value).toLocaleString("es-CO")}</div>
        <div class="hint">${esc(hint)}</div>
      </article>`).join("");

    $("tabAprCount").textContent = rows.length.toLocaleString("es-CO");
    $("tabFicCount").textContent = state.filteredFichas.length.toLocaleString("es-CO");
    $("tabPlanCount").textContent = state.filteredPlan.length.toLocaleString("es-CO");
  }

  function upsertChart(id, type, labels, values) {
    const canvas = $(id);
    if (!canvas) return;
    if (state.charts[id]) state.charts[id].destroy();
    const total = values.reduce((a, b) => a + b, 0) || 1;
    state.charts[id] = new Chart(canvas, {
      type,
      data: {
        labels,
        datasets: [{
          data: values,
          backgroundColor: labels.map((_, i) => SENA_COLORS[i % SENA_COLORS.length]),
          borderWidth: type === "bar" ? 0 : 1,
          borderColor: "#fff",
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: type !== "bar",
            position: "bottom",
            labels: { boxWidth: 10, font: { size: 10, family: "Public Sans" } },
          },
          tooltip: {
            callbacks: {
              label(ctx) {
                const v = ctx.parsed.y ?? ctx.parsed;
                return ` ${Number(v).toLocaleString("es-CO")} (${((v / total) * 100).toFixed(1)}%)`;
              },
            },
          },
        },
        scales: type === "bar" ? {
          x: { ticks: { font: { size: 10 }, maxRotation: 40 } },
          y: { beginAtZero: true, ticks: { precision: 0 } },
        } : undefined,
      },
    });
  }

  function renderCharts() {
    const estados = countMap(state.filtered, (r) => r.estadoAprendiz || "").slice(0, 8);
    upsertChart("chartEstado", "bar", estados.map((x) => x[0]), estados.map((x) => x[1]));
    const niveles = countMap(state.filtered, (r) => r.nivel || "");
    upsertChart("chartNivel", "doughnut", niveles.map((x) => x[0]), niveles.map((x) => x[1]));
    const alts = countMap(state.filtered, (r) => r.alternativa || "").slice(0, 7);
    upsertChart("chartAlt", "doughnut", alts.map((x) => x[0]), alts.map((x) => x[1]));
  }

  function fullName(r) {
    return `${r.nombres || ""} ${r.apellidos || ""}`.trim() || "Sin nombre";
  }

  function renderTable() {
    const start = (state.page - 1) * state.pageSize;
    const pageRows = state.filtered.slice(start, start + state.pageSize);
    const host = $("listAprendices");

    if (!pageRows.length) {
      host.innerHTML = `<div class="empty">Sin resultados. Pruebe otro modo o limpie filtros.</div>`;
    } else {
      const cards = pageRows.map((r, idx) => {
        const abs = start + idx;
        const id = `${r.documento || ""}-${r.ficha || ""}-${abs}`;
        const active = state.selectedId === id ? "active" : "";
        return `<article class="row-card mobile-only ${active}" data-idx="${abs}" data-id="${esc(id)}">
          <div class="who">
            <div class="name">${esc(fullName(r))}</div>
            <div class="meta">
              <span>Doc. <strong>${esc(r.documento || "—")}</strong></span>
              <span>Ficha <strong>${esc(r.ficha || "—")}</strong></span>
              <span>${badgeForEstado(r.estadoAprendiz)}</span>
            </div>
            <div class="meta">
              <span>Instructor: <strong>${esc(r.instructorAsignado || "sin asignar")}</strong></span>
            </div>
          </div>
          <div class="actions">
            <button type="button" class="btn btn-detalle" data-detalle="${abs}">Detalle</button>
          </div>
        </article>`;
      }).join("");

      const rows = pageRows.map((r, idx) => {
        const abs = start + idx;
        const id = `${r.documento || ""}-${r.ficha || ""}-${abs}`;
        const active = state.selectedId === id ? "active" : "";
        return `<tr class="${active}" data-idx="${abs}" data-id="${esc(id)}">
          <td class="col-name">
            <strong>${esc(fullName(r))}</strong>
            <span class="sub">Doc. ${esc(r.documento || "—")}</span>
          </td>
          <td>${esc(r.ficha || "—")}</td>
          <td>${badgeForEstado(r.estadoAprendiz)}</td>
          <td class="clip" title="${esc(r.instructorAsignado || "")}">${esc(r.instructorAsignado || "—")}</td>
          <td><button type="button" class="btn btn-sm btn-detalle" data-detalle="${abs}">Detalle</button></td>
        </tr>`;
      }).join("");

      host.innerHTML = `
        ${cards}
        <table class="compact-table desktop-only">
          <thead>
            <tr>
              <th>Aprendiz</th>
              <th>Ficha</th>
              <th>Estado</th>
              <th>Instructor</th>
              <th>Detalle</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>`;
    }

    const total = state.filtered.length;
    const pages = Math.max(1, Math.ceil(total / state.pageSize) || 1);
    $("resultCount").textContent = `${total.toLocaleString("es-CO")} resultados · pág. ${state.page}/${pages} · use Detalle para ver todo`;
    $("pageInfo").textContent = total ? `${start + 1}–${Math.min(start + state.pageSize, total)} de ${total.toLocaleString("es-CO")}` : "Sin filas";
    $("prevPage").disabled = state.page <= 1;
    $("nextPage").disabled = state.page >= pages;
  }

  function renderFichas() {
    const rows = state.filteredFichas;
    $("fichasCount").textContent = `${rows.length.toLocaleString("es-CO")} fichas · toque Detalle en aprendices para ver más`;
    const host = $("listFichas");
    if (!rows.length) {
      host.innerHTML = `<div class="empty">Sin fichas para el criterio.</div>`;
      return;
    }
    host.innerHTML = rows.slice(0, 400).map((r) => {
      const inst = r.instructorAsignado || resolveInstructor({ ficha: r.ficha, instructor: r.responsable }) || "—";
      return `<article class="row-card">
        <div class="who">
          <div class="name">Ficha ${esc(r.ficha || "—")}</div>
          <div class="meta">
            <span>${badgeForEstado(r.estado)}</span>
            <span>${esc(r.nivel || "")}</span>
            <span>${esc(r.jornada || "")}</span>
          </div>
          <div class="meta">
            <span class="clip" title="${esc(r.programa || "")}"><strong>${esc(String(r.programa || "—").slice(0, 60))}</strong></span>
          </div>
          <div class="meta">Instructor: <strong>${esc(inst)}</strong></div>
        </div>
      </article>`;
    }).join("");
  }

  function renderPlan() {
    const rows = state.filteredPlan;
    $("planCount").textContent = `${rows.length.toLocaleString("es-CO")} asignaciones`;
    const host = $("listPlan");
    if (!rows.length) {
      host.innerHTML = `<div class="empty">Sin registros de plan.</div>`;
      return;
    }
    host.innerHTML = rows.map((r) => {
      const obs = String(r.observaciones || "");
      return `<article class="row-card">
        <div class="who">
          <div class="name">${esc(cleanInstructor(r.responsable) || "Sin responsable")}</div>
          <div class="meta">
            <span>Ficha <strong>${esc(r.ficha || "—")}</strong></span>
            <span>${esc(r.modalidad || "")}</span>
          </div>
          <div class="meta">${esc(String(r.especialidad || "—").slice(0, 80))}</div>
          ${obs ? `<div class="meta">${esc(obs.slice(0, 120))}${obs.length > 120 ? "…" : ""}</div>` : ""}
        </div>
      </article>`;
    }).join("");
  }

  function renderAll() {
    renderBanner();
    renderKpis();
    if (state.tab === "aprendices") {
      if (state.chartsOpen) renderCharts();
      renderTable();
    } else if (state.tab === "fichas") {
      renderFichas();
    } else {
      renderPlan();
    }
  }

  function openDrawer(row) {
    $("drawerTitle").textContent = `${row.nombres || ""} ${row.apellidos || ""}`.trim() || "Aprendiz";
    $("drawerSub").textContent = `Doc. ${row.documento || "—"} · Ficha ${row.ficha || "—"} · Instructor: ${row.instructorAsignado || "sin asignar"}`;
    $("drawerBody").innerHTML = DETAIL_FIELDS
      .filter(([k]) => row[k] != null && row[k] !== "")
      .map(([k, label]) => `<div class="field"><div class="k">${esc(label)}</div><div class="v">${esc(row[k])}</div></div>`)
      .join("");
    $("drawer").classList.add("open");
    $("drawerBackdrop").classList.add("open");
    $("drawer").setAttribute("aria-hidden", "false");
    $("drawerBody").scrollTop = 0;
  }

  function closeDrawer() {
    $("drawer").classList.remove("open");
    $("drawerBackdrop").classList.remove("open");
    $("drawer").setAttribute("aria-hidden", "true");
  }

  function setFiltersOpen(open) {
    $("filtersPanel").dataset.open = open ? "true" : "false";
    $("btnToggleFilters").setAttribute("aria-expanded", open ? "true" : "false");
    $("filtersBackdrop").hidden = !open;
  }

  function updateModeHint() {
    const hints = {
      todo: "Busca en todos los campos (documento, nombre, ficha, instructor, empresa, correo…).",
      aprendiz: "Solo documento, nombres, apellidos, correo y teléfono. Verá ficha e instructor asignado.",
      instructor: "Filtra por instructor/responsable. Cuenta aprendices y fichas a cargo.",
      ficha: "Escriba el número de ficha. Muestra aprendices, instructor del plan y totales.",
    };
    $("lookupHint").textContent = hints[state.mode] || hints.todo;
    const ph = {
      todo: "Documento, nombre, ficha, instructor, correo, empresa…",
      aprendiz: "Documento o nombre del aprendiz…",
      instructor: "Nombre del instructor / responsable…",
      ficha: "Número de ficha (ej. 2740649)…",
    };
    $("globalSearch").placeholder = ph[state.mode] || ph.todo;
  }

  function readFiltersFromUI() {
    state.filters = {
      estado: $("fEstado").value,
      nivel: $("fNivel").value,
      jornada: $("fJornada").value,
      modalidad: $("fModalidad").value,
      etapa: $("fEtapa").value,
      estadoFicha: $("fEstadoFicha").value,
      alternativa: $("fAlternativa").value,
      instructor: $("fInstructor").value,
      rap: $("fRap").value,
      cmp: $("fCmp").value,
      programa: $("fPrograma").value,
      fuente: $("fFuente").value,
      ficha: $("fFicha").value.trim(),
      empresa: $("fEmpresa").value.trim(),
    };
    state.query = norm($("globalSearch").value);
    state.pageSize = Number($("pageSize").value) || 50;
  }

  function exportCsv() {
    const rows = state.tab === "aprendices" ? state.filtered : state.tab === "fichas" ? state.filteredFichas : state.filteredPlan;
    if (!rows.length) return alert("No hay filas para exportar.");
    const keys = Object.keys(rows[0]).filter((k) => k !== "_q");
    const lines = [keys.join(",")];
    for (const r of rows) {
      lines.push(keys.map((k) => {
        const raw = r[k] == null ? "" : String(r[k]).replace(/\r?\n/g, " ");
        return /[",]/.test(raw) ? `"${raw.replace(/"/g, '""')}"` : raw;
      }).join(","));
    }
    const blob = new Blob(["\ufeff" + lines.join("\n")], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `sena-consulta-${state.tab}.csv`;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function refreshFilterOptions() {
    const f = state.data.meta.filtros || {};
    fillSelect("fEstado", f.estadoAprendiz || [], "Todos");
    fillSelect("fNivel", f.nivel || [], "Todos");
    fillSelect("fJornada", f.jornada || [], "Todas");
    fillSelect("fModalidad", f.modalidad || [], "Todas");
    fillSelect("fEtapa", f.etapa || [], "Todas");
    fillSelect("fEstadoFicha", f.estadoFicha || [], "Todos");
    fillSelect("fAlternativa", f.alternativa || [], "Todas");
    fillSelect("fInstructor", state.instructorIndex.length ? state.instructorIndex : (f.instructor || []), "Todos");
    fillSelect("fRap", f.estadoRap || [], "Todos");
    fillSelect("fCmp", f.estadoCmp || [], "Todos");
    fillSelect("fPrograma", f.programa || [], "Todos");
    fillSelect("fFuente", f.fuente || [], "Todas");
  }

  function replaceWorkingData(data, note) {
    showBusy("Actualizando consulta…");
    setTimeout(() => {
      try {
        state.data = data;
        buildIndexes();
        enrichRows();
        refreshFilterOptions();
        $("metaStamp").textContent = `${state.data.meta.generado} · ${state.data.meta.totales.aprendices.toLocaleString("es-CO")} apr.${note ? " · " + note : ""}`;
        readFiltersFromUI();
        applyFilters({ busy: false });
      } finally {
        hideBusy();
      }
    }, 20);
  }

  function renderValidationBlock(validation) {
    if (!validation) return "";
    const errHtml = validation.errors?.length
      ? `<div class="validation-box err">
          <h3>Error de formato / inconsistencias (${validation.errors.length})</h3>
          <ul>${validation.errors.map((e) => `<li><strong>${esc(e.code)}</strong> — ${esc(e.msg)}${e.detail ? `<span class="detail">${esc(e.detail)}</span>` : ""}</li>`).join("")}</ul>
        </div>`
      : "";
    const warnHtml = validation.warnings?.length
      ? `<div class="validation-box warn">
          <h3>Advertencias (${validation.warnings.length})</h3>
          <ul>${validation.warnings.map((e) => `<li><strong>${esc(e.code)}</strong> — ${esc(e.msg)}${e.detail ? `<span class="detail">${esc(e.detail)}</span>` : ""}</li>`).join("")}</ul>
        </div>`
      : "";
    if (!errHtml && !warnHtml) {
      return `<div class="validation-box" style="border-left:4px solid var(--sena)"><h3>Formato válido</h3><p class="lookup-hint" style="margin:0">Hojas y columnas coinciden con el Excel unificado esperado.</p></div>`;
    }
    return `${errHtml}${warnHtml}`;
  }

  function renderDiffReport(diff, parsed) {
    const box = $("diffReport");
    const status = $("uploadStatus");
    const validation = parsed.validation || { ok: true, errors: [], warnings: [] };
    const hasErrors = !validation.ok || (validation.errors && validation.errors.length > 0);

    status.hidden = false;
    status.classList.remove("is-error", "is-ok");
    if (hasErrors) {
      status.classList.add("is-error");
      status.textContent = `ERROR: el Excel no pasó la validación (${validation.errors.length} problema(s)). No se puede aplicar a la consulta hasta corregirlo.`;
    } else {
      status.classList.add("is-ok");
      status.textContent = `Formato OK · Leído: ${parsed.meta.totales.aprendices.toLocaleString("es-CO")} aprendices · ${parsed.meta.totales.fichas} fichas · ${parsed.meta.totales.plan} plan · hojas: ${parsed.sheets.integral || "—"}${parsed.sheets.fichas ? ", " + parsed.sheets.fichas : ""}${parsed.sheets.plan ? ", " + parsed.sheets.plan : ""}`;
    }

    const addedPreview = diff.added.slice(0, 40).map((r) => {
      const name = `${r.nombres || ""} ${r.apellidos || ""}`.trim() || "Sin nombre";
      return `<div class="item"><strong>${esc(name)}</strong> · Doc ${esc(r.documento || "—")} · Ficha ${esc(r.ficha || "—")} · ${badgeForEstado(r.estadoAprendiz)}</div>`;
    }).join("") || `<div class="item">No hay aprendices nuevos.</div>`;

    const changedPreview = diff.changed.slice(0, 40).map((c) => {
      const r = c.row;
      const name = `${r.nombres || ""} ${r.apellidos || ""}`.trim() || "Sin nombre";
      const ch = c.fields.slice(0, 4).map((f) => `${f.field}: ${f.before ?? "∅"} → ${f.after ?? "∅"}`).join(" · ");
      return `<div class="item"><strong>${esc(name)}</strong> · Ficha ${esc(r.ficha || "—")}<div class="chg">${esc(ch)}</div></div>`;
    }).join("") || `<div class="item">Sin cambios de campos vigilados.</div>`;

    box.hidden = false;
    box.innerHTML = `
      ${renderValidationBlock(validation)}
      ${hasErrors ? "" : `
      <div>
        <strong>Comparación con la base actual</strong>
        <div class="lookup-hint">Antes ${diff.before.aprendices.toLocaleString("es-CO")} → ahora ${diff.after.aprendices.toLocaleString("es-CO")}</div>
      </div>
      <div class="diff-kpis">
        <div class="diff-kpi pos"><b>Aprendices nuevos</b><strong>+${diff.added.length}</strong></div>
        <div class="diff-kpi neg"><b>Ya no aparecen</b><strong>${diff.removed.length}</strong></div>
        <div class="diff-kpi warn"><b>Con cambios</b><strong>${diff.changed.length}</strong></div>
        <div class="diff-kpi"><b>Delta total</b><strong>${diff.deltaAprendices >= 0 ? "+" : ""}${diff.deltaAprendices}</strong></div>
        <div class="diff-kpi pos"><b>Fichas nuevas</b><strong>+${diff.fichasAdded.length}</strong></div>
        <div class="diff-kpi"><b>Plan (+ / −)</b><strong>+${diff.planAdded} / −${diff.planRemoved}</strong></div>
      </div>
      <div class="diff-actions">
        <button type="button" class="btn btn-detalle" id="btnApplyUpload">Usar este Excel en la consulta</button>
        <button type="button" class="btn" id="btnShowOnlyNew" ${diff.added.length ? "" : "disabled"}>Ver solo nuevos en la lista</button>
        <button type="button" class="btn" id="btnDownloadDiffCsv" ${diff.added.length || diff.changed.length ? "" : "disabled"}>CSV de novedades</button>
      </div>
      <div>
        <strong>Aprendices nuevos (${diff.added.length})</strong>
        <div class="diff-list">${addedPreview}${diff.added.length > 40 ? `<div class="item">… y ${diff.added.length - 40} más</div>` : ""}</div>
      </div>
      <div>
        <strong>Cambios detectados (${diff.changed.length})</strong>
        <div class="diff-list">${changedPreview}${diff.changed.length > 40 ? `<div class="item">… y ${diff.changed.length - 40} más</div>` : ""}</div>
      </div>
      ${diff.fichasAdded.length ? `<div class="lookup-hint">Fichas nuevas: ${esc(diff.fichasAdded.slice(0, 20).join(", "))}${diff.fichasAdded.length > 20 ? "…" : ""}</div>` : ""}
      `}
      ${hasErrors ? `<div class="diff-actions"><button type="button" class="btn" id="btnForcePreview" disabled title="Corrija el Excel">Aplicar bloqueado por errores</button></div>` : ""}
    `;

    if (hasErrors) return;

    $("btnApplyUpload").onclick = async () => {
      if (parsed.validation && !parsed.validation.ok) {
        alert("No se puede aplicar: el Excel tiene errores de formato.");
        return;
      }
      replaceWorkingData(parsed, "Excel aplicado");
      state.pendingUpload = parsed;
      state.lastDiff = diff;
      try {
        await window.ExcelImport.saveBaseline({
          savedAt: new Date().toISOString(),
          fileName: parsed.meta.fuente,
          totals: parsed.meta.totales,
          keys: parsed.aprendices.map((r) => window.ExcelImport.aprKey(r)),
          data: parsed,
        });
      } catch (e) {
        console.warn("No se pudo guardar baseline IndexedDB", e);
      }
      status.classList.remove("is-error");
      status.classList.add("is-ok");
      status.textContent = "Base de consulta actualizada con el Excel cargado. La próxima subida se comparará contra esta versión.";
    };

    $("btnShowOnlyNew").onclick = () => {
      if (!diff.added.length) return;
      replaceWorkingData({
        ...parsed,
        aprendices: diff.added,
        meta: {
          ...parsed.meta,
          totales: { ...parsed.meta.totales, aprendices: diff.added.length },
          generado: parsed.meta.generado + " (solo nuevos)",
        },
      }, "solo nuevos");
      $("uploadPanel").hidden = true;
    };

    $("btnDownloadDiffCsv").onclick = () => {
      const rows = [
        ...diff.added.map((r) => ({ tipo: "NUEVO", ...r })),
        ...diff.changed.map((c) => ({
          tipo: "CAMBIO",
          ...c.row,
          cambios: c.fields.map((f) => `${f.field}:${f.before}->${f.after}`).join(" | "),
        })),
      ];
      if (!rows.length) return;
      const keys = ["tipo", "documento", "nombres", "apellidos", "ficha", "estadoAprendiz", "programa", "instructor", "empresa", "cambios"];
      const lines = [keys.join(",")];
      for (const r of rows) {
        lines.push(keys.map((k) => {
          const raw = r[k] == null ? "" : String(r[k]).replace(/\r?\n/g, " ");
          return /[",]/.test(raw) ? `"${raw.replace(/"/g, '""')}"` : raw;
        }).join(","));
      }
      const blob = new Blob(["\ufeff" + lines.join("\n")], { type: "text/csv;charset=utf-8" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "sena-novedades-excel.csv";
      a.click();
      URL.revokeObjectURL(a.href);
    };
  }

  function bindUpload() {
    const panel = $("uploadPanel");
    const fileInput = $("excelFile");
    const drop = $("fileDrop");
    const analyzeBtn = $("btnAnalyze");

    $("btnOpenUpload").addEventListener("click", () => {
      panel.hidden = !panel.hidden;
    });
    $("btnCloseUpload").addEventListener("click", () => {
      panel.hidden = true;
    });

    const acceptFile = (file) => {
      if (!file) return;
      state._uploadFile = file;
      $("uploadFileName").textContent = `${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)`;
      analyzeBtn.disabled = false;
      $("diffReport").hidden = true;
      $("uploadStatus").hidden = true;
      $("uploadStatus").classList.remove("is-error", "is-ok");
      if (!/\.xlsx$/i.test(file.name || "") && !/\.xls$/i.test(file.name || "")) {
        $("uploadStatus").hidden = false;
        $("uploadStatus").classList.add("is-error");
        $("uploadStatus").textContent = "ERROR: extensión no válida. Use un archivo .xlsx";
        analyzeBtn.disabled = true;
      }
    };

    fileInput.addEventListener("change", () => {
      acceptFile(fileInput.files && fileInput.files[0]);
    });

    ["dragenter", "dragover"].forEach((ev) => {
      drop.addEventListener(ev, (e) => {
        e.preventDefault();
        drop.classList.add("drag");
      });
    });
    ["dragleave", "drop"].forEach((ev) => {
      drop.addEventListener(ev, (e) => {
        e.preventDefault();
        drop.classList.remove("drag");
      });
    });
    drop.addEventListener("drop", (e) => {
      acceptFile(e.dataTransfer.files && e.dataTransfer.files[0]);
    });

    $("uploadForm").addEventListener("submit", async (e) => {
      e.preventDefault();
      const file = state._uploadFile || (fileInput.files && fileInput.files[0]);
      if (!file) return;
      if (!window.ExcelImport || !window.XLSX) {
        alert("No se cargó la librería de Excel. Revise la conexión o recargue la página.");
        return;
      }
      analyzeBtn.disabled = true;
      const status = $("uploadStatus");
      status.hidden = false;
      status.classList.remove("is-error", "is-ok");
      status.textContent = "Leyendo Excel… esto puede tardar unos segundos con archivos grandes.";
      $("diffReport").hidden = true;
      showBusy("Importando y validando Excel…");
      try {
        // Validación previa de extensión
        if (!/\.xlsx$/i.test(file.name || "") && !/\.xls$/i.test(file.name || "")) {
          status.classList.add("is-error");
          status.textContent = "ERROR: el archivo no es .xlsx. Suba el Excel unificado de Etapa Productiva.";
          $("diffReport").hidden = false;
          $("diffReport").innerHTML = renderValidationBlock({
            ok: false,
            errors: [{ code: "EXT", msg: `Formato inválido: "${file.name}". Se espera .xlsx.` }],
            warnings: [],
          });
          return;
        }

        const parsed = await window.ExcelImport.readFile(file);
        showBusy("Comparando novedades…");
        await new Promise((r) => setTimeout(r, 30));
        const baseline = state.data;
        const diff = parsed.validation?.ok === false
          ? {
              before: { aprendices: baseline.aprendices.length, fichas: baseline.fichas.length, plan: baseline.plan.length },
              after: { aprendices: 0, fichas: 0, plan: 0 },
              added: [], removed: [], changed: [], fichasAdded: [], fichasRemoved: [],
              planAdded: 0, planRemoved: 0, deltaAprendices: 0,
            }
          : window.ExcelImport.compareData(baseline, parsed);
        state.pendingUpload = parsed.validation?.ok === false ? null : parsed;
        state.lastDiff = diff;
        renderDiffReport(diff, parsed);
      } catch (err) {
        console.error(err);
        status.hidden = false;
        status.classList.add("is-error");
        status.classList.remove("is-ok");
        status.textContent = `ERROR al leer el Excel: ${err.message || err}`;
        $("diffReport").hidden = false;
        $("diffReport").innerHTML = renderValidationBlock({
          ok: false,
          errors: [{ code: "LECTURA", msg: String(err.message || err) }],
          warnings: [],
        });
      } finally {
        hideBusy();
        analyzeBtn.disabled = false;
      }
    });
  }

  function bindEvents() {
    let timer = null;
    const refresh = () => {
      readFiltersFromUI();
      applyFilters();
    };
    const debounced = () => {
      clearTimeout(timer);
      timer = setTimeout(refresh, 160);
    };
    bindUpload();

    $("globalSearch").addEventListener("input", debounced);
    ["fEstado","fNivel","fJornada","fModalidad","fEtapa","fEstadoFicha","fAlternativa","fInstructor","fRap","fCmp","fPrograma","fFuente","pageSize"]
      .forEach((id) => $(id).addEventListener("change", refresh));
    ["fFicha","fEmpresa"].forEach((id) => $(id).addEventListener("input", debounced));

    $("btnToggleCharts").addEventListener("click", () => {
      state.chartsOpen = !state.chartsOpen;
      const wrap = $("chartsWrap");
      wrap.hidden = !state.chartsOpen;
      $("btnToggleCharts").textContent = state.chartsOpen ? "Ocultar gráficos" : "Ver gráficos";
      $("btnToggleCharts").setAttribute("aria-expanded", state.chartsOpen ? "true" : "false");
      if (state.chartsOpen) renderCharts();
    });

    document.querySelectorAll(".mode").forEach((btn) => {
      btn.addEventListener("click", () => {
        document.querySelectorAll(".mode").forEach((b) => b.classList.remove("active"));
        btn.classList.add("active");
        state.mode = btn.dataset.mode;
        updateModeHint();
        refresh();
        $("globalSearch").focus();
      });
    });

    $("btnClear").addEventListener("click", () => {
      $("globalSearch").value = "";
      refresh();
      $("globalSearch").focus();
    });

    $("btnResetFilters").addEventListener("click", () => {
      $("globalSearch").value = "";
      ["fEstado","fNivel","fJornada","fModalidad","fEtapa","fEstadoFicha","fAlternativa","fInstructor","fRap","fCmp","fPrograma","fFuente"]
        .forEach((id) => { $(id).value = ""; });
      $("fFicha").value = "";
      $("fEmpresa").value = "";
      setFiltersOpen(false);
      refresh();
    });

    $("btnToggleFilters").addEventListener("click", () => {
      const open = $("filtersPanel").dataset.open !== "true";
      setFiltersOpen(open);
    });
    $("btnCloseFilters").addEventListener("click", () => setFiltersOpen(false));
    $("filtersBackdrop").addEventListener("click", () => setFiltersOpen(false));

    $("btnExport").addEventListener("click", exportCsv);
    $("prevPage").addEventListener("click", () => {
      if (state.page > 1) { state.page -= 1; renderTable(); }
    });
    $("nextPage").addEventListener("click", () => {
      const pages = Math.max(1, Math.ceil(state.filtered.length / state.pageSize));
      if (state.page < pages) { state.page += 1; renderTable(); }
    });

    $("listAprendices").addEventListener("click", (e) => {
      const btn = e.target.closest("[data-detalle]");
      const rowEl = e.target.closest("[data-idx]");
      if (!btn && !rowEl) return;
      const idx = Number((btn || rowEl).getAttribute(btn ? "data-detalle" : "data-idx"));
      const row = state.filtered[idx];
      if (!row) return;
      state.selectedId = rowEl ? rowEl.getAttribute("data-id") : `${row.documento || ""}-${row.ficha || ""}-${idx}`;
      renderTable();
      openDrawer(row);
    });

    document.querySelectorAll(".tab").forEach((btn) => {
      btn.addEventListener("click", () => {
        document.querySelectorAll(".tab").forEach((b) => b.classList.remove("active"));
        btn.classList.add("active");
        state.tab = btn.dataset.tab;
        ["aprendices", "fichas", "plan"].forEach((t) => { $(`view-${t}`).hidden = t !== state.tab; });
        renderAll();
      });
    });

    ["drawerClose", "drawerBack", "drawerCloseBottom"].forEach((id) => {
      $(id).addEventListener("click", closeDrawer);
    });
    $("drawerBackdrop").addEventListener("click", closeDrawer);
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") { closeDrawer(); setFiltersOpen(false); hideBusy(); }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k") {
        e.preventDefault();
        $("globalSearch").focus();
        $("globalSearch").select();
      }
    });
  }

  async function boot() {
    try {
      const res = await fetch("data.json", { cache: "no-store" });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      state.data = await res.json();
      buildIndexes();
      enrichRows();

      const f = state.data.meta.filtros;
      fillSelect("fEstado", f.estadoAprendiz, "Todos");
      fillSelect("fNivel", f.nivel, "Todos");
      fillSelect("fJornada", f.jornada, "Todas");
      fillSelect("fModalidad", f.modalidad, "Todas");
      fillSelect("fEtapa", f.etapa, "Todas");
      fillSelect("fEstadoFicha", f.estadoFicha, "Todos");
      fillSelect("fAlternativa", f.alternativa, "Todas");
      fillSelect("fInstructor", state.instructorIndex.length ? state.instructorIndex : f.instructor, "Todos");
      fillSelect("fRap", f.estadoRap, "Todos");
      fillSelect("fCmp", f.estadoCmp, "Todos");
      fillSelect("fPrograma", f.programa, "Todos");
      fillSelect("fFuente", f.fuente, "Todas");

      $("metaStamp").textContent = `${state.data.meta.generado} · ${state.data.meta.totales.aprendices.toLocaleString("es-CO")} apr.`;
      updateModeHint();
      bindEvents();
      readFiltersFromUI();
      applyFilters({ busy: false });
      const loader = $("loading");
      if (loader) loader.remove();
      hideBusy();

      // Si hay una base aplicada antes en este navegador, avisar
      window.ExcelImport.loadBaseline().then((saved) => {
        if (!saved || !saved.fileName) return;
        const tip = document.createElement("div");
        tip.className = "lookup-hint";
        tip.style.cssText = "width:min(1280px,calc(100% - 1.25rem));margin:0.4rem auto 0;font-size:0.78rem;color:#5a6358";
        tip.textContent = `Último Excel aplicado en este equipo: ${saved.fileName} (${saved.totals?.aprendices ?? "?"} aprendices, ${saved.savedAt || ""}). Use «Actualizar Excel» para comparar uno nuevo.`;
        document.querySelector(".topbar")?.after(tip);
      }).catch(() => {});
    } catch (err) {
      console.error(err);
      $("loading").innerHTML = `<p><strong>No se pudo cargar.</strong></p><p>Use <code>iniciar.bat</code>.</p><p>${esc(err.message)}</p>`;
    }
  }

  window.__EP_API = {
    getData: () => state.data,
    replaceWorkingData,
    renderDiffReport,
  };

  boot();
})();
