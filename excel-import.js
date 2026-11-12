/* Parseo de Excel y detección de novedades vs base actual */
(() => {
  "use strict";

  const WATCH = [
    "estadoAprendiz", "estadoFicha", "programa", "nivel", "jornada", "modalidad",
    "etapa", "alternativa", "instructor", "empresa", "estadoRap", "estadoCmp",
  ];

  const INTEGRAL_MAP = {
    "Numero de Documento": "documento", Grupo: "ficha", Teléfono: "telefono",
    "Correo electónico": "correo", Nombres: "nombres", Apellidos: "apellidos",
    NIT: "nit", "Razon Social": "empresa", "Departamento empresa": "deptoEmpresa",
    "Ciudad empresa": "ciudadEmpresa", Dirección: "direccionEmpresa",
    "Teléfono empresa": "telEmpresa", "Correo electrónico": "correoEmpresa",
    "Tipo documento": "tipoDoc", "Centro ID": "centroId", Centro: "centro",
    Especialidad: "especialidad", "INICIO ETAPA LECTIVA": "inicioLectiva",
    "Fin lectiva": "finLectiva", "Inicio productiva": "inicioProductiva",
    "Fecha Fin": "fechaFin", "Contrato inicio": "contratoInicio", "Contrato Fin": "contratoFin",
    Regional: "regional", Fase: "fase", ESTADO_FICHA: "estadoFicha",
    CODIGO_PROGRAMA: "codPrograma", VERSION_PROGRANA: "versionPrograma",
    PROGRAMA: "programa", NIVEL_DE_FORMACION: "nivel", ESTADO_APRENDIZ: "estadoAprendiz",
    JORNADA: "jornada", "CALCULO FIN E.P.": "calculoFinEP", ETAPA: "etapa",
    MODALIDAD: "modalidad", Instructor: "instructor",
    "Observaciones de la ficha": "obsFicha", Alternativa: "alternativa",
    "inicio Etapa Productiva": "inicioEP", "Fin Etapa Productiva": "finEP",
    "AUTORIZACIÓN - FECHA": "autorizacion", CERTFICADO: "certificado", "nº Visitas": "visitas",
    "OTROS DOCUMENTOS": "otrosDocs", "Reconocimiento Especial": "reconocimiento",
    "Juicio Evaluación EP": "juicioEP", ACTA: "acta", Deserción: "desercion",
    "Número documento": "documentoSVP", "RAPS aprobados": "rapsOk", "RAPS no aprobados": "rapsNo",
    "RAPS por evaluar": "rapsPend", "Sumatorias RAPS": "rapsSum", "Fecha emisión RAP": "fechaRap",
    "Estado RAP EP": "estadoRap", "Estado CMP EP": "estadoCmp", "Fichas relacionadas": "fichasRel",
    "Fuente actualización": "fuente", "Fecha actualización SVP": "fechaActSVP",
    "Fecha registro alternativa EP": "fechaAlt", "Estado alternativa": "estadoAlt",
  };

  const FICHA_MAP = {
    PROGRAMA: "programa", ESTADO: "estado", JORNADA: "jornada", INICIO: "inicio",
    "FIN LECTIVA": "finLectiva", "Fecha Incio Productiva": "inicioProductiva",
    "Fecha Fin": "fechaFin", "CALCULO FIN E.P. REGLEMENTO": "calculoFinEP", ETAPA: "etapa",
    MODALIDAD: "modalidad", "ESTADO EP": "estadoEP", RESPONSABLE: "responsable",
    P04_NIVEL_FORMACION: "nivel", P04_NOMBRE_EMPRESA: "empresa",
    P04_TOTAL_APRENDICES: "aprendices", P04_TOTAL_APRENDICES_ACTIVOS: "aprendicesActivos",
    FUENTE: "fuente",
  };

  const PLAN_MAP = {
    Especialidad: "especialidad", Ficha: "ficha", Modalidad: "modalidad",
    "Inicio lectiva": "inicioLectiva", "Fin lectiva": "finLectiva",
    "Inicio productiva": "inicioProductiva", "Fin productiva": "finProductiva",
    Responsable: "responsable", OBSERVACIONES: "observaciones",
  };

  function cleanHeader(h, idx) {
    if (h == null || h === "") return `col_${idx}`;
    return String(h).replace(/\n/g, " ").replace(/\s+/g, " ").trim() || `col_${idx}`;
  }

  function cellVal(v) {
    if (v == null || v === "") return null;
    if (v instanceof Date && !isNaN(v)) {
      const y = v.getFullYear();
      const m = String(v.getMonth() + 1).padStart(2, "0");
      const d = String(v.getDate()).padStart(2, "0");
      if (v.getHours() || v.getMinutes() || v.getSeconds()) {
        return `${y}-${m}-${d} ${String(v.getHours()).padStart(2, "0")}:${String(v.getMinutes()).padStart(2, "0")}:${String(v.getSeconds()).padStart(2, "0")}`;
      }
      return `${y}-${m}-${d}`;
    }
    if (typeof v === "number" && Number.isFinite(v)) {
      if (Math.abs(v - Math.round(v)) < 1e-9 && Math.abs(v) < 1e15) return Math.round(v);
      return Math.round(v * 1e6) / 1e6;
    }
    const s = String(v).trim();
    return s || null;
  }

  function readCell(sheet, r, c) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = sheet[addr];
    if (!cell) return null;
    // Preferir texto formateado (evita notación científica en documentos largos)
    if (cell.w != null && String(cell.w).trim() !== "") {
      const w = String(cell.w).trim();
      if (/e\+/i.test(w) && typeof cell.v === "number") {
        // fallback: si el formato ya vino científico, usar entero redondeado en string
        return String(Math.round(cell.v));
      }
      return cellVal(w);
    }
    return cellVal(cell.v);
  }

  function getSheetHeaders(wb, name) {
    const sheet = wb.Sheets[name];
    if (!sheet || !sheet["!ref"]) return [];
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const headerVals = [];
    for (let c = range.s.c; c <= range.e.c; c++) headerVals.push(readCell(sheet, range.s.r, c));
    return headerVals.map((h, i) => cleanHeader(h, i));
  }

  function sheetRows(wb, name) {
    const sheet = wb.Sheets[name];
    if (!sheet || !sheet["!ref"]) return [];
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const headers = getSheetHeaders(wb, name);
    const seen = {};
    const fixed = headers.map((h) => {
      if (seen[h] != null) {
        seen[h] += 1;
        return `${h}_${seen[h]}`;
      }
      seen[h] = 0;
      return h;
    });

    const out = [];
    for (let r = range.s.r + 1; r <= range.e.r; r++) {
      const rec = {};
      let any = false;
      for (let c = range.s.c; c <= range.e.c; c++) {
        const v = readCell(sheet, r, c);
        if (v != null && v !== "") {
          rec[fixed[c - range.s.c]] = v;
          any = true;
        }
      }
      if (any) out.push(rec);
    }
    return out;
  }

  function headerPresent(headers, aliases) {
    const normH = headers.map((h) => normVal(h).replace(/[^a-z0-9]/g, ""));
    return aliases.some((alias) => {
      const a = normVal(alias).replace(/[^a-z0-9]/g, "");
      return normH.some((h) => h === a || h.includes(a) || a.includes(h));
    });
  }

  function validateStructure(wb, sheets, aprendices, fichas, plan, fileName) {
    const errors = [];
    const warnings = [];
    const names = wb.SheetNames || [];

    if (!fileName || !/\.xlsx$/i.test(fileName)) {
      errors.push({
        code: "EXT",
        msg: "El archivo debe ser .xlsx (formato Excel moderno). Si es .xls o CSV, conviértalo primero.",
      });
    }

    if (!sheets.integral) {
      errors.push({
        code: "HOJA_INTEGRAL",
        msg: `No se encontró la hoja obligatoria "Integral". Hojas presentes: ${names.join(", ") || "(ninguna)"}.`,
      });
    }

    if (!sheets.fichas) {
      warnings.push({
        code: "HOJA_FICHAS",
        msg: 'No se encontró hoja de fichas ("Actualizado 2017-2026" o similar). La comparación de fichas quedará incompleta.',
      });
    }

    if (!sheets.plan) {
      warnings.push({
        code: "HOJA_PLAN",
        msg: 'No se encontró "Plan de Trabajo Instructores". No se podrá cruzar instructor por plan.',
      });
    }

    if (sheets.integral) {
      const headers = getSheetHeaders(wb, sheets.integral);
      const required = [
        { label: "Numero de Documento", aliases: ["Numero de Documento", "Número documento", "documento"] },
        { label: "Grupo / Ficha", aliases: ["Grupo", "Ficha", "Ficha Grupo"] },
        { label: "Nombres", aliases: ["Nombres"] },
        { label: "Apellidos", aliases: ["Apellidos"] },
        { label: "PROGRAMA", aliases: ["PROGRAMA"] },
        { label: "ESTADO_APRENDIZ", aliases: ["ESTADO_APRENDIZ", "Estado aprendiz"] },
      ];
      const missing = required.filter((r) => !headerPresent(headers, r.aliases)).map((r) => r.label);
      if (missing.length) {
        errors.push({
          code: "COLUMNAS_INTEGRAL",
          msg: `La hoja "${sheets.integral}" no coincide con el formato esperado. Faltan columnas: ${missing.join(", ")}.`,
          detail: `Encabezados leídos: ${headers.slice(0, 20).join(" | ")}${headers.length > 20 ? "…" : ""}`,
        });
      }

      const recommended = [
        { label: "JORNADA", aliases: ["JORNADA"] },
        { label: "MODALIDAD", aliases: ["MODALIDAD"] },
        { label: "ESTADO_FICHA", aliases: ["ESTADO_FICHA"] },
        { label: "Alternativa", aliases: ["Alternativa"] },
        { label: "Instructor", aliases: ["Instructor"] },
      ];
      const missRec = recommended.filter((r) => !headerPresent(headers, r.aliases)).map((r) => r.label);
      if (missRec.length) {
        warnings.push({
          code: "COLUMNAS_OPCIONALES",
          msg: `Columnas recomendadas ausentes en Integral: ${missRec.join(", ")}.`,
        });
      }
    }

    if (sheets.fichas) {
      const headers = getSheetHeaders(wb, sheets.fichas);
      const need = [
        { label: "Ficha/Grupo", aliases: ["Ficha", "Grupo", "Ficha Grupo"] },
        { label: "PROGRAMA", aliases: ["PROGRAMA"] },
        { label: "ESTADO", aliases: ["ESTADO"] },
      ];
      const missing = need.filter((r) => !headerPresent(headers, r.aliases)).map((r) => r.label);
      if (missing.length) {
        errors.push({
          code: "COLUMNAS_FICHAS",
          msg: `La hoja de fichas "${sheets.fichas}" tiene formato inconsistente. Faltan: ${missing.join(", ")}.`,
        });
      }
    }

    if (sheets.plan) {
      const headers = getSheetHeaders(wb, sheets.plan);
      const need = [
        { label: "Ficha", aliases: ["Ficha"] },
        { label: "Responsable", aliases: ["Responsable"] },
      ];
      const missing = need.filter((r) => !headerPresent(headers, r.aliases)).map((r) => r.label);
      if (missing.length) {
        errors.push({
          code: "COLUMNAS_PLAN",
          msg: `La hoja de plan "${sheets.plan}" no tiene el formato esperado. Faltan: ${missing.join(", ")}.`,
        });
      }
    }

    if (!aprendices.length) {
      errors.push({
        code: "SIN_DATOS",
        msg: "La hoja Integral no trajo registros de aprendices. Revise que no esté vacía o con encabezados en otra fila.",
      });
    } else if (aprendices.length < 10) {
      warnings.push({
        code: "POCOS_REGISTROS",
        msg: `Solo se leyeron ${aprendices.length} aprendices. Si esperaba miles, el formato/hoja probablemente es incorrecto.`,
      });
    }

    // Data inconsistencies
    const keyCount = new Map();
    let sinIdentidad = 0;
    let docCientifico = 0;
    let sinFicha = 0;
    let sinPrograma = 0;
    let sinEstado = 0;

    for (const r of aprendices) {
      const k = aprKey(r);
      keyCount.set(k, (keyCount.get(k) || 0) + 1);
      if (!r.documento && !r.documentoSVP && !r.nombres && !r.ficha) sinIdentidad += 1;
      const docStr = String(r.documento || "");
      if (/e\+/i.test(docStr) || /e\+/i.test(String(r.documentoSVP || ""))) docCientifico += 1;
      if (!r.ficha) sinFicha += 1;
      if (!r.programa) sinPrograma += 1;
      if (!r.estadoAprendiz) sinEstado += 1;
    }

    const dupKeys = [...keyCount.entries()].filter(([, n]) => n > 1);
    if (dupKeys.length) {
      errors.push({
        code: "DUPLICADOS",
        msg: `Hay ${dupKeys.length} llaves duplicadas (documento+ficha). El archivo tiene inconsistencias de identidad.`,
        detail: dupKeys.slice(0, 5).map(([k, n]) => `${k} ×${n}`).join(" · "),
      });
    }

    if (docCientifico > 0) {
      errors.push({
        code: "DOC_CIENTIFICO",
        msg: `${docCientifico} documentos están en notación científica (ej. 8.17E+14). En Excel, formatee la columna como Texto antes de guardar.`,
      });
    }

    if (sinIdentidad > 0) {
      errors.push({
        code: "SIN_IDENTIDAD",
        msg: `${sinIdentidad} filas no tienen documento, nombre ni ficha.`,
      });
    }

    const n = aprendices.length || 1;
    if (sinFicha / n > 0.3) {
      errors.push({
        code: "SIN_FICHA",
        msg: `${sinFicha} aprendices (${Math.round((sinFicha / n) * 100)}%) no tienen ficha. Formato inconsistente o columna Grupo mal nombrada.`,
      });
    } else if (sinFicha > 0) {
      warnings.push({
        code: "SIN_FICHA_PARCIAL",
        msg: `${sinFicha} aprendices sin ficha.`,
      });
    }

    if (sinPrograma / n > 0.4) {
      errors.push({
        code: "SIN_PROGRAMA",
        msg: `${sinPrograma} registros (${Math.round((sinPrograma / n) * 100)}%) sin PROGRAMA. Revise el encabezado de la columna.`,
      });
    }

    if (sinEstado / n > 0.4) {
      errors.push({
        code: "SIN_ESTADO",
        msg: `${sinEstado} registros (${Math.round((sinEstado / n) * 100)}%) sin ESTADO_APRENDIZ.`,
      });
    }

    if (sheets.fichas && !fichas.length) {
      warnings.push({
        code: "FICHAS_VACIAS",
        msg: `La hoja "${sheets.fichas}" existe pero no se leyeron fichas útiles.`,
      });
    }

    if (sheets.plan && !plan.length) {
      warnings.push({
        code: "PLAN_VACIO",
        msg: `La hoja "${sheets.plan}" existe pero no se leyeron asignaciones.`,
      });
    }

    return {
      ok: errors.length === 0,
      errors,
      warnings,
      stats: {
        aprendices: aprendices.length,
        fichas: fichas.length,
        plan: plan.length,
        sinFicha,
        sinPrograma,
        sinEstado,
        duplicados: dupKeys.length,
      },
    };
  }

  function renameIntegral(rec) {
    const out = {};
    for (const [k, v] of Object.entries(rec)) {
      if (k.startsWith("Detalle Bitácoras")) out.detalleBitacoras = v;
      else if (k.startsWith("GFPI-F023")) out.gfpi = v;
      else if (k.startsWith("OBSERVACIONES")) out.observaciones = v;
      else if (k.startsWith("Comité Académico")) out.comite = v;
      else if (k.startsWith("Bitacora ")) {
        const n = k.replace("Bitacora ", "").trim();
        out[`bit${n}`] = v;
      } else if (INTEGRAL_MAP[k]) out[INTEGRAL_MAP[k]] = v;
      else if (!k.startsWith("Unnamed") && !k.startsWith("col_")) {
        out[k.replace(/[^a-zA-Z0-9]+/g, "_").slice(0, 40)] = v;
      }
    }
    out._q = Object.values(out).filter((x) => x != null).join(" ").toLowerCase();
    return out;
  }

  function renameFicha(rec) {
    const out = {};
    for (const [k, v] of Object.entries(rec)) {
      const kk = cleanHeader(k, 0);
      if (/^ficha/i.test(kk) || /grupo/i.test(kk)) out.ficha = out.ficha ?? v;
      else if (FICHA_MAP[kk]) out[FICHA_MAP[kk]] = v;
    }
    out._q = Object.values(out).filter((x) => x != null).join(" ").toLowerCase();
    return out;
  }

  function renamePlan(rec) {
    const out = {};
    for (const [k, v] of Object.entries(rec)) {
      if (PLAN_MAP[k]) out[PLAN_MAP[k]] = v;
    }
    out._q = Object.values(out).filter((x) => x != null).join(" ").toLowerCase();
    return out;
  }

  function normVal(v) {
    if (v == null || v === "") return "";
    if (typeof v === "number" && Number.isFinite(v)) {
      if (Math.abs(v - Math.round(v)) < 1e-9) return String(Math.round(v));
      return String(v);
    }
    let s = String(v).trim().replace(/\s+/g, " ");
    const iso = s.match(/(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
    if (iso) return `${iso[1]}-${iso[2].padStart(2, "0")}-${iso[3].padStart(2, "0")}`;
    const dmy = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})$/);
    if (dmy) {
      const y = dmy[3].length === 2 ? `20${dmy[3]}` : dmy[3];
      return `${y}-${dmy[2].padStart(2, "0")}-${dmy[1].padStart(2, "0")}`;
    }
    return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  }

  function aprKey(r) {
    const doc = String(r.documento || r.documentoSVP || "").replace(/\D/g, "");
    const ficha = String(r.ficha || "").replace(/\D/g, "");
    if (doc) return `d:${doc}|f:${ficha || "-"}`;
    const nom = normVal(`${r.nombres || ""}|${r.apellidos || ""}`);
    return `n:${nom}|f:${ficha || "-"}`;
  }

  function fichaKey(r) {
    return String(r.ficha || "").replace(/\D/g, "") || String(r.ficha || "");
  }

  function planKey(r) {
    return `${fichaKey(r)}|${normVal(r.responsable || "")}`;
  }

  function sig(r) {
    return WATCH.map((k) => `${k}=${normVal(r[k])}`).join("||");
  }

  function uniqueSorted(values) {
    return [...new Set(values.filter((v) => v != null && String(v).trim() !== "").map(String))].sort((a, b) =>
      a.localeCompare(b, "es", { sensitivity: "base" })
    );
  }

  function buildMeta(aprendices, fichas, plan, fileName) {
    return {
      generado: new Date().toISOString().slice(0, 19).replace("T", " "),
      fuente: fileName || "Excel cargado",
      centro: "Centro de Diseño y Metrología",
      regional: "Regional Distrito Capital",
      totales: { aprendices: aprendices.length, fichas: fichas.length, plan: plan.length },
      filtros: {
        estadoAprendiz: uniqueSorted(aprendices.map((r) => r.estadoAprendiz)),
        nivel: uniqueSorted(aprendices.map((r) => r.nivel)),
        jornada: uniqueSorted(aprendices.map((r) => r.jornada)),
        modalidad: uniqueSorted(aprendices.map((r) => r.modalidad)),
        etapa: uniqueSorted(aprendices.map((r) => r.etapa)),
        estadoFicha: uniqueSorted(aprendices.map((r) => r.estadoFicha)),
        alternativa: uniqueSorted(aprendices.map((r) => r.alternativa)),
        instructor: uniqueSorted(aprendices.map((r) => r.instructor)),
        estadoRap: uniqueSorted(aprendices.map((r) => r.estadoRap)),
        estadoCmp: uniqueSorted(aprendices.map((r) => r.estadoCmp)),
        programa: uniqueSorted(aprendices.map((r) => r.programa)),
        fuente: uniqueSorted(aprendices.map((r) => r.fuente)),
      },
    };
  }

  function findSheet(names, candidates) {
    for (const c of candidates) {
      const hit = names.find((n) => n === c || n.toLowerCase() === c.toLowerCase());
      if (hit) return hit;
    }
    // fuzzy contains
    for (const c of candidates) {
      const hit = names.find((n) => n.toLowerCase().includes(c.toLowerCase().slice(0, 12)));
      if (hit) return hit;
    }
    return null;
  }

  function parseWorkbook(wb, fileName) {
    const names = wb.SheetNames || [];
    const integralName = findSheet(names, ["Integral"]);
    const fichasName = findSheet(names, ["Actualizado 2017-2026", "E.P. 2017 - 2026 ACTUALIZADO"]);
    const planName = findSheet(names, ["Plan de Trabajo Instructores"]);
    const sheets = { integral: integralName, fichas: fichasName, plan: planName };

    const aprendices = integralName
      ? sheetRows(wb, integralName).map(renameIntegral).filter((r) => r.nombres || r.documento || r.ficha)
      : [];

    const fichas = fichasName
      ? sheetRows(wb, fichasName).map(renameFicha).filter((r) => r.ficha)
      : [];

    const plan = planName
      ? sheetRows(wb, planName).map(renamePlan).filter((r) => r.ficha || r.especialidad)
      : [];

    const validation = validateStructure(wb, sheets, aprendices, fichas, plan, fileName);

    return {
      meta: buildMeta(aprendices, fichas, plan, fileName),
      aprendices,
      fichas,
      plan,
      sheets,
      validation,
    };
  }

  function compareData(oldData, newData) {
    const oldMap = new Map((oldData.aprendices || []).map((r) => [aprKey(r), r]));
    const newMap = new Map((newData.aprendices || []).map((r) => [aprKey(r), r]));

    const added = [];
    const removed = [];
    const changed = [];

    for (const [k, row] of newMap) {
      if (!oldMap.has(k)) added.push(row);
      else {
        const prev = oldMap.get(k);
        if (sig(prev) !== sig(row)) {
          const fields = WATCH.filter((f) => normVal(prev[f]) !== normVal(row[f])).map((f) => ({
            field: f,
            before: prev[f] ?? null,
            after: row[f] ?? null,
          }));
          if (fields.length) changed.push({ key: k, row, fields });
        }
      }
    }
    for (const [k, row] of oldMap) {
      if (!newMap.has(k)) removed.push(row);
    }

    const oldF = new Set((oldData.fichas || []).map(fichaKey).filter(Boolean));
    const newF = new Set((newData.fichas || []).map(fichaKey).filter(Boolean));
    const fichasAdded = [...newF].filter((f) => !oldF.has(f));
    const fichasRemoved = [...oldF].filter((f) => !newF.has(f));

    const oldP = new Set((oldData.plan || []).map(planKey));
    const newP = new Set((newData.plan || []).map(planKey));
    const planAdded = [...newP].filter((p) => !oldP.has(p)).length;
    const planRemoved = [...oldP].filter((p) => !newP.has(p)).length;

    return {
      before: {
        aprendices: (oldData.aprendices || []).length,
        fichas: (oldData.fichas || []).length,
        plan: (oldData.plan || []).length,
      },
      after: {
        aprendices: (newData.aprendices || []).length,
        fichas: (newData.fichas || []).length,
        plan: (newData.plan || []).length,
      },
      added,
      removed,
      changed,
      fichasAdded,
      fichasRemoved,
      planAdded,
      planRemoved,
      deltaAprendices: (newData.aprendices || []).length - (oldData.aprendices || []).length,
    };
  }

  async function readFile(file) {
    if (!file) throw new Error("No se recibió archivo.");
    const name = file.name || "";
    if (!/\.xlsx$/i.test(name) && !/\.xls$/i.test(name)) {
      const validation = {
        ok: false,
        errors: [{ code: "EXT", msg: `Formato no soportado (${name || "sin nombre"}). Use un archivo .xlsx unificado de Etapa Productiva.` }],
        warnings: [],
        stats: {},
      };
      return {
        meta: buildMeta([], [], [], name),
        aprendices: [],
        fichas: [],
        plan: [],
        sheets: {},
        validation,
      };
    }
    const buf = await file.arrayBuffer();
    let wb;
    try {
      wb = XLSX.read(buf, { type: "array", cellDates: true });
    } catch (e) {
      throw new Error(`No se pudo abrir el Excel. Archivo dañado o no es un libro válido. (${e.message || e})`);
    }
    if (!wb.SheetNames || !wb.SheetNames.length) {
      throw new Error("El archivo no contiene hojas. No es un Excel válido para esta consulta.");
    }
    return parseWorkbook(wb, file.name);
  }

  const DB_NAME = "sena-ep-dashboard";
  const STORE = "baseline";

  function openDb() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE);
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function saveBaseline(payload) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, "readwrite");
      tx.objectStore(STORE).put(payload, "latest");
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => reject(tx.error);
    });
  }

  async function loadBaseline() {
    try {
      const db = await openDb();
      return new Promise((resolve) => {
        const tx = db.transaction(STORE, "readonly");
        const req = tx.objectStore(STORE).get("latest");
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => resolve(null);
      });
    } catch {
      return null;
    }
  }

  window.ExcelImport = {
    readFile,
    compareData,
    aprKey,
    saveBaseline,
    loadBaseline,
    validateStructure,
    WATCH,
  };
})();
