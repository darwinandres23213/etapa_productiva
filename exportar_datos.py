# -*- coding: utf-8 -*-
"""Regenera data.json desde el Excel unificado de Etapa Productiva."""
import json
import math
import re
from datetime import datetime, date, time
from pathlib import Path

HERE = Path(__file__).resolve().parent
SRC = HERE / "Etapa_Productiva_MG_Integral_Unificada_SVP_Tecnico_Tecnologo (1).xlsx"
OUT = HERE / "data.json"

import openpyxl


def clean_header(h, idx):
    if h is None:
        return f"col_{idx}"
    s = re.sub(r"\s+", " ", str(h).replace("\n", " ").strip())
    return s or f"col_{idx}"


def to_jsonable(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S") if (v.hour or v.minute or v.second) else v.strftime("%Y-%m-%d")
    if isinstance(v, date) and not isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return None
        if v == int(v) and abs(v) < 1e15:
            return int(v)
        return round(v, 6)
    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return v
    if isinstance(v, str):
        t = v.strip()
        return t or None
    return str(v).strip() or None


def sheet_to_records(ws):
    rows_iter = ws.iter_rows(values_only=True)
    raw_headers = next(rows_iter)
    headers = [clean_header(h, i) for i, h in enumerate(raw_headers)]
    seen, fixed = {}, []
    for h in headers:
        if h in seen:
            seen[h] += 1
            fixed.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            fixed.append(h)
    headers = fixed
    records = []
    for row in rows_iter:
        vals = [to_jsonable(c) for c in row]
        if not any(v is not None and v != "" for v in vals):
            continue
        if len(vals) < len(headers):
            vals.extend([None] * (len(headers) - len(vals)))
        vals = vals[: len(headers)]
        rec = {headers[j]: vals[j] for j in range(len(headers)) if vals[j] is not None}
        if rec:
            records.append(rec)
    return records


def rename_integral(rec):
    mapping = {
        "Numero de Documento": "documento", "Grupo": "ficha", "Teléfono": "telefono",
        "Correo electónico": "correo", "Nombres": "nombres", "Apellidos": "apellidos",
        "NIT": "nit", "Razon Social": "empresa", "Departamento empresa": "deptoEmpresa",
        "Ciudad empresa": "ciudadEmpresa", "Dirección": "direccionEmpresa",
        "Teléfono empresa": "telEmpresa", "Correo electrónico": "correoEmpresa",
        "Tipo documento": "tipoDoc", "Centro ID": "centroId", "Centro": "centro",
        "Especialidad": "especialidad", "INICIO ETAPA LECTIVA": "inicioLectiva",
        "Fin lectiva": "finLectiva", "Inicio productiva": "inicioProductiva",
        "Fecha Fin": "fechaFin", "Contrato inicio": "contratoInicio", "Contrato Fin": "contratoFin",
        "Regional": "regional", "Fase": "fase", "ESTADO_FICHA": "estadoFicha",
        "CODIGO_PROGRAMA": "codPrograma", "VERSION_PROGRANA": "versionPrograma",
        "PROGRAMA": "programa", "NIVEL_DE_FORMACION": "nivel", "ESTADO_APRENDIZ": "estadoAprendiz",
        "JORNADA": "jornada", "CALCULO FIN E.P.": "calculoFinEP", "ETAPA": "etapa",
        "MODALIDAD": "modalidad", "Instructor": "instructor",
        "Observaciones de la ficha": "obsFicha", "Alternativa": "alternativa",
        "inicio Etapa Productiva": "inicioEP", "Fin Etapa Productiva": "finEP",
        "Bitacora 1": "bit1", "Bitacora 2": "bit2", "Bitacora 3": "bit3", "Bitacora 4": "bit4",
        "Bitacora 5": "bit5", "Bitacora 6": "bit6", "Bitacora 7": "bit7", "Bitacora 8": "bit8",
        "Bitacora 9": "bit9", "Bitacora 10": "bit10", "Bitacora 11": "bit11", "Bitacora 12": "bit12",
        "AUTORIZACIÓN - FECHA": "autorizacion", "CERTFICADO": "certificado", "nº Visitas": "visitas",
        "OTROS DOCUMENTOS": "otrosDocs", "Reconocimiento Especial": "reconocimiento",
        "Juicio Evaluación EP": "juicioEP", "ACTA": "acta", "Deserción": "desercion",
        "Número documento": "documentoSVP", "RAPS aprobados": "rapsOk", "RAPS no aprobados": "rapsNo",
        "RAPS por evaluar": "rapsPend", "Sumatorias RAPS": "rapsSum", "Fecha emisión RAP": "fechaRap",
        "Estado RAP EP": "estadoRap", "Estado CMP EP": "estadoCmp", "Fichas relacionadas": "fichasRel",
        "Fuente actualización": "fuente", "Fecha actualización SVP": "fechaActSVP",
        "Fecha registro alternativa EP": "fechaAlt", "Estado alternativa": "estadoAlt",
    }
    out = {}
    for k, v in rec.items():
        if k.startswith("Detalle Bitácoras"):
            out["detalleBitacoras"] = v
        elif k.startswith("GFPI-F023"):
            out["gfpi"] = v
        elif k.startswith("OBSERVACIONES"):
            out["observaciones"] = v
        elif k.startswith("Comité Académico"):
            out["comite"] = v
        elif k in mapping:
            out[mapping[k]] = v
        elif not k.startswith("Unnamed") and not k.startswith("col_"):
            key = re.sub(r"[^a-zA-Z0-9]+", "_", k).strip("_")[:40]
            if key:
                out[key] = v
    out["_q"] = " ".join(str(v) for v in out.values() if v is not None).lower()
    return out


def rename_ficha(rec):
    mapping = {
        "PROGRAMA": "programa", "ESTADO": "estado", "JORNADA": "jornada", "INICIO": "inicio",
        "FIN LECTIVA": "finLectiva", "Fecha Incio Productiva": "inicioProductiva",
        "Fecha Fin": "fechaFin", "CALCULO FIN E.P. REGLEMENTO": "calculoFinEP", "ETAPA": "etapa",
        "MODALIDAD": "modalidad", "ESTADO EP": "estadoEP", "RESPONSABLE": "responsable",
        "P04_NIVEL_FORMACION": "nivel", "P04_NOMBRE_EMPRESA": "empresa",
        "P04_TOTAL_APRENDICES": "aprendices", "P04_TOTAL_APRENDICES_ACTIVOS": "aprendicesActivos",
        "P04_NOMBRE_SECTOR_PROGRAMA": "sector", "P04_NOMBRE_NUEVO_SECTOR": "nuevoSector",
        "P04_NOMBRE_OCUPACION": "ocupacion", "P04_CODIGO_PROGRAMA": "codPrograma",
        "P04_ETAPA_FICHA": "etapaP04", "P04_ESTADO_CURSO": "estadoP04",
        "P04_MODALIDAD_FORMACION": "modalidadP04", "P04_Fecha Vencimiento": "vencimiento",
        "FUENTE": "fuente", "P04_TOTAL_APRENDICES_MASCULINOS": "masc",
        "P04_TOTAL_APRENDICES_FEMENINOS": "fem", "P04_DURACION_PROGRAMA": "duracion",
        "P04_NOMBRE_PROGRAMA_FORMACION": "programaP04", "P04_NOMBRE_RESPONSABLE": "responsableP04",
    }
    out = {}
    for k, v in rec.items():
        kk = re.sub(r"\s+", " ", k.replace("\n", " ")).strip()
        if kk.lower().startswith("ficha") or "grupo" in kk.lower():
            out.setdefault("ficha", v)
        elif kk in mapping:
            out[mapping[kk]] = v
    out["_q"] = " ".join(str(v) for v in out.values() if v is not None).lower()
    return out


def rename_plan(rec):
    mapping = {
        "Especialidad": "especialidad", "Ficha": "ficha", "Modalidad": "modalidad",
        "Inicio lectiva": "inicioLectiva", "Fin lectiva": "finLectiva",
        "Inicio productiva": "inicioProductiva", "Fin productiva": "finProductiva",
        "Responsable": "responsable", "OBSERVACIONES": "observaciones",
    }
    out = {}
    for k, v in rec.items():
        if k in mapping:
            out[mapping[k]] = v
        elif k.startswith("Unnamed") and v:
            out.setdefault("nota", v)
    out["_q"] = " ".join(str(v) for v in out.values() if v is not None).lower()
    return out


def unique_sorted(values):
    return sorted({str(v).strip() for v in values if v is not None and str(v).strip()}, key=str.lower)


def main():
    if not SRC.exists():
        raise SystemExit(f"No se encontró el Excel: {SRC}")
    print("Cargando", SRC.name)
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    aprendices = [rename_integral(r) for r in sheet_to_records(wb["Integral"])]
    aprendices = [r for r in aprendices if r.get("nombres") or r.get("documento") or r.get("ficha")]
    fichas = [rename_ficha(r) for r in sheet_to_records(wb["Actualizado 2017-2026"])]
    plan = [rename_plan(r) for r in sheet_to_records(wb["Plan de Trabajo Instructores"]) if r.get("Ficha") or r.get("Especialidad")]
    control = sheet_to_records(wb["Control_Unificacion"])
    wb.close()

    meta = {
        "generado": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "fuente": SRC.name,
        "centro": "Centro de Diseño y Metrología",
        "regional": "Regional Distrito Capital",
        "totales": {"aprendices": len(aprendices), "fichas": len(fichas), "plan": len(plan)},
        "filtros": {
            "estadoAprendiz": unique_sorted(r.get("estadoAprendiz") for r in aprendices),
            "nivel": unique_sorted(r.get("nivel") for r in aprendices),
            "jornada": unique_sorted(r.get("jornada") for r in aprendices),
            "modalidad": unique_sorted(r.get("modalidad") for r in aprendices),
            "etapa": unique_sorted(r.get("etapa") for r in aprendices),
            "estadoFicha": unique_sorted(r.get("estadoFicha") for r in aprendices),
            "alternativa": unique_sorted(r.get("alternativa") for r in aprendices),
            "instructor": unique_sorted(r.get("instructor") for r in aprendices),
            "estadoRap": unique_sorted(r.get("estadoRap") for r in aprendices),
            "estadoCmp": unique_sorted(r.get("estadoCmp") for r in aprendices),
            "programa": unique_sorted(r.get("programa") for r in aprendices),
            "fuente": unique_sorted(r.get("fuente") for r in aprendices),
        },
        "control": control,
    }
    with OUT.open("w", encoding="utf-8") as f:
        json.dump({"meta": meta, "aprendices": aprendices, "fichas": fichas, "plan": plan}, f, ensure_ascii=False, separators=(",", ":"))
    print(f"OK -> {OUT} ({OUT.stat().st_size/1024/1024:.2f} MB)")


if __name__ == "__main__":
    main()
