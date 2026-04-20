"""RYD — single-file Streamlit app for daily per-table Excel uploads.

User flow:
    1. Pick which table is in the file (radio)
    2. Upload an .xlsx
    3. App extracts → loads raw → refreshes dims → refreshes that table's facts
    4. Success message → "Upload another" resets the form

Run:
    source .venv/bin/activate
    streamlit run app.py
"""

import hashlib
import hmac
import json
import os
from datetime import datetime, timedelta

import psycopg2
import streamlit as st
from dotenv import load_dotenv
from openpyxl import load_workbook
from psycopg2.extras import execute_values

load_dotenv()


# ══════════════════════════════════════════════════════════════
# Secret access — reads from Streamlit Cloud secrets OR .env
# ══════════════════════════════════════════════════════════════

def _secret(key, default=None):
    """Prefer st.secrets (Streamlit Cloud). Fall back to env (local .env)."""
    try:
        if key in st.secrets:
            return st.secrets[key]
    except (FileNotFoundError, Exception):
        pass
    return os.getenv(key, default)


# ══════════════════════════════════════════════════════════════
# DB connection
# ══════════════════════════════════════════════════════════════

def get_connection():
    return psycopg2.connect(
        host=_secret("DB_HOST"),
        port=int(_secret("DB_PORT", 5432)),
        dbname=_secret("DB_NAME"),
        user=_secret("DB_USER"),
        password=_secret("DB_PASSWORD"),
        sslmode="require",
    )


# ══════════════════════════════════════════════════════════════
# Value coercion helpers
# ══════════════════════════════════════════════════════════════

def _str(val):
    if val is None:
        return None
    return str(val).strip() or None


def _excel_serial_to_datetime(serial):
    if serial is None:
        return None
    try:
        return datetime(1899, 12, 30) + timedelta(days=float(serial))
    except (ValueError, TypeError):
        return None


def _dt_str(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        val = _excel_serial_to_datetime(val)
    if isinstance(val, datetime):
        return val.isoformat()
    return str(val).strip() or None


def _num_str(val):
    if val is None:
        return None
    return str(val)


# ══════════════════════════════════════════════════════════════
# Extractors — each takes a worksheet and yields dicts
# ══════════════════════════════════════════════════════════════

def extract_leads(ws):
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or not row[0]:
            continue
        yield {
            "source_row_number": row_num,
            "aglead": _str(row[0]),
            "distribuidor": _str(row[1]),
            "fecha_origen_raw": _dt_str(row[2]),
            "hora_origen_raw": _num_str(row[3]),
            "grupo": _str(row[17]),
            "marca": _str(row[16]),
            "producto": _str(row[5]),
            "asesor": _str(row[7]),
            "campania": _str(row[10]),
            "subcampania": _str(row[11]),
            "fuente": _str(row[8]),
            "medio_atencion": _str(row[14]),
            "titulo_lead": _str(row[15]),
            "temperatura": _str(row[6]),
            "estatus": _str(row[9]),
            "motivo_finalizacion": _str(row[13]),
            "raw_payload": json.dumps(
                {"mes_lead": row[4], "tiempo_respuesta": row[12]}, default=str
            ) if (row[4] is not None or row[12] is not None) else None,
        }


def extract_lead_activities(ws):
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or not row[0]:
            continue
        fecha_act = row[2]
        hora_act = None
        if isinstance(fecha_act, datetime):
            hora_act = str(fecha_act.hour)
        fecha_act_raw = _dt_str(fecha_act)

        fecha_plan = row[9]
        hora_plan = None
        if isinstance(fecha_plan, (int, float)):
            dt = _excel_serial_to_datetime(fecha_plan)
            if dt:
                hora_plan = str(dt.hour)
                fecha_plan = dt
        fecha_plan_raw = _dt_str(fecha_plan)

        yield {
            "source_row_number": row_num,
            "aglead": _str(row[0]),
            "fecha_actividad_raw": fecha_act_raw,
            "hora_actividad_raw": hora_act,
            "fecha_programada_raw": fecha_plan_raw,
            "hora_programada_raw": hora_plan,
            "distribuidor": _str(row[1]),
            "grupo": _str(row[12]),
            "marca": _str(row[11]),
            "producto": _str(row[3]),
            "asesor": _str(row[5]),
            "campania": _str(row[7]),
            "subcampania": None,
            "fuente": _str(row[6]),
            "actividad": _str(row[8]),
            "estatus_actividad": _str(row[10]),
            "temperatura": _str(row[4]),
            "raw_payload": None,
        }


def extract_finance_applications(ws):
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or row[5] is None:
            continue
        yield {
            "source_row_number": row_num,
            "folio": _str(row[5]),
            "codigo_distribuidor": _str(row[18]),
            "distribuidor_ok": _str(row[4]),
            "nombre_distribuidor": _str(row[19]),
            "grupo": _str(row[20]),
            "marca": _str(row[21]),
            "oem": None,
            "producto": _str(row[24]),
            "carline": _str(row[12]),
            # Column 11 header is "Modelo" but values are model years (e.g. 2024, 2026)
            "modelo": None,
            "version": _str(row[13]),
            "anio_modelo_raw": _num_str(row[11]),
            "fecha_hora_recibida_raw": _str(row[7]),
            "fecha_ok_raw": _dt_str(row[0]),
            "periodo_yyyymm_raw": _num_str(row[6]),
            "tipo_vehiculo": _str(row[8]),
            "estatus": _str(row[9]),
            "subestatus": _str(row[10]),
            "tipo_persona": _str(row[22]),
            "plazo_meses_raw": _num_str(row[23]),
            "enganche_raw": _num_str(row[15]),
            "porcentaje_enganche_raw": _num_str(row[16]),
            "monto_unidad_raw": _num_str(row[14]),
            "monto_financiado_raw": _num_str(row[17]),
            "aprobado_raw": _num_str(row[1]),
            "base_flag_raw": _num_str(row[2]),
            "anf_flag_raw": _num_str(row[3]),
            "raw_payload": None,
        }


def extract_market_prices(ws):
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or row[3] is None:
            continue
        yield {
            "source_row_number": row_num,
            "mes_raw": _dt_str(row[0]),
            "marca": _str(row[2]),
            "oem": None,
            "modelo": _str(row[3]),
            "version": _str(row[4]),
            "anio_modelo_raw": _num_str(row[7]),
            "segmento_ryd": _str(row[1]),
            "tipo_carroceria": _str(row[5]),
            "puertas_raw": _num_str(row[6]),
            "precio_lista_raw": _num_str(row[8]),
            "precio_contado_neto_raw": _num_str(row[9]),
            "precio_financiado_raw": _num_str(row[10]),
            "moneda": _str(row[11]),
            "raw_payload": None,
        }


def extract_sales(ws):
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or row[4] is None:
            continue
        yield {
            "source_row_number": row_num,
            "month_raw": _dt_str(row[0]),
            "dealer": _str(row[2]),
            "group_name": _str(row[1]),
            "brand": "MG",
            "oem": "MG MOTOR",
            "carline": _str(row[5]),
            "sale_type": _str(row[3]),
            "units_raw": _num_str(row[4]),
            "raw_payload": None,
        }


def extract_claims(ws):
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or row[9] is None:
            continue
        razon_social = row[5]
        yield {
            "source_row_number": row_num,
            "fecha_raw": _dt_str(row[0]),
            "poliza": _str(row[1]),
            "aseguradora": _str(row[2]),
            "programa": _str(row[3]),
            "nombre_agencia": _str(row[4]),
            "nombre_agencia_2": _str(row[6]),
            "grupo": _str(row[7]),
            "cliente": _str(row[8]),
            "numero_siniestro": _str(row[9]),
            "cobertura": _str(row[10]),
            "conductor": _str(row[11]),
            "telefono": _str(row[12]),
            "descripcion_vehiculo": _str(row[13]),
            "vin": _str(row[14]),
            "modelo": _str(row[15]),
            "ciudad": _str(row[16]),
            "estado": _str(row[17]),
            "raw_payload": json.dumps({"razon_social": razon_social}, default=str)
                if razon_social is not None else None,
        }


def extract_inegi_sales(ws):
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or row[3] is None:
            continue
        yield {
            "source_row_number": row_num,
            "tema": _str(row[0]),
            "anio_raw": _num_str(row[1]),
            "mes": _str(row[2]),
            "marca": _str(row[3]),
            "oem": None,
            "modelo": _str(row[4]),
            "tipo_vehiculo": _str(row[5]),
            "segmento": _str(row[6]),
            "tipo_origen": _str(row[7]),
            "pais_origen": _str(row[8]),
            "cantidad_raw": _num_str(row[9]),
            "raw_payload": None,
        }


# ══════════════════════════════════════════════════════════════
# Raw loading
# ══════════════════════════════════════════════════════════════

def create_import_batch(cur, filename, template, checksum):
    cur.execute(
        """
        INSERT INTO ingest.import_batches
            (source_file_name, source_sheet_name, source_system, template_type,
             uploaded_by, status, checksum)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
        RETURNING id
        """,
        (filename, template, "streamlit_upload", template, "streamlit_app",
         "processing", checksum),
    )
    return cur.fetchone()[0]


def update_batch_status(cur, batch_id, rows_detected, rows_loaded, rows_failed=0):
    status = "success" if rows_failed == 0 else "partial_success"
    if rows_loaded == 0:
        status = "failed"
    cur.execute(
        """
        UPDATE ingest.import_batches
        SET rows_detected = %s, rows_loaded = %s, rows_failed = %s,
            status = %s, updated_at = now()
        WHERE id = %s
        """,
        (rows_detected, rows_loaded, rows_failed, status, batch_id),
    )


def bulk_insert(cur, table_name, rows_iter, batch_id, batch_size=1000):
    buffered = []
    columns = None
    total_loaded = 0
    total_failed = 0

    def flush():
        nonlocal total_loaded, total_failed
        if not buffered:
            return
        cols_sql = ", ".join(columns)
        template = "(" + ", ".join(["%s"] * len(columns)) + ")"
        try:
            execute_values(
                cur,
                f"INSERT INTO {table_name} ({cols_sql}) VALUES %s",
                buffered,
                template=template,
                page_size=len(buffered),
            )
            total_loaded += len(buffered)
        except Exception:
            total_failed += len(buffered)
            raise

    for row in rows_iter:
        row["import_batch_id"] = str(batch_id)
        if columns is None:
            columns = list(row.keys())
        buffered.append(tuple(row[c] for c in columns))
        if len(buffered) >= batch_size:
            flush()
            buffered = []
    flush()
    return total_loaded, total_failed


# ══════════════════════════════════════════════════════════════
# Dimension population (idempotent, always full-scan)
# ══════════════════════════════════════════════════════════════

def populate_dimensions(cur):
    cur.execute("""
        INSERT INTO core.dim_group (group_name)
        SELECT DISTINCT val FROM (
            SELECT grupo AS val FROM raw.raw_leads WHERE grupo IS NOT NULL
            UNION SELECT grupo FROM raw.raw_lead_activities WHERE grupo IS NOT NULL
            UNION SELECT grupo FROM raw.raw_finance_applications WHERE grupo IS NOT NULL
            UNION SELECT group_name FROM raw.raw_sales WHERE group_name IS NOT NULL
            UNION SELECT grupo FROM raw.raw_claims WHERE grupo IS NOT NULL
            UNION SELECT grupo FROM raw.raw_dealer_catalog WHERE grupo IS NOT NULL
        ) t
        ON CONFLICT (group_name) DO NOTHING
    """)
    cur.execute("""
        INSERT INTO core.dim_brand (brand_name)
        SELECT DISTINCT val FROM (
            SELECT marca AS val FROM raw.raw_leads WHERE marca IS NOT NULL
            UNION SELECT marca FROM raw.raw_lead_activities WHERE marca IS NOT NULL
            UNION SELECT marca FROM raw.raw_finance_applications WHERE marca IS NOT NULL
            UNION SELECT marca FROM raw.raw_market_prices WHERE marca IS NOT NULL
            UNION SELECT brand FROM raw.raw_sales WHERE brand IS NOT NULL
            UNION SELECT marca FROM raw.raw_dealer_catalog WHERE marca IS NOT NULL
            UNION SELECT marca FROM raw.raw_inegi_sales WHERE marca IS NOT NULL
        ) t
        ON CONFLICT (brand_name) DO NOTHING
    """)
    cur.execute("""
        INSERT INTO core.dim_oem (oem_name)
        SELECT DISTINCT val FROM (
            SELECT oem AS val FROM raw.raw_market_prices WHERE oem IS NOT NULL
            UNION SELECT oem FROM raw.raw_sales WHERE oem IS NOT NULL
            UNION SELECT oem FROM raw.raw_dealer_catalog WHERE oem IS NOT NULL
            UNION SELECT oem FROM raw.raw_inegi_sales WHERE oem IS NOT NULL
        ) t
        ON CONFLICT (oem_name) DO NOTHING
    """)
    cur.execute("""
        INSERT INTO core.dim_source_channel (source_name)
        SELECT DISTINCT val FROM (
            SELECT fuente AS val FROM raw.raw_leads WHERE fuente IS NOT NULL
            UNION SELECT fuente FROM raw.raw_lead_activities WHERE fuente IS NOT NULL
        ) t
        ON CONFLICT (source_name) DO NOTHING
    """)
    cur.execute("""
        INSERT INTO core.dim_campaign (campaign_name, subcampaign_name)
        SELECT DISTINCT campania, subcampania FROM (
            SELECT campania, subcampania FROM raw.raw_leads WHERE campania IS NOT NULL
            UNION SELECT campania, subcampania FROM raw.raw_lead_activities WHERE campania IS NOT NULL
        ) t
        ON CONFLICT (campaign_name, subcampaign_name) DO NOTHING
    """)
    cur.execute("""
        INSERT INTO core.dim_insurer (insurer_name)
        SELECT DISTINCT aseguradora FROM raw.raw_claims WHERE aseguradora IS NOT NULL
        ON CONFLICT (insurer_name) DO NOTHING
    """)
    cur.execute("""
        INSERT INTO core.dim_product (brand_id, oem_id, product_name)
        SELECT DISTINCT b.id, NULL::uuid, t.producto FROM (
            SELECT marca, producto FROM raw.raw_leads WHERE producto IS NOT NULL
            UNION SELECT marca, producto FROM raw.raw_lead_activities WHERE producto IS NOT NULL
            UNION SELECT marca, producto FROM raw.raw_finance_applications WHERE producto IS NOT NULL
        ) t
        LEFT JOIN core.dim_brand b ON b.brand_name = t.marca
        ON CONFLICT (brand_id, oem_id, product_name) DO NOTHING
    """)
    cur.execute("""
        INSERT INTO core.dim_product (brand_id, oem_id, product_name)
        SELECT DISTINCT b.id, o.id, s.carline
        FROM raw.raw_sales s
        LEFT JOIN core.dim_brand b ON b.brand_name = s.brand
        LEFT JOIN core.dim_oem o ON o.oem_name = s.oem
        WHERE s.carline IS NOT NULL
        ON CONFLICT (brand_id, oem_id, product_name) DO NOTHING
    """)
    cur.execute("""
        INSERT INTO core.dim_model_version
            (brand_id, oem_id, model_name, version_name, body_type, doors, model_year, ryd_segment)
        SELECT DISTINCT
            b.id, NULL::uuid, p.modelo, p.version, p.tipo_carroceria,
            CASE WHEN p.puertas_raw ~ '^[0-9]+$' THEN p.puertas_raw::integer ELSE NULL END,
            CASE WHEN p.anio_modelo_raw ~ '^[0-9]+$' THEN p.anio_modelo_raw::integer ELSE NULL END,
            p.segmento_ryd
        FROM raw.raw_market_prices p
        LEFT JOIN core.dim_brand b ON b.brand_name = p.marca
        WHERE p.modelo IS NOT NULL AND p.version IS NOT NULL
        ON CONFLICT (brand_id, oem_id, model_name, version_name, model_year) DO NOTHING
    """)
    cur.execute("""
        INSERT INTO core.dim_dealer
            (dealer_code, commercial_name, legal_name, rfc, group_id, brand_id, oem_id,
             classification, dms, crm, workshop_status, start_sales_date)
        SELECT DISTINCT ON (dc.codigo_distribuidor)
            dc.codigo_distribuidor, dc.nombre_comercial, dc.razon_social, dc.rfc,
            g.id, b.id, o.id, dc.clasificacion, dc.dms, dc.crm, dc.estatus_taller,
            CASE WHEN dc.fecha_inicio_ventas_raw IS NOT NULL
                 THEN dc.fecha_inicio_ventas_raw::date ELSE NULL END
        FROM raw.raw_dealer_catalog dc
        LEFT JOIN core.dim_group g ON g.group_name = dc.grupo
        LEFT JOIN core.dim_brand b ON b.brand_name = dc.marca
        LEFT JOIN core.dim_oem o ON o.oem_name = dc.oem
        WHERE dc.nombre_comercial IS NOT NULL
        ON CONFLICT (dealer_code) DO NOTHING
    """)
    cur.execute("""
        INSERT INTO core.dim_dealer (commercial_name, group_id, brand_id)
        SELECT DISTINCT t.distribuidor, g.id, b.id FROM (
            SELECT distribuidor, grupo, marca FROM raw.raw_leads WHERE distribuidor IS NOT NULL
            UNION SELECT distribuidor, grupo, marca FROM raw.raw_lead_activities WHERE distribuidor IS NOT NULL
            UNION SELECT distribuidor_ok, grupo, marca FROM raw.raw_finance_applications WHERE distribuidor_ok IS NOT NULL
            UNION SELECT dealer, group_name, brand FROM raw.raw_sales WHERE dealer IS NOT NULL
        ) t
        LEFT JOIN core.dim_group g ON g.group_name = t.grupo
        LEFT JOIN core.dim_brand b ON b.brand_name = t.marca
        WHERE NOT EXISTS (SELECT 1 FROM core.dim_dealer d WHERE d.commercial_name = t.distribuidor)
        ON CONFLICT DO NOTHING
    """)
    cur.execute("""
        INSERT INTO core.dim_advisor (advisor_name, dealer_id)
        SELECT DISTINCT t.asesor, d.id FROM (
            SELECT asesor, distribuidor FROM raw.raw_leads WHERE asesor IS NOT NULL
            UNION SELECT asesor, distribuidor FROM raw.raw_lead_activities WHERE asesor IS NOT NULL
        ) t
        LEFT JOIN core.dim_dealer d ON d.commercial_name = t.distribuidor
        ON CONFLICT (advisor_name, dealer_id) DO NOTHING
    """)


# ══════════════════════════════════════════════════════════════
# Per-table fact transforms — each rebuilds its fact table
# from the FULL raw history. Tables with unique keys use
# ON CONFLICT; tables without use DELETE+INSERT.
# ══════════════════════════════════════════════════════════════

def fact_leads(cur):
    cur.execute("""
        INSERT INTO core.fact_leads
            (aglead, dealer_id, group_id, brand_id, product_id, advisor_id,
             campaign_id, source_channel_id,
             lead_origin_ts, lead_origin_date, lead_origin_hour, lead_month_num,
             temperature, status, finalization_reason, attention_channel, lead_title,
             source_import_batch_id)
        SELECT
            r.aglead, d.id, g.id, b.id, p.id, adv.id, camp.id, sc.id,
            CASE WHEN r.fecha_origen_raw IS NOT NULL THEN r.fecha_origen_raw::timestamptz ELSE NULL END,
            CASE WHEN r.fecha_origen_raw IS NOT NULL THEN r.fecha_origen_raw::date ELSE NULL END,
            CASE WHEN r.hora_origen_raw ~ '^[0-9]+$' THEN r.hora_origen_raw::integer ELSE NULL END,
            EXTRACT(MONTH FROM r.fecha_origen_raw::date),
            r.temperatura, r.estatus, r.motivo_finalizacion, r.medio_atencion, r.titulo_lead,
            r.import_batch_id
        FROM raw.raw_leads r
        LEFT JOIN core.dim_dealer d ON d.commercial_name = r.distribuidor
        LEFT JOIN core.dim_group g ON g.group_name = r.grupo
        LEFT JOIN core.dim_brand b ON b.brand_name = r.marca
        LEFT JOIN core.dim_product p ON p.product_name = r.producto AND p.brand_id = b.id
        LEFT JOIN core.dim_advisor adv ON adv.advisor_name = r.asesor AND adv.dealer_id = d.id
        LEFT JOIN core.dim_campaign camp ON camp.campaign_name = r.campania
            AND (camp.subcampaign_name = r.subcampania OR (camp.subcampaign_name IS NULL AND r.subcampania IS NULL))
        LEFT JOIN core.dim_source_channel sc ON sc.source_name = r.fuente
        ON CONFLICT (aglead) DO NOTHING
    """)
    return cur.rowcount


def fact_lead_activities(cur):
    # No unique key → full refresh
    cur.execute("TRUNCATE core.fact_lead_activities RESTART IDENTITY")
    cur.execute("""
        INSERT INTO core.fact_lead_activities
            (aglead, dealer_id, group_id, brand_id, product_id, advisor_id,
             campaign_id, source_channel_id,
             activity_ts, planned_ts, activity_name, activity_status, temperature,
             source_import_batch_id)
        SELECT
            r.aglead, d.id, g.id, b.id, p.id, adv.id, camp.id, sc.id,
            CASE WHEN r.fecha_actividad_raw IS NOT NULL THEN r.fecha_actividad_raw::timestamptz ELSE NULL END,
            CASE WHEN r.fecha_programada_raw IS NOT NULL THEN r.fecha_programada_raw::timestamptz ELSE NULL END,
            r.actividad, r.estatus_actividad, r.temperatura, r.import_batch_id
        FROM raw.raw_lead_activities r
        INNER JOIN core.fact_leads fl ON fl.aglead = r.aglead
        LEFT JOIN core.dim_dealer d ON d.commercial_name = r.distribuidor
        LEFT JOIN core.dim_group g ON g.group_name = r.grupo
        LEFT JOIN core.dim_brand b ON b.brand_name = r.marca
        LEFT JOIN core.dim_product p ON p.product_name = r.producto AND p.brand_id = b.id
        LEFT JOIN core.dim_advisor adv ON adv.advisor_name = r.asesor AND adv.dealer_id = d.id
        LEFT JOIN core.dim_campaign camp ON camp.campaign_name = r.campania
            AND (camp.subcampaign_name = r.subcampania OR (camp.subcampaign_name IS NULL AND r.subcampania IS NULL))
        LEFT JOIN core.dim_source_channel sc ON sc.source_name = r.fuente
    """)
    return cur.rowcount


def fact_finance_applications(cur):
    cur.execute("""
        INSERT INTO core.fact_finance_applications
            (folio, dealer_id, group_id, brand_id, product_id,
             received_ts, ok_date, request_period_yyyymm,
             vehicle_type, status, substatus, person_type, term_months,
             model_year, model_name, carline, version_name,
             unit_amount, down_payment, down_payment_pct, financed_amount,
             approved_flag, base_flag, anf_flag,
             source_import_batch_id)
        SELECT
            r.folio::bigint, d.id, g.id, b.id, p.id,
            CASE WHEN r.fecha_hora_recibida_raw IS NOT NULL
                 THEN to_timestamp(r.fecha_hora_recibida_raw, 'DD/MM/YY HH24:MI:SS') ELSE NULL END,
            CASE WHEN r.fecha_ok_raw IS NOT NULL THEN r.fecha_ok_raw::date ELSE NULL END,
            CASE WHEN r.periodo_yyyymm_raw ~ '^[0-9]+$' THEN r.periodo_yyyymm_raw::integer ELSE NULL END,
            r.tipo_vehiculo, r.estatus, r.subestatus, r.tipo_persona,
            CASE WHEN r.plazo_meses_raw ~ '^[0-9]+$' THEN r.plazo_meses_raw::integer ELSE NULL END,
            CASE WHEN r.anio_modelo_raw ~ '^[0-9]+$' THEN r.anio_modelo_raw::integer ELSE NULL END,
            NULL::text, r.carline, r.version,
            CASE WHEN r.monto_unidad_raw ~ '^[[0-9].]+$' THEN r.monto_unidad_raw::numeric(14,2) ELSE NULL END,
            CASE WHEN r.enganche_raw ~ '^[[0-9].]+$' THEN r.enganche_raw::numeric(14,2) ELSE NULL END,
            CASE WHEN r.porcentaje_enganche_raw ~ '^[[0-9].]+$' THEN r.porcentaje_enganche_raw::numeric(8,4) ELSE NULL END,
            CASE WHEN r.monto_financiado_raw ~ '^[[0-9].]+$' THEN r.monto_financiado_raw::numeric(14,2) ELSE NULL END,
            CASE WHEN r.aprobado_raw = '1' THEN true WHEN r.aprobado_raw = '0' THEN false ELSE NULL END,
            CASE WHEN r.base_flag_raw = '1' THEN true WHEN r.base_flag_raw = '0' THEN false ELSE NULL END,
            CASE WHEN r.anf_flag_raw = '1' THEN true WHEN r.anf_flag_raw = '0' THEN false ELSE NULL END,
            r.import_batch_id
        FROM raw.raw_finance_applications r
        LEFT JOIN core.dim_dealer d ON d.commercial_name = r.distribuidor_ok
        LEFT JOIN core.dim_group g ON g.group_name = r.grupo
        LEFT JOIN core.dim_brand b ON b.brand_name = r.marca
        LEFT JOIN core.dim_product p ON p.product_name = r.producto AND p.brand_id = b.id
        ON CONFLICT (folio) DO NOTHING
    """)
    return cur.rowcount


def fact_market_prices(cur):
    cur.execute("""
        INSERT INTO core.fact_market_prices
            (month_date, brand_id, oem_id, model_version_id,
             retail_price, cash_net_price, finance_price, currency,
             source_import_batch_id)
        SELECT
            r.mes_raw::date, b.id, NULL::uuid, mv.id,
            CASE WHEN r.precio_lista_raw ~ '^[[0-9].]+$' THEN r.precio_lista_raw::numeric(14,2) ELSE NULL END,
            CASE WHEN r.precio_contado_neto_raw ~ '^[[0-9].]+$' THEN r.precio_contado_neto_raw::numeric(14,2) ELSE NULL END,
            CASE WHEN r.precio_financiado_raw ~ '^[[0-9].]+$' THEN r.precio_financiado_raw::numeric(14,2) ELSE NULL END,
            r.moneda, r.import_batch_id
        FROM raw.raw_market_prices r
        LEFT JOIN core.dim_brand b ON b.brand_name = r.marca
        LEFT JOIN core.dim_model_version mv
            ON mv.model_name = r.modelo AND mv.version_name = r.version AND mv.brand_id = b.id
            AND (mv.model_year = CASE WHEN r.anio_modelo_raw ~ '^[0-9]+$'
                                      THEN r.anio_modelo_raw::integer ELSE NULL END
                 OR mv.model_year IS NULL)
        WHERE mv.id IS NOT NULL
        ON CONFLICT (month_date, model_version_id) DO NOTHING
    """)
    return cur.rowcount


def fact_sales(cur):
    cur.execute("""
        INSERT INTO core.fact_sales
            (month_date, dealer_id, group_id, brand_id, oem_id, product_id,
             sale_type, units, source_import_batch_id)
        SELECT
            r.month_raw::date, d.id, g.id, b.id, o.id, p.id,
            r.sale_type,
            CASE WHEN r.units_raw ~ '^-?[0-9]+$' THEN r.units_raw::integer ELSE 0 END,
            r.import_batch_id
        FROM raw.raw_sales r
        LEFT JOIN core.dim_dealer d ON d.commercial_name = r.dealer
        LEFT JOIN core.dim_group g ON g.group_name = r.group_name
        LEFT JOIN core.dim_brand b ON b.brand_name = r.brand
        LEFT JOIN core.dim_oem o ON o.oem_name = r.oem
        LEFT JOIN core.dim_product p ON p.product_name = r.carline AND p.brand_id = b.id AND p.oem_id = o.id
        ON CONFLICT (month_date, dealer_id, product_id, sale_type) DO NOTHING
    """)
    return cur.rowcount


def fact_claims(cur):
    cur.execute("TRUNCATE core.fact_claims RESTART IDENTITY")
    cur.execute("""
        INSERT INTO core.fact_claims
            (claim_date, policy_number, insurer_id, dealer_id, group_id,
             client_name, claim_number, coverage, driver_name, phone,
             vehicle_description, vin, model_name, city, state, program_name,
             source_import_batch_id)
        SELECT
            CASE WHEN r.fecha_raw IS NOT NULL THEN r.fecha_raw::date ELSE NULL END,
            r.poliza, ins.id, d.id, g.id,
            r.cliente, r.numero_siniestro, r.cobertura, r.conductor, r.telefono,
            r.descripcion_vehiculo, r.vin, r.modelo, r.ciudad, r.estado, r.programa,
            r.import_batch_id
        FROM raw.raw_claims r
        LEFT JOIN core.dim_insurer ins ON ins.insurer_name = r.aseguradora
        LEFT JOIN core.dim_dealer d ON d.commercial_name = COALESCE(r.nombre_agencia, r.nombre_agencia_2)
        LEFT JOIN core.dim_group g ON g.group_name = r.grupo
    """)
    return cur.rowcount


def fact_market_sales_inegi(cur):
    cur.execute("TRUNCATE core.fact_market_sales_inegi RESTART IDENTITY")
    cur.execute("""
        INSERT INTO core.fact_market_sales_inegi
            (year_num, month_name, brand_id, oem_id,
             model_name, vehicle_type, segment, origin_type, origin_country,
             quantity, topic, source_import_batch_id)
        SELECT
            r.anio_raw::integer, r.mes, b.id, NULL::uuid,
            r.modelo, r.tipo_vehiculo, r.segmento, r.tipo_origen, r.pais_origen,
            CASE WHEN r.cantidad_raw ~ '^-?[0-9]+$' THEN r.cantidad_raw::integer ELSE 0 END,
            r.tema, r.import_batch_id
        FROM raw.raw_inegi_sales r
        LEFT JOIN core.dim_brand b ON b.brand_name = r.marca
        WHERE r.anio_raw ~ '^[0-9]+$' AND r.modelo IS NOT NULL AND r.mes IS NOT NULL
    """)
    return cur.rowcount


# ══════════════════════════════════════════════════════════════
# Expected headers per table (case-insensitive, order-sensitive).
# These guard against uploading a file with the wrong radio selection.
# ══════════════════════════════════════════════════════════════

EXPECTED_HEADERS = {
    "leads": ["AgLead", "Distribuidor", "Fecha Origen", "Hora Origen", "MES Lead",
              "Producto", "Temperatura", "Asesor", "Fuente", "Estatus",
              "Campaña", "Subcampaña", "Tiempo Respuesta", "Motivo de Finalización",
              "Medio Atencion", "TituloLead", "Marca", "Grupo"],
    "activities": ["AgLead", "Distribuidor", "Fecha Actividad", "Producto",
                   "Temperatura", "Asesor", "Fuente", "Campaña", "Actividad",
                   "Fecha Planeada", "Estatus", "Marca", "Grupo"],
    "finance": ["Fecha OK", "Cuenta Aprobado", "Cuenta Base", "Cuenta ANF",
                "Distribuidor_OK", "Folio", "Mes Solicitud", "Fecha y Hora Recibida",
                "Tipo_Auto", "Status", "Substatus", "Modelo", "Carline", "Versión",
                "Importe Unidad", "Enganche", "Porcentaje de Enganche",
                "Monto a Financiar", "Codigo Distribuidor", "Nombre Distribuidor",
                "Grupo", "Marca", "Tipo de Persona", "Plazo", "Producto"],
    "prices": ["Mes", "Segmento RYD", "Marca", "Modelo", "Version", "Body type",
               "Puertas", "Model year", "Retail price", "Cash/Net price",
               "Finance price", "Price currency"],
    "sales": ["Month", "Group", "Dealer", "Tipo de Venta", "Unidades", "Carline"],
    "claims": ["Fecha", "Póliza", "Aseguradora", "Programa", "Nombre Agencia",
               "Razón Social", "Nombre de Agencia 2", "Grupo", "Cliente",
               "Siniestro", "Cobertura", "Conductor", "Teléfono", "Vehículo",
               "Serie", "Modelo", "Ciudad", "Estado"],
    "inegi": ["Tema", "Año", "Mes", "Marca", "Modelo", "Tipo", "Segmento",
              "Origen", "País origen", "Cantidad"],
}


def _read_headers(ws):
    first_row = next(ws.iter_rows(max_row=1, values_only=True), None)
    if not first_row:
        return []
    return [str(h).strip() if h is not None else "" for h in first_row]


def validate_headers(ws, table_code):
    """Return None if headers match; otherwise a user-facing error string."""
    expected = EXPECTED_HEADERS[table_code]
    found = _read_headers(ws)

    if len(found) < len(expected):
        return (f"This file has {len(found)} columns but '{TABLES[table_code]['label']}' "
                f"expects at least {len(expected)}. Did you pick the right table?")

    exp_norm = [h.strip().lower() for h in expected]
    got_norm = [h.strip().lower() for h in found[:len(expected)]]
    if exp_norm != got_norm:
        mismatches = [
            f"  col {i+1}: expected '{expected[i]}', got '{found[i]}'"
            for i in range(len(expected)) if exp_norm[i] != got_norm[i]
        ]
        return ("The uploaded file's headers don't match the selected table.\n\n"
                + "\n".join(mismatches[:5])
                + ("\n  …" if len(mismatches) > 5 else ""))
    return None


# ══════════════════════════════════════════════════════════════
# Table registry — one entry per user-selectable table
# ══════════════════════════════════════════════════════════════

TABLES = {
    "leads": {
        "label": "Leads (BaseLeads1)",
        "raw_table": "raw.raw_leads",
        "template": "BaseLeads1",
        "extractor": extract_leads,
        "fact_fn": fact_leads,
    },
    "activities": {
        "label": "Lead Activities (BaseLeads2)",
        "raw_table": "raw.raw_lead_activities",
        "template": "BaseLeads2",
        "extractor": extract_lead_activities,
        "fact_fn": fact_lead_activities,
    },
    "finance": {
        "label": "Finance Applications (DBSolicitudes_Cetelem)",
        "raw_table": "raw.raw_finance_applications",
        "template": "DBSolicitudes_Cetelem",
        "extractor": extract_finance_applications,
        "fact_fn": fact_finance_applications,
    },
    "prices": {
        "label": "Market Prices (DBPreciosMexico_ConMG)",
        "raw_table": "raw.raw_market_prices",
        "template": "DBPreciosMexico_ConMG",
        "extractor": extract_market_prices,
        "fact_fn": fact_market_prices,
    },
    "sales": {
        "label": "Sales (DBVentas_ConMG)",
        "raw_table": "raw.raw_sales",
        "template": "DBVentas_ConMG",
        "extractor": extract_sales,
        "fact_fn": fact_sales,
    },
    "claims": {
        "label": "Claims (DBSiniestros_Marsh)",
        "raw_table": "raw.raw_claims",
        "template": "DBSiniestros_Marsh",
        "extractor": extract_claims,
        "fact_fn": fact_claims,
    },
    "inegi": {
        "label": "INEGI Sales (BaseINEGIAutosLigerosMexico)",
        "raw_table": "raw.raw_inegi_sales",
        "template": "BaseINEGIAutosLigerosMexico",
        "extractor": extract_inegi_sales,
        "fact_fn": fact_market_sales_inegi,
    },
}


# ══════════════════════════════════════════════════════════════
# ETL entry point
# ══════════════════════════════════════════════════════════════

def run_upload(file_obj, filename, table_code):
    """Process one uploaded file for one table. Returns a stats dict."""
    cfg = TABLES[table_code]

    file_bytes = file_obj.read()
    checksum = hashlib.md5(file_bytes).hexdigest()
    from io import BytesIO
    wb = load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.worksheets[0]

    # Guard: refuse the upload if the file's header row doesn't match
    # the selected table. Prevents silent misrouting of data.
    header_error = validate_headers(ws, table_code)
    if header_error:
        wb.close()
        raise ValueError(header_error)

    conn = get_connection()
    conn.autocommit = False
    try:
        cur = conn.cursor()
        batch_id = create_import_batch(cur, filename, cfg["template"], checksum)
        loaded, failed = bulk_insert(cur, cfg["raw_table"], cfg["extractor"](ws), batch_id)
        update_batch_status(cur, batch_id, loaded + failed, loaded, failed)

        populate_dimensions(cur)
        facts_inserted = cfg["fact_fn"](cur)

        conn.commit()
        return {
            "batch_id": str(batch_id),
            "raw_loaded": loaded,
            "raw_failed": failed,
            "facts_inserted": facts_inserted,
        }
    except Exception:
        conn.rollback()
        raise
    finally:
        wb.close()
        conn.close()


# ══════════════════════════════════════════════════════════════
# Password gate
# ══════════════════════════════════════════════════════════════

def check_password():
    """Block the app behind a shared password. Call before rendering UI."""
    expected = _secret("APP_PASSWORD")
    if not expected:
        st.error("APP_PASSWORD is not configured. Set it in Streamlit secrets "
                 "(or a local .env file) before running the app.")
        st.stop()

    def _on_submit():
        entered = st.session_state.get("pw_input", "") or ""
        # Compare as bytes so non-ASCII (e.g. smart-quote autocorrect) returns
        # False instead of raising TypeError.
        if hmac.compare_digest(entered.encode("utf-8"), str(expected).encode("utf-8")):
            st.session_state["auth_ok"] = True
            del st.session_state["pw_input"]
        else:
            st.session_state["auth_ok"] = False

    if st.session_state.get("auth_ok"):
        return

    st.text_input("Password", type="password", on_change=_on_submit, key="pw_input")
    if "auth_ok" in st.session_state and not st.session_state["auth_ok"]:
        st.error("Incorrect password.")
    st.stop()


# ══════════════════════════════════════════════════════════════
# Streamlit UI
# ══════════════════════════════════════════════════════════════

st.set_page_config(page_title="RYD — Data Upload", layout="centered")
check_password()
st.title("RYD — Data Upload")
st.caption("Upload one Excel file per table, one at a time.")

if "history" not in st.session_state:
    st.session_state.history = []
if "form_nonce" not in st.session_state:
    st.session_state.form_nonce = 0

label_to_code = {cfg["label"]: code for code, cfg in TABLES.items()}

with st.form(key=f"upload_form_{st.session_state.form_nonce}", clear_on_submit=True):
    table_label = st.radio(
        "Which table is in this file?",
        options=list(label_to_code.keys()),
        index=0,
    )
    uploaded = st.file_uploader("Excel file (.xlsx)", type=["xlsx"])
    submit = st.form_submit_button("Process upload", type="primary")

if submit:
    if uploaded is None:
        st.error("Please choose a file before submitting.")
    else:
        table_code = label_to_code[table_label]
        with st.spinner(f"Loading {table_label}…"):
            try:
                stats = run_upload(uploaded, uploaded.name, table_code)
                st.session_state.history.append({"label": table_label, **stats})
                st.success(
                    f"**Success!** `{table_label}` uploaded.\n\n"
                    f"- Raw rows loaded: **{stats['raw_loaded']:,}**\n"
                    f"- Fact rows populated: **{stats['facts_inserted']:,}**\n"
                    f"- Batch ID: `{stats['batch_id']}`"
                )
            except Exception as e:
                st.error(f"Upload failed: {e}")

if st.button("Upload another file"):
    st.session_state.form_nonce += 1
    st.rerun()

if st.session_state.history:
    with st.expander(f"This session's uploads ({len(st.session_state.history)})", expanded=False):
        for u in reversed(st.session_state.history):
            st.text(
                f"✓ {u['label']}  —  "
                f"{u['raw_loaded']:,} raw, {u['facts_inserted']:,} facts  "
                f"(batch {u['batch_id'][:8]})"
            )
