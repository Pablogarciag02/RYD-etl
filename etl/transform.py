"""Transform layer: populate core dimensions and fact tables from raw data."""


def populate_dimensions(cur):
    """Extract distinct values from raw tables into core dimension tables."""
    print("  Populating dim_group...")
    cur.execute("""
        INSERT INTO core.dim_group (group_name)
        SELECT DISTINCT val FROM (
            SELECT grupo AS val FROM raw.raw_finance_applications WHERE grupo IS NOT NULL
            UNION SELECT group_name FROM raw.raw_sales WHERE group_name IS NOT NULL
            UNION SELECT grupo FROM raw.raw_claims WHERE grupo IS NOT NULL
        ) t
        ON CONFLICT (group_name) DO NOTHING
    """)
    print(f"    dim_group: {cur.rowcount} inserted")

    print("  Populating dim_brand...")
    cur.execute("""
        INSERT INTO core.dim_brand (brand_name)
        SELECT DISTINCT val FROM (
            SELECT marca AS val FROM raw.raw_finance_applications WHERE marca IS NOT NULL
            UNION SELECT marca FROM raw.raw_market_prices WHERE marca IS NOT NULL
            UNION SELECT brand FROM raw.raw_sales WHERE brand IS NOT NULL
            UNION SELECT marca FROM raw.raw_inegi_sales WHERE marca IS NOT NULL
        ) t
        ON CONFLICT (brand_name) DO NOTHING
    """)
    print(f"    dim_brand: {cur.rowcount} inserted")

    print("  Populating dim_oem...")
    cur.execute("""
        INSERT INTO core.dim_oem (oem_name)
        SELECT DISTINCT val FROM (
            SELECT oem AS val FROM raw.raw_market_prices WHERE oem IS NOT NULL
            UNION SELECT oem FROM raw.raw_sales WHERE oem IS NOT NULL
            UNION SELECT oem FROM raw.raw_inegi_sales WHERE oem IS NOT NULL
        ) t
        ON CONFLICT (oem_name) DO NOTHING
    """)
    print(f"    dim_oem: {cur.rowcount} inserted")

    # dim_source_channel and dim_campaign were only ever populated from raw_leads /
    # raw_lead_activities. Now that leads no longer flow through this pipeline, those
    # dimensions are frozen at whatever values were loaded historically. Existing
    # fact_leads rows can still join to them; new ingestion does not extend them.

    print("  Populating dim_insurer...")
    cur.execute("""
        INSERT INTO core.dim_insurer (insurer_name)
        SELECT DISTINCT aseguradora FROM raw.raw_claims
        WHERE aseguradora IS NOT NULL
        ON CONFLICT (insurer_name) DO NOTHING
    """)
    print(f"    dim_insurer: {cur.rowcount} inserted")

    print("  Populating dim_product...")
    cur.execute("""
        INSERT INTO core.dim_product (brand_id, oem_id, product_name)
        SELECT DISTINCT
            b.id,
            NULL::uuid,
            t.producto
        FROM (
            SELECT marca, producto FROM raw.raw_finance_applications WHERE producto IS NOT NULL
        ) t
        LEFT JOIN core.dim_brand b ON b.brand_name = t.marca
        ON CONFLICT (brand_id, oem_id, product_name) DO NOTHING
    """)
    print(f"    dim_product: {cur.rowcount} inserted")

    # For sales, carline is the product — map it with brand/oem
    cur.execute("""
        INSERT INTO core.dim_product (brand_id, oem_id, product_name)
        SELECT DISTINCT
            b.id,
            o.id,
            s.carline
        FROM raw.raw_sales s
        LEFT JOIN core.dim_brand b ON b.brand_name = s.brand
        LEFT JOIN core.dim_oem o ON o.oem_name = s.oem
        WHERE s.carline IS NOT NULL
        ON CONFLICT (brand_id, oem_id, product_name) DO NOTHING
    """)
    print(f"    dim_product (sales carlines): {cur.rowcount} inserted")

    print("  Populating dim_model_version...")
    cur.execute("""
        INSERT INTO core.dim_model_version
            (brand_id, oem_id, model_name, version_name, body_type, doors, model_year, ryd_segment)
        SELECT DISTINCT
            b.id,
            NULL::uuid,
            p.modelo,
            p.version,
            p.tipo_carroceria,
            CASE WHEN p.puertas_raw ~ '^[0-9]+$' THEN p.puertas_raw::integer ELSE NULL END,
            CASE WHEN p.anio_modelo_raw ~ '^[0-9]+$' THEN p.anio_modelo_raw::integer ELSE NULL END,
            p.segmento_ryd
        FROM raw.raw_market_prices p
        LEFT JOIN core.dim_brand b ON b.brand_name = p.marca
        WHERE p.modelo IS NOT NULL AND p.version IS NOT NULL
        ON CONFLICT (brand_id, oem_id, model_name, version_name, model_year) DO NOTHING
    """)
    print(f"    dim_model_version: {cur.rowcount} inserted")

    print("  Populating dim_dealer...")
    # Dealer catalog is managed manually (authoritative) and must not be mutated by uploads.
    # Therefore, this ETL does not INSERT/UPDATE `core.dim_dealer`.
    print("    dim_dealer: skipped (manual DB-managed dimension)")

    # dim_advisor was only fed from raw_leads / raw_lead_activities. Frozen for the same reason.

    print("  Dimensions populated.")


def populate_fact_finance_applications(cur):
    """raw.raw_finance_applications → core.fact_finance_applications."""
    print("  Loading fact_finance_applications...")
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
            r.folio::bigint,
            d.id,
            g.id,
            b.id,
            p.id,
            CASE WHEN r.fecha_hora_recibida_raw IS NOT NULL
                 THEN to_timestamp(r.fecha_hora_recibida_raw, 'DD/MM/YY HH24:MI:SS')
                 ELSE NULL END,
            CASE WHEN r.fecha_ok_raw IS NOT NULL
                 THEN r.fecha_ok_raw::date ELSE NULL END,
            CASE WHEN r.periodo_yyyymm_raw ~ '^[0-9]+$'
                 THEN r.periodo_yyyymm_raw::integer ELSE NULL END,
            r.tipo_vehiculo,
            r.estatus,
            r.subestatus,
            r.tipo_persona,
            CASE WHEN r.plazo_meses_raw ~ '^[0-9]+$'
                 THEN r.plazo_meses_raw::integer ELSE NULL END,
            CASE WHEN r.anio_modelo_raw ~ '^[0-9]+$'
                 THEN r.anio_modelo_raw::integer ELSE NULL END,
            r.modelo,
            r.carline,
            r.version,
            CASE
                WHEN r.monto_unidad_raw IS NULL THEN NULL
                ELSE NULLIF(regexp_replace(r.monto_unidad_raw, '[^0-9\\.-]', '', 'g'), '')::numeric(14,2)
            END,
            CASE
                WHEN r.enganche_raw IS NULL THEN NULL
                ELSE NULLIF(regexp_replace(r.enganche_raw, '[^0-9\\.-]', '', 'g'), '')::numeric(14,2)
            END,
            CASE
                WHEN r.porcentaje_enganche_raw IS NULL THEN NULL
                ELSE NULLIF(regexp_replace(r.porcentaje_enganche_raw, '[^0-9\\.-]', '', 'g'), '')::numeric(8,4)
            END,
            CASE
                WHEN r.monto_financiado_raw IS NULL THEN NULL
                ELSE NULLIF(regexp_replace(r.monto_financiado_raw, '[^0-9\\.-]', '', 'g'), '')::numeric(14,2)
            END,
            CASE WHEN r.aprobado_raw = '1' THEN true
                 WHEN r.aprobado_raw = '0' THEN false ELSE NULL END,
            CASE WHEN r.base_flag_raw = '1' THEN true
                 WHEN r.base_flag_raw = '0' THEN false ELSE NULL END,
            CASE WHEN r.anf_flag_raw = '1' THEN true
                 WHEN r.anf_flag_raw = '0' THEN false ELSE NULL END,
            r.import_batch_id
        FROM raw.raw_finance_applications r
        LEFT JOIN core.dim_dealer d ON d.commercial_name = r.distribuidor_ok
        LEFT JOIN core.dim_group g ON g.group_name = r.grupo
        LEFT JOIN core.dim_brand b ON b.brand_name = r.marca
        LEFT JOIN core.dim_product p ON p.product_name = r.producto AND p.brand_id = b.id
        ON CONFLICT (folio) DO NOTHING
    """)
    print(f"    fact_finance_applications: {cur.rowcount} inserted")


def populate_fact_market_prices(cur):
    """raw.raw_market_prices → core.fact_market_prices."""
    print("  Loading fact_market_prices...")
    cur.execute("""
        INSERT INTO core.fact_market_prices
            (month_date, brand_id, oem_id, model_version_id,
             retail_price, cash_net_price, finance_price, currency,
             source_import_batch_id)
        SELECT
            r.mes_raw::date,
            b.id,
            NULL::uuid,
            mv.id,
            CASE
                WHEN r.precio_lista_raw IS NULL THEN NULL
                ELSE NULLIF(regexp_replace(r.precio_lista_raw, '[^0-9\\.-]', '', 'g'), '')::numeric(14,2)
            END,
            CASE
                WHEN r.precio_contado_neto_raw IS NULL THEN NULL
                ELSE NULLIF(regexp_replace(r.precio_contado_neto_raw, '[^0-9\\.-]', '', 'g'), '')::numeric(14,2)
            END,
            CASE
                WHEN r.precio_financiado_raw IS NULL THEN NULL
                ELSE NULLIF(regexp_replace(r.precio_financiado_raw, '[^0-9\\.-]', '', 'g'), '')::numeric(14,2)
            END,
            r.moneda,
            r.import_batch_id
        FROM raw.raw_market_prices r
        LEFT JOIN core.dim_brand b ON b.brand_name = r.marca
        LEFT JOIN core.dim_model_version mv
            ON mv.model_name = r.modelo
            AND mv.version_name = r.version
            AND mv.brand_id = b.id
            AND (mv.model_year = CASE WHEN r.anio_modelo_raw ~ '^[0-9]+$' THEN r.anio_modelo_raw::integer ELSE NULL END
                 OR mv.model_year IS NULL)
        WHERE mv.id IS NOT NULL
        ON CONFLICT (month_date, model_version_id) DO NOTHING
    """)
    print(f"    fact_market_prices: {cur.rowcount} inserted")


def populate_fact_sales(cur):
    """raw.raw_sales → core.fact_sales.

    Idempotent across re-uploads, including for sales whose product/dealer/etc.
    don't match a dim row (product_id ends up NULL). The plain ON CONFLICT
    unique constraint cannot dedupe rows with NULL in the key — Postgres treats
    NULLs as distinct — so we additionally:
      1. DISTINCT ON the business key inside this batch (handles intra-batch dupes)
      2. WHERE NOT EXISTS … IS NOT DISTINCT FROM … against the existing fact
         table (handles re-uploads that overlap previous loads).
    The ON CONFLICT stays as a final safety net for the non-NULL fast path.
    """
    print("  Loading fact_sales...")
    cur.execute("""
        INSERT INTO core.fact_sales
            (month_date, dealer_id, group_id, brand_id, oem_id, product_id,
             sale_type, units,
             source_import_batch_id)
        SELECT DISTINCT ON (s.month_date, s.dealer_id, s.product_id, s.sale_type)
            s.month_date, s.dealer_id, s.group_id, s.brand_id, s.oem_id, s.product_id,
            s.sale_type, s.units, s.import_batch_id
        FROM (
            SELECT
                r.month_raw::date AS month_date,
                d.id              AS dealer_id,
                g.id              AS group_id,
                b.id              AS brand_id,
                o.id              AS oem_id,
                p.id              AS product_id,
                r.sale_type       AS sale_type,
                CASE WHEN r.units_raw ~ '^-?[0-9]+$' THEN r.units_raw::integer ELSE 0 END AS units,
                r.import_batch_id AS import_batch_id
            FROM raw.raw_sales r
            LEFT JOIN core.dim_dealer  d ON d.commercial_name = r.dealer
            LEFT JOIN core.dim_group   g ON g.group_name      = r.group_name
            LEFT JOIN core.dim_brand   b ON b.brand_name      = r.brand
            LEFT JOIN core.dim_oem     o ON o.oem_name        = r.oem
            LEFT JOIN core.dim_product p ON p.product_name = r.carline
                                        AND p.brand_id     = b.id
                                        AND p.oem_id       = o.id
        ) s
        WHERE NOT EXISTS (
            SELECT 1 FROM core.fact_sales f
            WHERE f.month_date IS NOT DISTINCT FROM s.month_date
              AND f.dealer_id  IS NOT DISTINCT FROM s.dealer_id
              AND f.product_id IS NOT DISTINCT FROM s.product_id
              AND f.sale_type  IS NOT DISTINCT FROM s.sale_type
        )
        ORDER BY s.month_date, s.dealer_id, s.product_id, s.sale_type
        ON CONFLICT (month_date, dealer_id, product_id, sale_type) DO NOTHING
    """)
    print(f"    fact_sales: {cur.rowcount} inserted")


def populate_fact_claims(cur):
    """raw.raw_claims → core.fact_claims."""
    print("  Loading fact_claims...")
    # No unique key → full refresh
    cur.execute("TRUNCATE core.fact_claims RESTART IDENTITY")
    cur.execute("""
        INSERT INTO core.fact_claims
            (claim_date, policy_number, insurer_id,
             dealer_id, group_id,
             client_name, claim_number, coverage, driver_name, phone,
             vehicle_description, vin, model_name, city, state, program_name,
             source_import_batch_id)
        SELECT
            CASE WHEN r.fecha_raw IS NOT NULL
                 THEN r.fecha_raw::date ELSE NULL END,
            r.poliza,
            ins.id,
            d.id,
            g.id,
            r.cliente,
            r.numero_siniestro,
            r.cobertura,
            r.conductor,
            r.telefono,
            r.descripcion_vehiculo,
            r.vin,
            r.modelo,
            r.ciudad,
            r.estado,
            r.programa,
            r.import_batch_id
        FROM raw.raw_claims r
        LEFT JOIN core.dim_insurer ins ON ins.insurer_name = r.aseguradora
        LEFT JOIN core.dim_dealer d ON d.commercial_name = COALESCE(r.nombre_agencia, r.nombre_agencia_2)
        LEFT JOIN core.dim_group g ON g.group_name = r.grupo
    """)
    print(f"    fact_claims: {cur.rowcount} inserted")


def populate_fact_market_sales_inegi(cur):
    """raw.raw_inegi_sales → core.fact_market_sales_inegi."""
    print("  Loading fact_market_sales_inegi...")
    # No unique key → full refresh
    cur.execute("TRUNCATE core.fact_market_sales_inegi RESTART IDENTITY")
    cur.execute("""
        INSERT INTO core.fact_market_sales_inegi
            (year_num, month_name, brand_id, oem_id,
             model_name, vehicle_type, segment, origin_type, origin_country,
             quantity, topic,
             source_import_batch_id)
        SELECT
            r.anio_raw::integer,
            r.mes,
            b.id,
            NULL::uuid,
            r.modelo,
            r.tipo_vehiculo,
            r.segmento,
            r.tipo_origen,
            r.pais_origen,
            CASE WHEN r.cantidad_raw ~ '^-?[0-9]+$'
                 THEN r.cantidad_raw::integer ELSE 0 END,
            r.tema,
            r.import_batch_id
        FROM raw.raw_inegi_sales r
        LEFT JOIN core.dim_brand b ON b.brand_name = r.marca
        WHERE r.anio_raw ~ '^[0-9]+$'
          AND r.modelo IS NOT NULL
          AND r.mes IS NOT NULL
    """)
    print(f"    fact_market_sales_inegi: {cur.rowcount} inserted")


def run_transforms(cur):
    """Run all transformations in order: dimensions first, then facts."""
    print("\n=== TRANSFORM: Populating dimensions ===")
    populate_dimensions(cur)

    print("\n=== TRANSFORM: Populating fact tables ===")
    populate_fact_finance_applications(cur)
    populate_fact_market_prices(cur)
    populate_fact_sales(cur)
    populate_fact_claims(cur)
    populate_fact_market_sales_inegi(cur)
    print("\nTransform complete.")
