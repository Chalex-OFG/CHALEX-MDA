import streamlit as st
import pandas as pd

# =========================================================
# CONFIG_MDA ONLINE - GOOGLE SHEETS (GLOBAL)
# =========================================================
CONFIG_MDA_SHEET_ID = "1Lvsb31aFps3FaJjURH_pqGQHqpsotmEU505L32OeF2E"

@st.cache_data(ttl=60, show_spinner=False)
def read_config_sheet(sheet_name):
    """
    Lee una pestaña pública de CONFIG_MDA_ONLINE.
    Disponible para todos los módulos de la aplicación.
    """
    from urllib.parse import quote

    url = (
        f"https://docs.google.com/spreadsheets/d/{CONFIG_MDA_SHEET_ID}/"
        f"gviz/tq?tqx=out:csv&sheet={quote(sheet_name)}"
    )
    return pd.read_csv(url).dropna(axis=1, how="all")


st.set_page_config(page_title="CHALEX-MDA", page_icon="🐺", layout="wide")
try:
    st.image("chalex_network.png", width=1000)
except Exception:
    pass

with st.sidebar:
    st.header("📂 Carga única de datos")
    shared_energy = st.file_uploader("1. Site Energy Dashboard", type=["xlsx", "xls"], key="shared_energy")
    shared_alarms = st.file_uploader("2. Current Alarms", type=["xlsx", "xls"], key="shared_alarms")
    shared_wos = st.file_uploader("3. WOs List", type=["xlsx", "xls"], key="shared_wos")
    st.divider()
    modulo = st.radio("Módulo", ["📊 Corte / Monitoreo", "🔄 Cambio de Turno"])
    if st.button("🔄 Actualizar CONFIG_MDA", use_container_width=True):
        read_config_sheet.clear()
        st.rerun()

if modulo == "📊 Corte / Monitoreo":
    st.header("📊 Corte / Monitoreo")

    import re
    import json
    import pandas as pd
    import streamlit as st



    # =========================================================
    # UTILIDADES
    # =========================================================

    def normalize(text):
        text = str(text).strip().lower()
        text = (
            text.replace("á", "a").replace("é", "e").replace("í", "i")
                .replace("ó", "o").replace("ú", "u").replace("ñ", "n")
        )
        return re.sub(r"[^a-z0-9]+", " ", text).strip()

    def clean_text(value):
        if pd.isna(value):
            return ""
        text = str(value).strip()
        if text.lower() in {"nan", "none", "<na>"}:
            return ""
        return text

    def site_key(value):
        """
        Cruza Energy Dashboard y Current Alarms usando lo que viene
        después del código numérico inicial.
        Ej.: 0131305_IC_San_Antonio_Ica -> ic san antonio ica
        """
        text = clean_text(value)
        if not text:
            return ""
        suffix = re.sub(r"^\s*[0-9]+_?", "", text).strip()
        return normalize(suffix if suffix else text)

    def region_code(value):
        text = clean_text(value)
        if not text:
            return ""
        m = re.match(r"^\s*[0-9]+_([A-Za-z]{2})_", text)
        return m.group(1).upper() if m else ""

    def normalize_cm(value):
        return clean_text(value).upper()

    def to_number(value):
        text = clean_text(value)
        if not text:
            return None
        text = (
            text.replace("%", "")
                .replace(" V", "")
                .replace(" A", "")
                .replace(" h", "")
                .replace(",", ".")
                .strip()
        )
        try:
            return float(text)
        except Exception:
            return None

    def fmt_number(value, suffix=""):
        number = to_number(value)
        if number is None:
            return "-"
        if float(number).is_integer():
            return f"{int(number)}{suffix}"
        return f"{number:.1f}{suffix}"

    def detect_header_row(file_obj, required_names, max_rows=25):
        raw = pd.read_excel(file_obj, header=None, nrows=max_rows)
        targets = {normalize(x) for x in required_names}

        for idx in range(len(raw)):
            values = {normalize(v) for v in raw.iloc[idx].tolist() if clean_text(v)}
            if targets.issubset(values):
                return idx
        return 0

    def read_energy(file_obj):
        return pd.read_excel(file_obj, header=0).dropna(axis=1, how="all")

    def read_alarms(file_obj):
        header_row = detect_header_row(
            file_obj,
            required_names=["Name", "Alarm Source", "Clearance Status"]
        )
        df = pd.read_excel(file_obj, header=header_row).dropna(axis=1, how="all")
        return df, header_row

    def read_wos(file_obj):
        """
        Detecta la fila real de encabezados del WOs List.
        Reinicia el puntero antes de cada lectura para evitar
        lecturas incompletas del archivo subido por Streamlit.
        """
        file_obj.seek(0)

        raw = pd.read_excel(
            file_obj,
            header=None,
            nrows=25
        )

        header_row = 0

        for idx in range(len(raw)):
            vals = {
                normalize(v)
                for v in raw.iloc[idx].tolist()
                if clean_text(v)
            }

            has_state = any(
                x in vals
                for x in {
                    "wo state",
                    "estado de la tarea",
                    "estado de la tarea wo state",
                    "task state"
                }
            )

            has_site = any(
                x in vals
                for x in {
                    "nombre de site",
                    "site name",
                    "site"
                }
            )

            has_cm = any(
                x in vals
                for x in {
                    "numero de wo",
                    "cm",
                    "wo no",
                    "wo number",
                    "wo",
                    "work order"
                }
            )

            if has_state and has_site and has_cm:
                header_row = idx
                break

        # Volvemos al inicio antes de leer el archivo completo.
        file_obj.seek(0)

        df = pd.read_excel(
            file_obj,
            header=header_row
        ).dropna(axis=1, how="all")

        return df, header_row

    def find_col(columns, aliases):
        norm_cols = {normalize(c): c for c in columns}

        for alias in aliases:
            a = normalize(alias)
            if a in norm_cols:
                return norm_cols[a]

        candidates = []
        for c in columns:
            nc = normalize(c)
            for alias in aliases:
                na = normalize(alias)
                if na and (na in nc or nc in na):
                    candidates.append((len(na), c))
        if candidates:
            candidates.sort(reverse=True)
            return candidates[0][1]
        return None


    # =========================================================
    # CONFIG_MDA ONLINE - GOOGLE SHEETS
    # =========================================================

    monitored_site_keys = set()

    try:
        sdf = read_config_sheet("SITES_MONITOREADOS")

        site_col_cfg = find_col(
            sdf.columns,
            ["SITIO", "Nombre de Site", "SITE"]
        )

        if site_col_cfg:
            for v in sdf[site_col_cfg]:
                if clean_text(v):
                    monitored_site_keys.add(site_key(v))
        else:
            st.warning(
                "CONFIG_MDA_ONLINE → SITES_MONITOREADOS: "
                "no encontré la columna SITIO."
            )

    except Exception as exc:
        st.warning(
            "No pude consultar CONFIG_MDA_ONLINE → "
            f"SITES_MONITOREADOS: {exc}"
        )

    # =========================================================
    # GOOGLE SHEETS - COMENTARIOS MDA
    # =========================================================

    def sheets_configured():
        """
        Si Render/Streamlit no tiene secrets configurados,
        simplemente desactiva Google Sheets sin romper la app.
        """
        try:
            secrets = st.secrets
            return (
                "gcp_service_account" in secrets
                and "comments_sheet" in secrets
                and "spreadsheet_url" in secrets["comments_sheet"]
            )
        except Exception:
            return False

    @st.cache_resource(show_spinner=False)
    def get_comments_worksheet():
        import gspread
        from google.oauth2.service_account import Credentials

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]

        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]),
            scopes=scopes,
        )

        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(
            st.secrets["comments_sheet"]["spreadsheet_url"]
        )

        worksheet_name = st.secrets["comments_sheet"].get(
            "worksheet_name", "COMENTARIOS_MDA"
        )

        try:
            ws = spreadsheet.worksheet(worksheet_name)
        except gspread.WorksheetNotFound:
            ws = spreadsheet.add_worksheet(
                title=worksheet_name,
                rows=1000,
                cols=3
            )
            ws.update(
                values=[["SITIO", "COMENTARIO MDA", "ULTIMA ACTUALIZACION"]],
                range_name="A1:C1"
            )

        return ws

    def load_saved_comments():
        if not sheets_configured():
            return {}

        try:
            ws = get_comments_worksheet()
            records = ws.get_all_records()
            comments = {}
            for r in records:
                site = clean_text(r.get("SITIO", ""))
                comment = clean_text(r.get("COMENTARIO MDA", ""))
                if site:
                    comments[site] = comment
            return comments
        except Exception as exc:
            st.warning(f"No pude leer comentarios de Google Sheets: {exc}")
            return {}

    def save_comments_to_sheet(comments_dict):
        if not sheets_configured():
            return False, "Google Sheets aún no está configurado."

        try:
            ws = get_comments_worksheet()

            now = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M:%S")
            rows = [["SITIO", "COMENTARIO MDA", "ULTIMA ACTUALIZACION"]]

            for site in sorted(comments_dict):
                rows.append([
                    site,
                    clean_text(comments_dict[site]),
                    now
                ])

            ws.clear()
            ws.update(values=rows, range_name=f"A1:C{len(rows)}")
            return True, "Comentarios guardados."
        except Exception as exc:
            return False, f"No pude guardar comentarios: {exc}"

    # =========================================================
    # ARCHIVOS COMPARTIDOS
    # =========================================================
    energy_file = shared_energy
    alarms_file = shared_alarms
    wos_file = shared_wos

    if energy_file is None:
        st.info("Carga Site Energy Dashboard desde la barra lateral.")
        st.stop()

    # =========================================================
    # ENERGY DASHBOARD
    # =========================================================

    try:
        energy_df = read_energy(energy_file)
    except Exception as exc:
        st.error(f"No pude leer Site Energy Dashboard: {exc}")
        st.stop()

    energy_required = {
        "WO": "Related Task List",
        "SITIO": "Site Name",
        "PRIORIDAD": "Site Priority",
        "ESTADO": "Energy Site Status",
        "AUTONOMIA": "Site Autonomy",
        "SOC": "Soc",
        "VOLTAJE": "Site Voltage",
        "CORRIENTE": "Site Current Amperes",
        "DEPARTAMENTO": "Departamento",
        "COMENTARIO DASHBOARD": "Last Comment",
        "FECHA OCURRENCIA": "Firstoccurrence",
    }

    missing_energy = [
        source for source in energy_required.values()
        if source not in energy_df.columns
    ]

    if missing_energy:
        st.error(
            "Faltan columnas esperadas en Site Energy Dashboard: "
            + ", ".join(missing_energy)
        )
        st.stop()

    base = pd.DataFrame({
        output: energy_df[source]
        for output, source in energy_required.items()
    })

    # El Dashboard manda: una sola fila por SITE.
    # Si por alguna razón viniera duplicado, conservamos la primera aparición.
    base = base.drop_duplicates(subset=["SITIO"], keep="first").copy()

    base["_SITE_KEY"] = base["SITIO"].apply(site_key)

    # Si se cargó CONFIG_MDA, solo estos sites forman parte del monitoreo.
    if monitored_site_keys:
        base = base[
            base["_SITE_KEY"].isin(monitored_site_keys)
        ].copy()

    base["_REGION_CODE"] = base["SITIO"].apply(region_code)
    base["_CM_KEY"] = base["WO"].apply(normalize_cm)

    energy_site_list = base["SITIO"].astype(str).tolist()

    department_region_map = {}
    for _, r in base.iterrows():
        dep = clean_text(r["DEPARTAMENTO"])
        reg = clean_text(r["_REGION_CODE"]).upper()
        if dep and reg:
            department_region_map.setdefault(dep, set()).add(reg)

    # =========================================================
    # CURRENT ALARMS
    # =========================================================

    alarm_by_site = {}
    unmatched_alarm_rows = pd.DataFrame()

    if alarms_file is not None:
        try:
            alarms_df, alarms_header_row = read_alarms(alarms_file)
        except Exception as exc:
            st.error(f"No pude leer Current Alarms: {exc}")
            st.stop()

        required_alarm_cols = ["Name", "Alarm Source", "Clearance Status"]

        missing_alarm = [
            c for c in required_alarm_cols
            if c not in alarms_df.columns
        ]

        if missing_alarm:
            st.error(
                "No encontré estas columnas en Current Alarms: "
                + ", ".join(missing_alarm)
            )
            st.stop()

        # Fecha: Current usa First Occurred (NT)
        occurrence_col = None
        for c in alarms_df.columns:
            if normalize(c) == normalize("First Occurred (NT)"):
                occurrence_col = c
                break

        if occurrence_col is None:
            # fallback seguro
            occurrence_col = find_col(
                alarms_df.columns,
                ["first occurred nt", "first occurred", "first occurrence"]
            )

        selected = ["Name", "Alarm Source", "Clearance Status"]
        if occurrence_col:
            selected.append(occurrence_col)

        alarms = alarms_df[selected].copy()

        if occurrence_col:
            alarms.columns = [
                "ALARMA",
                "ALARM SOURCE",
                "STATUS DE LA ALARMA",
                "FECHA ALARMA",
            ]
        else:
            alarms.columns = [
                "ALARMA",
                "ALARM SOURCE",
                "STATUS DE LA ALARMA",
            ]
            alarms["FECHA ALARMA"] = pd.NaT

        alarms["_FECHA_DT"] = pd.to_datetime(
            alarms["FECHA ALARMA"],
            errors="coerce",
            dayfirst=True
        )

        def exact_or_suffix_match(alarm_source):
            key = site_key(alarm_source)
            if not key:
                return None

            exact = [
                s for s in energy_site_list
                if site_key(s) == key
            ]
            if len(exact) == 1:
                return exact[0]

            contained = [
                s for s in energy_site_list
                if key and (
                    key in site_key(s)
                    or site_key(s) in key
                )
            ]
            if len(contained) == 1:
                return contained[0]

            return None

        alarms["_MATCHED_SITE"] = alarms["ALARM SOURCE"].apply(
            exact_or_suffix_match
        )
        alarms["_SITE_KEY"] = alarms["_MATCHED_SITE"].apply(site_key)

        unmatched_alarm_rows = alarms[
            alarms["_MATCHED_SITE"].isna()
        ].copy()

        matched_alarms = alarms[
            alarms["_MATCHED_SITE"].notna()
        ].copy()

        for key, group in matched_alarms.groupby("_SITE_KEY"):
            group = group.sort_values(
                "_FECHA_DT",
                ascending=True,
                na_position="last"
            )

            pairs = []
            for _, r in group.iterrows():
                name = clean_text(r["ALARMA"]) or "-"
                status = clean_text(r["STATUS DE LA ALARMA"]) or "-"
                pair = (name, status)

                if pair not in pairs:
                    pairs.append(pair)

            alarm_by_site[key] = pairs

    # =========================================================
    # WOs LIST - TÉCNICO SOLO POR CM DEL DASHBOARD
    # =========================================================

    technician_by_cm = {}
    wos_header_row = None

    if wos_file is not None:
        try:
            wos_df, wos_header_row = read_wos(wos_file)
        except Exception as exc:
            st.error(f"No pude leer WOs List: {exc}")
            st.stop()

        # Columnas confirmadas del WOs List.
        cm_col = "Número de WO"
        tech_col = "Nombre de personal FLM asignado"

        # CM sí es obligatorio para cruzar con el Dashboard.
        # La columna de técnico es opcional: si no existe, queda vacía.
        if cm_col not in wos_df.columns:
            st.error(
                "Falta la columna esperada en WOs List: "
                + cm_col
            )
            st.write("Columnas encontradas:", list(wos_df.columns))
            st.stop()

        # NO cruzamos por SITE del WOs.
        # Solo por el CM que ya manda el Dashboard.
        if tech_col in wos_df.columns:
            wos_tmp = wos_df[[cm_col, tech_col]].copy()
            wos_tmp.columns = ["CM", "TECNICO"]
        else:
            wos_tmp = wos_df[[cm_col]].copy()
            wos_tmp["TECNICO"] = ""
            wos_tmp.columns = ["CM", "TECNICO"]

        for _, r in wos_tmp.iterrows():
            cm = normalize_cm(r["CM"])
            tech = clean_text(r["TECNICO"])

            if cm and cm not in technician_by_cm:
                technician_by_cm[cm] = tech or "-"

    # =========================================================
    # COMENTARIOS MDA PERSISTENTES
    # =========================================================

    saved_comments = load_saved_comments()

    if "comments_mda" not in st.session_state:
        st.session_state.comments_mda = saved_comments.copy()
    else:
        # Agrega desde Sheets solo los que aún no existen en la sesión.
        for site, comment in saved_comments.items():
            st.session_state.comments_mda.setdefault(site, comment)

    # =========================================================
    # OVERRIDES MANUALES DESDE CONFIG_MDA_ONLINE
    # Hoja: ESTADO_MANUAL_SITE
    # Clave: SITE (ignorando el prefijo numérico inicial)
    # =========================================================

    manual_site_map = {}

    try:
        mdf = read_config_sheet("ESTADO_MANUAL_SITE")

        # Normalizamos encabezados para tolerar tildes en ACTUALIZACIÓN.
        manual_cols = {
            normalize(str(c)): c
            for c in mdf.columns
        }

        m_site = manual_cols.get(normalize("SITE")) or manual_cols.get(normalize("SITIO"))

        if m_site:
            for _, mr in mdf.iterrows():
                raw_site = clean_text(mr.get(m_site, ""))
                if not raw_site:
                    continue

                mk = site_key(raw_site)

                def mvalue(label):
                    col = manual_cols.get(normalize(label))
                    return clean_text(mr.get(col, "")) if col else ""

                manual_site_map[mk] = {
                    "ESTADO": mvalue("ESTADO MANUAL"),
                    "AUTONOMIA": mvalue("AUTONOMIA MANUAL"),
                    "SOC": mvalue("SOC MANUAL"),
                    "VOLTAJE": mvalue("VOLTAJE MANUAL"),
                    "CORRIENTE": mvalue("CORRIENTE MANUAL"),
                    "COMENTARIO MDA": mvalue("COMENTARIO MDA"),
                }
        else:
            st.warning(
                "CONFIG_MDA_ONLINE → ESTADO_MANUAL_SITE: "
                "no encontré la columna SITE."
            )

    except Exception as exc:
        st.warning(
            "No pude consultar CONFIG_MDA_ONLINE → "
            f"ESTADO_MANUAL_SITE: {exc}"
        )

    # =========================================================
    # TABLA PRINCIPAL
    # =========================================================

    rows = []

    for _, r in base.iterrows():
        key = r["_SITE_KEY"]
        cm_key = r["_CM_KEY"]

        pairs = alarm_by_site.get(key, [])

        if pairs:
            alarm_text = " / ".join(p[0] for p in pairs)
            status_text = " / ".join(p[1] for p in pairs)
        else:
            alarm_text = "-"
            status_text = "-"

        site = clean_text(r["SITIO"]) or "-"
        technician = technician_by_cm.get(cm_key, "-")

        # Si existe dato manual para el SITE, manda el manual.
        # Si la celda manual está vacía, se conserva el valor del Energy Dashboard.
        manual = manual_site_map.get(key, {})

        estado_final = clean_text(manual.get("ESTADO")) or clean_text(r["ESTADO"]) or "-"
        autonomia_final = clean_text(manual.get("AUTONOMIA")) or r["AUTONOMIA"]
        soc_final = clean_text(manual.get("SOC")) or r["SOC"]
        voltaje_final = clean_text(manual.get("VOLTAJE")) or r["VOLTAJE"]
        corriente_final = clean_text(manual.get("CORRIENTE")) or r["CORRIENTE"]

        comentario_manual = clean_text(manual.get("COMENTARIO MDA"))
        comentario_mda_final = (
            comentario_manual
            if comentario_manual
            else st.session_state.comments_mda.get(site, "")
        )

        dashboard_date = pd.to_datetime(
            r["FECHA OCURRENCIA"],
            errors="coerce",
            dayfirst=True
        )

        date_text = (
            dashboard_date.strftime("%d/%m/%Y %H:%M:%S")
            if pd.notna(dashboard_date)
            else "-"
        )

        rows.append({
            "WO": clean_text(r["WO"]) or "-",
            "SITIO": site,
            "PRIORIDAD": clean_text(r["PRIORIDAD"]) or "-",
            "ESTADO": estado_final,
            "AUTONOMIA": fmt_number(autonomia_final, " h"),
            "SOC": fmt_number(soc_final, "%"),
            "VOLTAJE": fmt_number(voltaje_final, " V"),
            "CORRIENTE": fmt_number(corriente_final, " A"),
            "ALARMA": alarm_text,
            "STATUS DE LA ALARMA": status_text,
            "TECNICO ASIGNADO": technician,
            "DEPARTAMENTO": clean_text(r["DEPARTAMENTO"]) or "-",
            "COMENTARIO DASHBOARD": clean_text(r["COMENTARIO DASHBOARD"]) or "-",
            "COMENTARIO MDA": comentario_mda_final,
            "FECHA OCURRENCIA": date_text,
            "_ORDEN_FECHA": dashboard_date,
            "_REGION_CODE": r["_REGION_CODE"],
        })

    result = pd.DataFrame(rows).sort_values(
        "_ORDEN_FECHA",
        ascending=True,
        na_position="last"
    ).reset_index(drop=True)

    # =========================================================
    # FILTROS
    # =========================================================

    st.divider()
    st.subheader("Filtros")

    f1, f2, f3, f4 = st.columns(4)

    def filter_values(col):
        vals = result[col].fillna("-").astype(str).str.strip()
        return sorted([x for x in vals.unique() if x and x != "-"])

    with f1:
        f_dep = st.multiselect(
            "Departamento",
            filter_values("DEPARTAMENTO")
        )

    with f2:
        f_estado = st.multiselect(
            "Estado",
            filter_values("ESTADO")
        )

    with f3:
        f_prioridad = st.multiselect(
            "Prioridad",
            filter_values("PRIORIDAD")
        )

    with f4:
        f_alarm_status = st.multiselect(
            "Status de la alarma",
            filter_values("STATUS DE LA ALARMA")
        )

    search = st.text_input(
        "Buscar por SITE o WO",
        placeholder="Ej.: 013162560_CA_Namballe o CM-..."
    )

    filtered = result.copy()

    if f_dep:
        filtered = filtered[
            filtered["DEPARTAMENTO"].isin(f_dep)
        ]

    if f_estado:
        filtered = filtered[
            filtered["ESTADO"].isin(f_estado)
        ]

    if f_prioridad:
        filtered = filtered[
            filtered["PRIORIDAD"].isin(f_prioridad)
        ]

    if f_alarm_status:
        filtered = filtered[
            filtered["STATUS DE LA ALARMA"].isin(f_alarm_status)
        ]

    if search.strip():
        q = re.escape(search.strip())
        filtered = filtered[
            filtered["SITIO"].astype(str).str.contains(
                q, case=False, na=False
            )
            |
            filtered["WO"].astype(str).str.contains(
                q, case=False, na=False
            )
        ]

    # =========================================================
    # EDITOR DE COMENTARIO MDA
    # =========================================================

    st.subheader("Tabla consolidada")

    visible_cols = [
        "WO",
        "SITIO",
        "PRIORIDAD",
        "ESTADO",
        "AUTONOMIA",
        "SOC",
        "VOLTAJE",
        "CORRIENTE",
        "ALARMA",
        "STATUS DE LA ALARMA",
        "TECNICO ASIGNADO",
        "DEPARTAMENTO",
        "COMENTARIO DASHBOARD",
        "COMENTARIO MDA",
        "FECHA OCURRENCIA",
    ]

    edit_df = filtered[visible_cols].copy()

    edited = st.data_editor(
        edit_df,
        use_container_width=True,
        hide_index=True,
        height=650,
        disabled=[
            c for c in visible_cols
            if c != "COMENTARIO MDA"
        ],
        key="comments_editor"
    )

    # Persistencia inmediata en session_state por SITE.
    for idx in edited.index:
        site = clean_text(edited.loc[idx, "SITIO"])
        comment = clean_text(edited.loc[idx, "COMENTARIO MDA"])

        if site:
            st.session_state.comments_mda[site] = comment

    c_save1, c_save2 = st.columns([1, 3])

    with c_save1:
        if st.button("💾 Guardar comentarios MDA", use_container_width=True):
            ok, msg = save_comments_to_sheet(
                st.session_state.comments_mda
            )
            if ok:
                st.success(msg)
            else:
                st.warning(msg)

    with c_save2:
        if sheets_configured():
            st.caption(
                "Google Sheets conectado: los comentarios quedan vinculados al SITE "
                "aunque cambies los Excel o ingreses desde otra PC."
            )
        else:
            st.caption(
                "Aún falta configurar Google Sheets. Mientras tanto, los comentarios "
                "se conservan solo durante la sesión actual."
            )

    # =========================================================
    # SEGUNDO CUADRO: ALARMAS SIN SITE EN DASHBOARD
    # =========================================================

    st.divider()
    st.subheader("Alarmas de sites fuera del Energy Dashboard")

    if alarms_file is None:
        st.info("Carga Current Alarms para ver este cuadro.")
    elif unmatched_alarm_rows.empty:
        st.info("No hay alarmas de sites fuera del Energy Dashboard.")
    else:
        unmatched = unmatched_alarm_rows.copy()

        # Solo mostrar alarmas pendientes.
        unmatched = unmatched[
            unmatched["STATUS DE LA ALARMA"]
            .astype(str)
            .map(normalize)
            .eq("uncleared")
        ]

        # Si existe lista maestra, solo mostrar los sites que sí monitoreamos.
        if monitored_site_keys:
            unmatched = unmatched[
                unmatched["ALARM SOURCE"]
                .apply(site_key)
                .isin(monitored_site_keys)
            ]

        # El filtro Departamento afecta también al segundo cuadro
        if f_dep:
            selected_codes = set()
            for dep in f_dep:
                selected_codes.update(
                    department_region_map.get(dep, set())
                )

            unmatched["_REGION_CODE"] = unmatched[
                "ALARM SOURCE"
            ].apply(region_code)

            if selected_codes:
                unmatched = unmatched[
                    unmatched["_REGION_CODE"].isin(selected_codes)
                ]
            else:
                unmatched = unmatched.iloc[0:0]

        unmatched = unmatched.sort_values(
            "_FECHA_DT",
            ascending=True,
            na_position="last"
        )

        if unmatched.empty:
            st.info("No hay alarmas fuera del Dashboard para el filtro seleccionado.")
        else:
            second = unmatched[
                [
                    "ALARM SOURCE",
                    "ALARMA",
                    "STATUS DE LA ALARMA",
                    "FECHA ALARMA",
                ]
            ].copy()

            second.columns = [
                "SITIO",
                "ALARMA",
                "STATUS DE LA ALARMA",
                "FECHA OCURRENCIA",
            ]

            second["FECHA OCURRENCIA"] = pd.to_datetime(
                second["FECHA OCURRENCIA"],
                errors="coerce",
                dayfirst=True
            ).dt.strftime("%d/%m/%Y %H:%M:%S").fillna("-")

            st.dataframe(
                second,
                use_container_width=True,
                hide_index=True
            )

    # =========================================================
    # DIAGNÓSTICO
    # =========================================================

    with st.expander("Diagnóstico técnico"):
        st.write(f"Sites únicos del Dashboard: {len(base)}")
        if alarms_file is not None:
            st.write(
                f"Current Alarms encabezado en fila: {alarms_header_row + 1}"
            )
        if wos_file is not None:
            st.write(
                f"WOs List encabezado en fila: {wos_header_row + 1}"
            )
            st.write(
                "Cruce WOs confirmado: Número de WO → Nombre de personal FLM asignado"
            )
            matched_techs = sum(
                1 for cm in base["_CM_KEY"]
                if cm in technician_by_cm
            )
            st.write(
                f"CM del Dashboard encontrados en WOs List: "
                f"{matched_techs} de {len(base)}"
            )

    # =========================================================
    # COPIAR
    # =========================================================

    copy_df = edited.copy()
    tsv = copy_df.to_csv(sep="\t", index=False)

    with st.expander("Copiar y pegar en Excel"):
        st.code(tsv, language=None)

else:
    st.header("🔄 Cambio de Turno")

    import re
    import json
    import html
    from datetime import date
    import pandas as pd
    import streamlit as st


    st.caption("Carga el WOs List y genera texto listo para copiar y pegar en WhatsApp.")

    # =========================
    # UTILIDADES
    # =========================

    def normalize(text):
        text = str(text).strip().lower()
        text = (
            text.replace("á","a").replace("é","e").replace("í","i")
                .replace("ó","o").replace("ú","u").replace("ñ","n")
        )
        return re.sub(r"[^a-z0-9]+", " ", text).strip()

    def clean(value):
        if pd.isna(value):
            return ""
        text = str(value).strip()
        if text.lower() in {"nan", "none", "<na>"}:
            return ""
        return text

    def auto_match(columns, aliases):
        norm_cols = {normalize(c): c for c in columns}
        for alias in aliases:
            n = normalize(alias)
            if n in norm_cols:
                return norm_cols[n]

        candidates = []
        for col in columns:
            nc = normalize(col)
            for alias in aliases:
                na = normalize(alias)
                if na and (na in nc or nc in na):
                    candidates.append((len(na), col))
        if candidates:
            candidates.sort(reverse=True)
            return candidates[0][1]
        return None


    def copy_button(text, label="📋 Copiar para WhatsApp", key="copy"):
        """
        Botón de copia al portapapeles usando HTML/JS embebido.
        """
        safe_text = json.dumps(text)
        safe_label = html.escape(label)
        button_id = re.sub(r"[^a-zA-Z0-9_-]", "_", key)

        st.components.v1.html(
            f"""
            <div>
              <button id="{button_id}"
                style="
                  width:100%;
                  padding:0.55rem 0.85rem;
                  border-radius:0.5rem;
                  border:1px solid rgba(128,128,128,.35);
                  background:transparent;
                  cursor:pointer;
                  font-size:0.95rem;
                  font-weight:600;
                "
              >{safe_label}</button>
              <span id="{button_id}_msg"
                style="margin-left:10px;font-size:0.9rem;"></span>
            </div>

            <script>
              const btn = document.getElementById("{button_id}");
              const msg = document.getElementById("{button_id}_msg");
              const text = {safe_text};

              btn.addEventListener("click", async () => {{
                try {{
                  await navigator.clipboard.writeText(text);
                  msg.textContent = "✓ Copiado";
                  setTimeout(() => msg.textContent = "", 1800);
                }} catch (err) {{
                  const ta = document.createElement("textarea");
                  ta.value = text;
                  document.body.appendChild(ta);
                  ta.select();
                  document.execCommand("copy");
                  document.body.removeChild(ta);
                  msg.textContent = "✓ Copiado";
                  setTimeout(() => msg.textContent = "", 1800);
                }}
              }});
            </script>
            """,
            height=55,
        )

    ALIASES = {
        "ESTADO": [
            "Estado de la tarea (WO State)"
        ],
        "CM": [
            "Número de WO"
        ],
        "SITE": [
            "Nombre de Site"
        ],
        "TECNICO": [
            "Nombre de personal FLM asignado"
        ],
        "TIPO_TAREA": [
            "Tipo tarea"
        ],
        "DEPARTAMENTO": [
            "departamento", "department", "region", "región"
        ],
        "CRITICIDAD": [
            "Fault Level"
        ],
        "PRIORIDAD_SITE": [
            "Prioridad del Site"
        ],
        "HORA_TICKET": [
            "First Occurred",
            "First Occurrence",
            "Fecha de creación",
            "Fecha de creacion",
            "Created Time",
            "Creation Time",
            "Hora del ticket",
            "Hora Ticket"
        ],
    }

    def short_technician_name(value):
        text = clean(value)
        if not text:
            return "-"
        people = re.split(r"\s*[;/]\s*", text)
        result = []
        for person in people:
            t = person.split()
            if len(t) <= 2:
                result.append(" ".join(t))
            elif len(t) == 3:
                result.append(f"{t[0]} {t[1]}")
            else:
                result.append(f"{t[0]} {t[2]}")
        return " / ".join(result)

    def site_region_code(site):
        text = clean(site)
        m = re.match(r"^\s*[0-9]+_([A-Za-z]{2})_", text)
        return m.group(1).upper() if m else ""

    def derive_department_from_site(site):
        code = site_region_code(site)
        mapping = {
            "CA": "CAJAMARCA",
            "PI": "PIURA",
            "LI": "LA LIBERTAD",
            "LA": "LAMBAYEQUE",
            "AN": "ANCASH",
            "LO": "LORETO",
            "SM": "SAN MARTIN",
            "PN": "PUNO",
            "JU": "JUNIN",
            "IC": "ICA",
            "CS": "CUSCO",
            "AQ": "AREQUIPA",
            "TU": "TUMBES",
            "CP": "PASCO",
            "LH": "HUANUCO",
        }
        return mapping.get(code, code or "SIN DEPARTAMENTO")

    # Comentarios persistentes para SITIOS EN MONITOREO.
    # Clave: CM + SITE normalizado.
    seguimiento_cm_map = {}

    def make_line(row, section):
        cm = clean(row.get("CM")) or "-"
        site = clean(row.get("SITE")) or "-"
        tecnico = short_technician_name(row.get("TECNICO"))
        tipo = clean(row.get("TIPO_TAREA")) or "-"

        if section == "monitor":
            line = f"{cm} / {site} / {tipo}"
        else:
            line = f"{cm} / {site} / {tecnico}"

        acceso = clean(row.get("ESTADO_ACCESO"))
        obs = clean(row.get("OBSERVACION_ACCESO"))

        # Reglas de acceso:
        # - CLOSED: nunca mostrar acceso.
        # - TAREAS EN CURSO: siempre mostrar acceso.
        #   Si el site no figura en el Excel de accesos, usar "SIN ACCESO".
        # - SITIOS EN MONITOREO: mostrar acceso solo si existe dato para ese site.
        if section == "course":
            acceso_mostrar = acceso if acceso else "SIN ACCESO"
            line += f"\n↳ Acceso: {acceso_mostrar}"
            if acceso and obs:
                line += f" - {obs}"

        elif section == "monitor":
            if acceso:
                line += f"\n↳ Acceso: {acceso}"
                if obs:
                    line += f" - {obs}"

            cm_key_comment = normalize(cm)
            site_key_comment = normalize(
                re.sub(r"^\s*[0-9]+_?", "", site)
            )
            comentario = seguimiento_cm_map.get(
                (cm_key_comment, site_key_comment),
                ""
            )

            if comentario:
                line += f"\n↳ Comentario: {comentario}"

        return line

    def build_whatsapp_text(
        df, zona, mda_salida, mda_ingreso, supervisor, fecha_reporte,
        titulo_departamento=None, filtrar_monitoreo_general=False
    ):
        header = []
        header.append("*CAMBIO DE TURNO:*")
        header.append(f"*Fecha:* {fecha_reporte.strftime('%d/%m/%Y')}")
        if titulo_departamento:
            header.append(f"*Departamento:* {titulo_departamento}")
        header.append(f"*Zona:* {zona}")
        header.append(f"*MDA Salida:* {mda_salida}")
        header.append(f"*MDA Ingreso:* {mda_ingreso}")
        header.append(f"*Supervisor:* {supervisor}")
        header.append("")

        state_norm = df["ESTADO"].astype(str).map(normalize)

        closed_mask = state_norm.eq("closed")
        course_mask = state_norm.isin(["dispatched", "accepted", "inprocess"])
        # SITIOS EN MONITOREO siempre parte de los Unscheduled.
        unscheduled_mask = state_norm.isin([
            "unscheduled", "uncheduled", "uncheluded", "unschedule",
            "unschuduled", "unshuduled"
        ])

        # Regla obligatoria:
        # Todo ticket Unscheduled asociado a un site P0, P0+ o P1
        # debe monitorearse sí o sí.
        priority_norm = (
            df["PRIORIDAD_SITE"]
            .astype(str)
            .str.upper()
            .str.replace(" ", "", regex=False)
        )
        mandatory_priority = priority_norm.isin(["P0", "P0+", "P1"])

        monitor_mask = unscheduled_mask.copy()

        # REGLA PRINCIPAL:
        # Solo se monitorean sites incluidos en SITES_MONITOREADOS.
        # La prioridad P0/P0+/P1 NO permite saltarse esta lista.
        if monitored_turno_keys:
            site_keys_turno = df["SITE"].astype(str).apply(
                lambda x: normalize(
                    re.sub(r"^\s*[0-9]+_?", "", clean(x))
                )
            )
            listed_site = site_keys_turno.isin(monitored_turno_keys)
            monitor_mask = unscheduled_mask & listed_site

        # REPORTE GENERAL:
        # Mostrar Unscheduled cuando cumpla cualquiera de estas reglas:
        # - Fault Level = Media o Alta
        # - Tipo tarea contiene "Ausencia"
        # - Prioridad del Site = P0, P0+ o P1
        #
        # P0/P0+/P1 evita el filtro de criticidad/ausencia,
        # pero NUNCA evita el filtro de SITES_MONITOREADOS.
        if filtrar_monitoreo_general:
            criticality_norm = df["CRITICIDAD"].astype(str).map(normalize)
            task_type_norm = df["TIPO_TAREA"].astype(str).map(normalize)

            medium_or_high_criticality = criticality_norm.str.contains(
                r"\bmedia\b|\bmedium\b|\balta\b|\bhigh\b",
                regex=True,
                na=False
            )

            ausencia = task_type_norm.str.contains(
                r"\bausencia\b",
                regex=True,
                na=False
            )

            general_rule = (
                medium_or_high_criticality
                | ausencia
                | mandatory_priority
            )

            monitor_mask = monitor_mask & general_rule

        sections = []

        sections.append("*TAREAS CERRADAS*")
        closed = df[closed_mask]
        if closed.empty:
            sections.append("-")
        else:
            sections.extend(make_line(r, "closed") for _, r in closed.iterrows())

        sections.append("")
        sections.append("*TAREAS EN CURSO*")
        course = df[course_mask]
        if course.empty:
            sections.append("-")
        else:
            sections.extend(make_line(r, "course") for _, r in course.iterrows())

        sections.append("")
        sections.append("*SITIOS EN MONITOREO*")
        monitor = df[monitor_mask]
        if monitor.empty:
            sections.append("-")
        else:
            sections.extend(make_line(r, "monitor") for _, r in monitor.iterrows())

        return "\n".join(header + sections)

    # =========================
    # DATOS DE CABECERA
    # =========================

    st.subheader("Datos del cambio de turno")

    fecha_reporte = st.date_input(
        "Fecha",
        value=date.today(),
        format="DD/MM/YYYY"
    )

    h1, h2, h3, h4 = st.columns(4)

    with h1:
        zona = st.text_input("Zona", value="")
    with h2:
        mda_salida = st.text_input("MDA Salida", value="")
    with h3:
        mda_ingreso = st.text_input("MDA Ingreso", value="")
    with h4:
        supervisor = st.text_input("Supervisor", value="")

    # =========================
    # CARGA EXCEL
    # =========================

    uploaded = shared_wos

    if uploaded is None:
        st.info("Carga WOs List desde la barra lateral.")
        st.stop()

    try:
        xls = pd.ExcelFile(uploaded)
        sheet = st.selectbox("Hoja", xls.sheet_names, index=0)
        df = pd.read_excel(uploaded, sheet_name=sheet)
    except Exception as exc:
        st.error(f"No pude leer el Excel: {exc}")
        st.stop()

    df = df.dropna(axis=1, how="all")

    mapping = {field: auto_match(df.columns, aliases) for field, aliases in ALIASES.items()}

    missing = [
        field for field, col in mapping.items()
        if col is None and field not in [
            "DEPARTAMENTO",
            "CRITICIDAD",
            "HORA_TICKET",
            "TECNICO"
        ]
    ]

    if missing:
        st.warning("No pude detectar automáticamente: " + ", ".join(missing))
        with st.expander("Asignar columnas manualmente"):
            options = ["— No encontrada —"] + list(df.columns)
            for field in missing:
                selected = st.selectbox(field, options, key=f"map_{field}")
                if selected != "— No encontrada —":
                    mapping[field] = selected

    required = ["ESTADO", "CM", "SITE", "TIPO_TAREA"]
    if any(mapping.get(x) is None for x in required):
        st.stop()

    work = pd.DataFrame({
        "ESTADO": df[mapping["ESTADO"]],
        "CM": df[mapping["CM"]],
        "SITE": df[mapping["SITE"]],
        "TECNICO": (
            df[mapping["TECNICO"]]
            if mapping.get("TECNICO")
            else ""
        ),
        "TIPO_TAREA": df[mapping["TIPO_TAREA"]],
        "PRIORIDAD_SITE": df[mapping["PRIORIDAD_SITE"]],
    })

    if mapping.get("HORA_TICKET"):
        work["HORA_TICKET"] = df[mapping["HORA_TICKET"]]
    else:
        work["HORA_TICKET"] = pd.NaT

    if mapping.get("DEPARTAMENTO"):
        work["DEPARTAMENTO"] = df[mapping["DEPARTAMENTO"]].apply(clean)
    else:
        work["DEPARTAMENTO"] = work["SITE"].apply(derive_department_from_site)

    if mapping.get("CRITICIDAD"):
        work["CRITICIDAD"] = df[mapping["CRITICIDAD"]].apply(clean)
    else:
        work["CRITICIDAD"] = ""

    work["DEPARTAMENTO"] = work["DEPARTAMENTO"].replace("", "SIN DEPARTAMENTO")
    work["ESTADO_ACCESO"] = ""
    work["OBSERVACION_ACCESO"] = ""

    monitored_turno_keys = set()

    # CONFIG_MDA_ONLINE se consulta directamente desde Google Sheets.
    try:
        # Hoja ACCESOS
        adf = read_config_sheet("ACCESOS")

        sc = auto_match(adf.columns, ["SITIO", "Nombre de Site", "SITE"])
        ac = auto_match(adf.columns, ["ESTADO ACCESO", "Estado de acceso", "ACCESO"])
        oc = auto_match(adf.columns, ["OBSERVACIÓN", "OBSERVACION", "Comentario", "NOTA"])

        if sc is None or ac is None:
            st.warning(
                "CONFIG_MDA_ONLINE → ACCESOS necesita las columnas "
                "SITIO y ESTADO ACCESO."
            )
        else:
            amap = {}
            for _, ar in adf.iterrows():
                k = normalize(
                    re.sub(r"^\s*[0-9]+_?", "", clean(ar[sc]))
                )
                if k:
                    amap[k] = (
                        clean(ar[ac]),
                        clean(ar[oc]) if oc else ""
                    )

            for i in work.index:
                k = normalize(
                    re.sub(
                        r"^\s*[0-9]+_?",
                        "",
                        clean(work.at[i, "SITE"])
                    )
                )
                info = amap.get(k)
                if info:
                    work.at[i, "ESTADO_ACCESO"], work.at[i, "OBSERVACION_ACCESO"] = info

        # Hoja SITES_MONITOREADOS
        sdf = read_config_sheet("SITES_MONITOREADOS")

        scc = auto_match(
            sdf.columns,
            ["SITIO", "Nombre de Site", "SITE"]
        )

        if scc:
            for v in sdf[scc]:
                if clean(v):
                    monitored_turno_keys.add(
                        normalize(
                            re.sub(
                                r"^\s*[0-9]+_?",
                                "",
                                clean(v)
                            )
                        )
                    )

        # Hoja SEGUIMIENTO_CM
        segdf = read_config_sheet("SEGUIMIENTO_CM")

        cmc = auto_match(segdf.columns, ["CM", "Número de WO", "Numero de WO"])
        sgc = auto_match(segdf.columns, ["SITE", "SITIO", "Nombre de Site"])
        cgc = auto_match(segdf.columns, ["COMENTARIO MDA", "COMENTARIO", "Comentario MDA"])

        if cmc and sgc and cgc:
            for _, sr in segdf.iterrows():
                scm = clean(sr[cmc])
                ssite = clean(sr[sgc])
                scomment = clean(sr[cgc])

                if not scm or not ssite or not scomment:
                    continue

                kcm = normalize(scm)
                ksite = normalize(
                    re.sub(r"^\s*[0-9]+_?", "", ssite)
                )
                seguimiento_cm_map[(kcm, ksite)] = scomment

    except Exception as exc:
        st.warning(f"No pude consultar CONFIG_MDA_ONLINE: {exc}")


    # =========================
    # SEGUIMIENTO DE TICKETS
    # =========================

    st.divider()
    st.subheader("🎫 Seguimiento de tickets")

    def site_key_turno(value):
        text = clean(value)
        if not text:
            return ""
        suffix = re.sub(r"^\s*[0-9]+_?", "", text).strip()
        return normalize(suffix if suffix else text)

    energy_status_by_site = {}
    if shared_energy is not None:
        try:
            shared_energy.seek(0)
            edf = pd.read_excel(shared_energy).dropna(axis=1, how="all")
            if "Site Name" in edf.columns and "Energy Site Status" in edf.columns:
                for _, er in edf.iterrows():
                    k = site_key_turno(er["Site Name"])
                    if k:
                        energy_status_by_site[k] = clean(er["Energy Site Status"]) or "-"
        except Exception as exc:
            st.warning(f"No pude cruzar Energy Dashboard para seguimiento de tickets: {exc}")

    active_alarm_by_site = {}
    if shared_alarms is not None:
        try:
            shared_alarms.seek(0)
            raw = pd.read_excel(shared_alarms, header=None, nrows=25)
            target = {
                normalize("Name"),
                normalize("Alarm Source"),
                normalize("Clearance Status")
            }
            header_row = 0
            for i in range(len(raw)):
                vals = {
                    normalize(v)
                    for v in raw.iloc[i].tolist()
                    if clean(v)
                }
                if target.issubset(vals):
                    header_row = i
                    break

            shared_alarms.seek(0)
            cadf = pd.read_excel(
                shared_alarms,
                header=header_row
            ).dropna(axis=1, how="all")

            if all(c in cadf.columns for c in ["Name", "Alarm Source", "Clearance Status"]):
                cadf = cadf[
                    cadf["Clearance Status"]
                    .astype(str)
                    .map(normalize)
                    .eq("uncleared")
                ].copy()

                cadf["_SITE_KEY"] = cadf["Alarm Source"].apply(site_key_turno)

                for k, grp in cadf.groupby("_SITE_KEY"):
                    alarms_list = []
                    for value in grp["Name"]:
                        name = clean(value)
                        if name and name not in alarms_list:
                            alarms_list.append(name)
                    if k:
                        active_alarm_by_site[k] = " / ".join(alarms_list) if alarms_list else "-"
        except Exception as exc:
            st.warning(f"No pude cruzar Current Alarms para seguimiento de tickets: {exc}")

    ticket_rows = []

    for _, tr in work.iterrows():
        site = clean(tr["SITE"])
        k = site_key_turno(site)

        if monitored_turno_keys and k not in monitored_turno_keys:
            continue

        ticket_dt = pd.to_datetime(
            tr.get("HORA_TICKET"),
            errors="coerce",
            dayfirst=True
        )

        ticket_rows.append({
            "CM": clean(tr["CM"]) or "-",
            "STATUS CM": clean(tr["ESTADO"]) or "-",
            "SITE": site or "-",
            "PRIORIDAD": clean(tr["PRIORIDAD_SITE"]) or "-",
            "TIPO TAREA": clean(tr["TIPO_TAREA"]) or "-",
            "ALARMA ACTIVA": active_alarm_by_site.get(k, "-"),
            "ESTADO ENERGY": energy_status_by_site.get(k, "-"),
            "TECNICO": short_technician_name(tr["TECNICO"]),
            "HORA TICKET": (
                ticket_dt.strftime("%d/%m/%Y %H:%M")
                if pd.notna(ticket_dt)
                else "-"
            ),
            "_DT_TICKET": ticket_dt,
        })

    ticket_table = pd.DataFrame(ticket_rows)

    if ticket_table.empty:
        st.info("No hay tickets de sites monitoreados para mostrar.")
    else:
        ticket_table = ticket_table.sort_values(
            "_DT_TICKET",
            ascending=True,
            na_position="last"
        ).reset_index(drop=True)

        tf1, tf2, tf3 = st.columns(3)

        with tf1:
            ticket_status_filter = st.multiselect(
                "Status CM",
                sorted([x for x in ticket_table["STATUS CM"].dropna().unique() if clean(x)]),
                key="ticket_status_filter"
            )

        with tf2:
            ticket_priority_filter = st.multiselect(
                "Prioridad ticket",
                sorted([x for x in ticket_table["PRIORIDAD"].dropna().unique() if clean(x)]),
                key="ticket_priority_filter"
            )

        with tf3:
            ticket_search = st.text_input(
                "Buscar CM o SITE",
                key="ticket_search"
            )

        ticket_view = ticket_table.copy()

        if ticket_status_filter:
            ticket_view = ticket_view[
                ticket_view["STATUS CM"].isin(ticket_status_filter)
            ]

        if ticket_priority_filter:
            ticket_view = ticket_view[
                ticket_view["PRIORIDAD"].isin(ticket_priority_filter)
            ]

        if ticket_search.strip():
            q = re.escape(ticket_search.strip())
            ticket_view = ticket_view[
                ticket_view["CM"].astype(str).str.contains(q, case=False, na=False)
                |
                ticket_view["SITE"].astype(str).str.contains(q, case=False, na=False)
            ]

        st.dataframe(
            ticket_view.drop(columns=["_DT_TICKET"], errors="ignore"),
            use_container_width=True,
            hide_index=True,
            height=430
        )

        st.caption(
            "WOs List manda el ticket y su estado. Current Alarms añade alarmas Uncleared "
            "y Energy Dashboard indica si el site ya aparece Cargando, Descargando o Caído."
        )

    # =========================
    # SALIDA GENERAL
    # =========================

    st.divider()
    st.subheader("Reporte general")

    general_text = build_whatsapp_text(
        work,
        zona,
        mda_salida,
        mda_ingreso,
        supervisor,
        fecha_reporte,
        filtrar_monitoreo_general=True
    )

    st.text_area(
        "Texto general para WhatsApp",
        value=general_text,
        height=430,
        key="general_text"
    )

    st.markdown("**Copiar reporte:**")
    copy_button(
        general_text,
        label="📋 COPIAR REPORTE GENERAL",
        key="copy_general"
    )

    # =========================
    # SALIDA POR DEPARTAMENTO
    # =========================

    st.divider()
    st.subheader("Reportes por departamento")

    departments = sorted(
        [d for d in work["DEPARTAMENTO"].dropna().astype(str).unique() if d.strip()]
    )

    for dep in departments:
        dep_df = work[work["DEPARTAMENTO"].astype(str) == dep]

        with st.expander(dep, expanded=False):
            dep_text = build_whatsapp_text(
                dep_df,
                zona,
                mda_salida,
                mda_ingreso,
                supervisor,
                fecha_reporte,
                titulo_departamento=dep,
                filtrar_monitoreo_general=False
            )

            st.text_area(
                f"Texto - {dep}",
                value=dep_text,
                height=390,
                key=f"text_{dep}"
            )

            st.markdown("**Copiar reporte:**")
            copy_button(
                dep_text,
                label=f"📋 COPIAR {dep}",
                key=f"copy_{dep}"
            )

    # =========================
    # VISTA DE DATOS DETECTADOS
    # =========================

    with st.expander("Ver columnas detectadas"):
        st.caption(
            "Columnas confirmadas: Estado = Estado de la tarea (WO State) | CM = Número de WO | Site = Nombre de Site | Técnico = Nombre de personal FLM asignado | Criticidad = Fault Level | Prioridad = Prioridad del Site | Ausencia = Tipo tarea"
        )
        detected = pd.DataFrame({
            "Campo": list(mapping.keys()),
            "Columna Excel": [mapping[k] or "-" for k in mapping]
        })
        st.dataframe(detected, hide_index=True, use_container_width=True)
