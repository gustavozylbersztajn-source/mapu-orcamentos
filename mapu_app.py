#!/usr/bin/env python3
"""MAPU — Interface Web para Geração de Orçamentos"""

import importlib.util, json
from datetime import date, timedelta
from pathlib import Path
import streamlit as st

# ── Carrega o módulo principal ─────────────────────────────────────────────────
@st.cache_resource
def load_mapu():
    _path = Path(__file__).parent / "gerar_orcamento.py"
    spec = importlib.util.spec_from_file_location("mapu", str(_path))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m

@st.cache_data
def load_agencies():
    p = Path(__file__).parent / "agencies.json"
    if p.exists():
        return json.loads(p.read_text())
    # Nuvem: lê de st.secrets
    try:
        result = {}
        for key in st.secrets["agencies"]:
            result[key] = dict(st.secrets["agencies"][key])
        return result
    except Exception:
        return {}

m = load_mapu()

CABANAS = ["COIGUE", "NIRE", "CHAITEN", "CORCOVADO"]

# ── Layout base ────────────────────────────────────────────────────────────────
st.set_page_config(page_title="MAPU LODGE | Estimates APP", page_icon="🏕️", layout="centered")

st.markdown("""
<style>
    .block-container { max-width: 720px; }
    h1 { font-size: 1.6rem; }
    h3 { font-size: 1.1rem; margin-top: 1.5rem; }
</style>
""", unsafe_allow_html=True)

# ── Tela de aprovação (Gustavo) ────────────────────────────────────────────────
_approve_key = st.query_params.get("approve")
if _approve_key:
    st.title("MAPU | Aprovação de Orçamento")
    st.divider()

    json_path, pdf_path = m.find_budget_files(_approve_key)

    if json_path is None:
        st.error(f"Orçamento `{_approve_key}` não encontrado.")
        # debug temporário
        dbx = m._get_dropbox_client()
        if dbx is None:
            st.warning("Dropbox client: None — secrets não configurados")
        else:
            st.info("Dropbox client: OK")
            slug = _approve_key
            year, month = slug[:4], slug[4:6]
            try:
                result = dbx.files_list_folder(f"/{year}/{month}")
                names = [e.name for e in result.entries]
                st.write(f"Pastas em /{year}/{month}:", names)
            except Exception as ex:
                st.write(f"Erro listando /{year}/{month}: {ex}")
        st.stop()

    raw = json.loads(Path(json_path).read_text())
    # reconstruct basic display data
    ag_name    = raw.get("agency_name", "—")
    ag_contact = raw.get("agency_contact", "—")
    ag_email   = raw.get("agency_email", "—")
    client_raw = raw.get("client_raw", "—")
    checkin    = raw.get("checkin", "—")
    checkout   = raw.get("checkout", "—")
    nights     = raw.get("nights", "—")
    cabins     = raw.get("cabins", {})
    total_cc   = raw.get("total_cc", 0)

    st.markdown(f"**Agência:** {ag_name} · {ag_contact} · {ag_email}")
    st.markdown(f"**Cliente:** {client_raw}")
    st.markdown(f"**Check-in:** {checkin} · **Check-out:** {checkout} · **{nights} noite(s)**")
    st.markdown(f"**Cabanas:** {', '.join(cabins.keys())}")
    if total_cc:
        st.metric("TOTAL (c/ CC 4%)", f"CLP {total_cc:,.0f}".replace(",", "."))

    st.divider()

    if pdf_path and Path(pdf_path).exists():
        with open(pdf_path, "rb") as f:
            st.download_button(
                "⬇️ Baixar PDF para revisão",
                f.read(),
                file_name=Path(pdf_path).name,
                mime="application/pdf",
                use_container_width=True,
            )
        st.divider()

    col_a, col_b = st.columns(2)
    if col_a.button("✓ Aprovar e Enviar PDF", type="primary", use_container_width=True):
        # rebuild calcs dict from saved json
        calcs_saved = {k: raw.get(k, 0) for k in
                       ("lodging", "food", "total", "total_cc", "agency_fee",
                        "food_embedded", "usd_ref", "per_adult")}
        ok = m.send_approved_budget_email(raw, calcs_saved, pdf_path)
        if ok:
            st.success(f"PDF enviado para {ag_email}")
        else:
            st.error("Erro ao enviar — verifique configuração SMTP.")

    if col_b.button("✗ Recusar", use_container_width=True):
        st.warning("Orçamento recusado. Nenhum email enviado.")

    st.stop()

# ── Login ──────────────────────────────────────────────────────────────────────
if "agency_user" not in st.session_state:
    st.session_state["agency_user"] = None

if st.session_state["agency_user"] is None:
    st.title("MAPU | Experiences Lodge")
    st.divider()
    st.markdown("### Acesso para Agências")
    with st.form("login"):
        usuario  = st.text_input("Usuário")
        senha    = st.text_input("Senha", type="password")
        entrar   = st.form_submit_button("Entrar", use_container_width=True, type="primary")

    if entrar:
        agencies = load_agencies()
        u = usuario.strip().lower()
        if u in agencies and agencies[u]["password"] == senha.strip():
            st.session_state["agency_user"] = u
            st.session_state["agency_info"] = agencies[u]
            st.rerun()
        else:
            st.error("Usuário ou senha incorretos.")
    st.stop()

# ── App principal (autenticado) ────────────────────────────────────────────────
agency_info = st.session_state["agency_info"]

col_title, col_lang, col_logout = st.columns([3, 1, 1])
col_title.title("MAPU | Experiences Lodge")
idioma = col_lang.selectbox("Idioma", ["Português", "Español", "English"], label_visibility="collapsed")
if col_logout.button("Sair", use_container_width=True):
    st.session_state["agency_user"] = None
    st.session_state["agency_info"] = None
    st.rerun()

st.caption(f"Agência: **{agency_info['agency_name']}** · {agency_info['contact']}")
st.divider()

tipo = 1

date_fmt = "MM/DD/YYYY" if idioma == "English" else "DD/MM/YYYY"
col1, col2 = st.columns(2)
checkin  = col1.date_input("Check-in",  value=date.today() + timedelta(days=30), format=date_fmt)
checkout = col2.date_input("Check-out", value=checkin + timedelta(days=3), min_value=checkin + timedelta(days=3), format=date_fmt)
nights = max((checkout - checkin).days, 0)
st.caption(f"{nights} noite(s)")

st.divider()

with st.form("orcamento"):

    st.markdown("### Cliente")
    col1, col2 = st.columns(2)
    client_raw   = col1.text_input("Nome do cliente")
    email        = col1.text_input("Email", placeholder="cliente@email.com")
    telefone     = col2.text_input("Telefone", placeholder="+55 11 99999-9999")

    CAB_LABELS = {
        "COIGUE":    "COIGUE (até 2 adultos)",
        "NIRE":      "NIRE (até 2 adultos)",
        "CHAITEN":   "CHAITEN (até 2 adultos + 1 criança)",
        "CORCOVADO": "CORCOVADO (até 4 adultos ou 2 adultos + 3 crianças)",
    }

    st.markdown("### Cabanas")
    st.caption("Deixe Adultos = 0 para não incluir a cabana.")
    cabins_data = {}
    for cab in CABANAS:
        with st.expander(CAB_LABELS[cab], expanded=True):
            c1, c2 = st.columns(2)
            adults  = c1.number_input("Adultos", 0, 8, 0, key=f"ad_{cab}")
            infants = c2.number_input("Crianças", 0, 4, 0, key=f"inf_{cab}")
            if adults > 0:
                cabins_data[cab] = {"adults": int(adults), "infants": int(infants)}

    st.markdown("### Refeições")
    _month = checkin.month
    _meia_pensao = _month in (12, 1, 2, 3)
    if _meia_pensao:
        st.info(f"🍽️ **Meia pensão incluída** — café da manhã + jantar × {nights} noite(s) (Temporada Dez–Mar)")
    else:
        st.info(f"☕ **Café da manhã incluído** × {nights} noite(s) (Baixa temporada Abr–Nov)")

    st.markdown("### Comercial")
    agency      = st.checkbox("Com agência (15%)", value=True)
    adjustments = {}

    extras_sel = {}
    st.divider()
    submitted = st.form_submit_button("⚡ Gerar Orçamento", use_container_width=True, type="primary")

# ── Geração ────────────────────────────────────────────────────────────────────
if submitted:
    errors = []
    if not client_raw.strip():
        errors.append("Nome do cliente obrigatório.")
    if not cabins_data:
        errors.append("Selecione ao menos uma cabana.")
    if checkout <= checkin:
        errors.append("Check-out deve ser após check-in.")
    if 0 < (checkout - checkin).days < 3:
        errors.append("Mínimo de 3 noites.")

    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    lang_map = {"Português": "pt", "Español": "es", "English": "en"}
    lang = lang_map[idioma]

    _month = checkin.month
    meals = {}
    meals["breakfast"] = nights
    if _month in (12, 1, 2, 3):
        meals["dinner"] = nights

    total_adults  = sum(v["adults"]  for v in cabins_data.values())
    total_infants = sum(v["infants"] for v in cabins_data.values())

    client = client_raw.strip().upper().replace(" ", "")

    data = {
        "tipo":             tipo,
        "client_raw":       client_raw.strip(),
        "client":           client,
        "checkin":          checkin,
        "checkout":         checkout,
        "nights":           (checkout - checkin).days,
        "cabins":           cabins_data,
        "total_adults":     total_adults,
        "total_infants":    total_infants,
        "exchange":         m._fetch_exchange_rate(),
        "meals":            meals,
        "mapu_team":        0,
        "agency":           agency,
        "adjustments":      adjustments,
        "lang":             lang,
        "email_cliente":    email.strip(),
        "telefone_cliente": telefone.strip(),
        "extras_sel":       {},
        "agency_name":      agency_info["agency_name"],
        "agency_contact":   agency_info["contact"],
        "agency_email":     agency_info["email"],
        "agency_phone":     agency_info["phone"],
        "agency_user":      st.session_state["agency_user"],
    }

    with st.spinner("Gerando orçamento..."):
        try:
            prices = m.load_precos(checkin_year=checkin.year, checkin_month=checkin.month)
            calcs  = m.calculate(data, prices)

            for key in ("beverages", "activities", "transport", "experiences", "extras", "quincho"):
                calcs.setdefault(key, 0)

            client_path, slug = m.create_folders(data)
            prop_num  = m.next_prop_num(client_path)
            xlsx_path = m.populate_excel(data, client_path, slug, calcs, prices)

            # save totals into data so they're persisted in data.json for approval flow
            data["total_cc"]      = calcs.get("total_cc", 0)
            data["total"]         = calcs.get("total", 0)
            data["lodging"]       = calcs.get("lodging", 0)
            data["food"]          = calcs.get("food", 0)
            data["food_embedded"] = calcs.get("food_embedded", False)
            data["agency_fee"]    = calcs.get("agency_fee", 0)
            data["usd_ref"]       = calcs.get("usd_ref", 0)
            data["per_adult"]     = calcs.get("per_adult", 0)

            m._save_data_json(data, client_path, slug)

            pdf_path = m.generate_pdf(data, calcs, client_path, slug, prop_num)

            # Upload para Dropbox + notificação de aprovação para Gustavo
            if data.get("agency_user"):
                m.upload_budget_to_dropbox(pdf_path, xlsx_path, data)

            # Detecta URL do app (local vs nuvem)
            try:
                _host = st.context.headers.get("host", "localhost:8502")
                _scheme = "https" if "." in _host.split(":")[0] else "http"
                _app_url = f"{_scheme}://{_host}"
            except Exception:
                _app_url = "http://localhost:8502"

            m.send_approval_notification(data, calcs, pdf_path, app_url=_app_url)

            st.session_state["generated"]    = True
            st.session_state["last_pdf"]     = pdf_path
            st.session_state["last_data"]    = data
            st.session_state["last_calcs"]   = calcs
            st.session_state["last_slug"]    = slug
            st.session_state["last_email"]   = email.strip()

        except Exception as e:
            st.session_state["generated"] = False
            st.error(f"Erro ao gerar orçamento: {e}")
            st.exception(e)

# ── Resultados ─────────────────────────────────────────────────────────────────
if st.session_state.get("generated"):
    _pdf   = st.session_state["last_pdf"]
    _data  = st.session_state["last_data"]
    _calcs = st.session_state["last_calcs"]
    _slug  = st.session_state["last_slug"]

    st.success("✓ Orçamento enviado para aprovação MAPU!")
    st.info(
        "O orçamento foi gerado e enviado para Gustavo revisar. "
        "Assim que aprovado, você receberá o PDF por email."
    )

    st.divider()
    st.markdown("### Resumo")
    _hosp_lbl = "Hospedagem + Refeições" if _calcs.get("food_embedded") else "Hospedagem"
    st.metric(_hosp_lbl, f"CLP {_calcs['lodging']:,.0f}".replace(",", "."))
    if _calcs.get("food", 0) > 0:
        st.metric("Alimentação", f"CLP {_calcs['food']:,.0f}".replace(",", "."))
    st.metric("TOTAL (c/ 4% CC)", f"CLP {_calcs['total_cc']:,.0f}".replace(",", "."))
