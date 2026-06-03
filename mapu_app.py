#!/usr/bin/env python3
"""MAPU — Interface Web para Geração de Orçamentos v2"""

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

def load_agencies():
    """Carrega agências: arquivo local → Dropbox → Streamlit secrets (admin)."""
    agencies = {}

    # 1. Arquivo local (desenvolvimento)
    p = Path(__file__).parent / "agencies.json"
    if p.exists():
        agencies.update(json.loads(p.read_text()))

    # 2. Dropbox (fonte principal em produção)
    try:
        dropbox_agencies = m.load_agencies_dropbox()
        agencies.update(dropbox_agencies)
    except Exception:
        pass

    # 3. Streamlit secrets — admin + fallback
    try:
        for key in st.secrets.get("agencies", {}):
            entry = dict(st.secrets["agencies"][key])
            # admin do secrets tem prioridade e nunca é sobrescrito pelo Dropbox
            if entry.get("is_admin") or key not in agencies:
                agencies[key] = entry
    except Exception:
        pass

    return agencies

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

    # Lê dados direto da URL — sem precisar buscar no Dropbox
    p = st.query_params
    client_raw = p.get("cl", "—")
    checkin    = p.get("ci", "—")
    checkout   = p.get("co", "—")
    nights     = p.get("n", "—")
    cabins_str = p.get("cab", "—")
    total_cc   = int(p.get("tot", 0))
    ag_name    = p.get("ag", "—")
    ag_contact = p.get("agc", "—")
    ag_email   = p.get("age", "—")
    ag_phone   = p.get("agp", "—")
    ag_user    = p.get("agu", "")

    st.markdown(f"**Agência:** {ag_name} · {ag_contact} · {ag_email}")
    st.markdown(f"**Cliente:** {client_raw}")
    st.markdown(f"**Check-in:** {checkin} · **Check-out:** {checkout} · **{nights} noite(s)**")
    st.markdown(f"**Cabanas:** {cabins_str}")
    if total_cc:
        st.metric("TOTAL (c/ CC 4%)", f"CLP {total_cc:,.0f}".replace(",", "."))

    st.divider()

    col_a, col_b = st.columns(2)
    if col_a.button("✓ Aprovar e Enviar PDF", type="primary", use_container_width=True):
        with st.spinner("Baixando PDF e enviando..."):
            pdf_path = m.download_pdf_from_dropbox(_approve_key, ag_user)
            if pdf_path is None:
                st.warning("PDF não encontrado no Dropbox — email será enviado sem anexo.")
            data_min = {
                "client_raw":    client_raw,
                "agency_name":   ag_name,
                "agency_contact": ag_contact,
                "agency_email":  ag_email,
                "agency_phone":  ag_phone,
                "checkin":       checkin,
                "checkout":      checkout,
                "nights":        nights,
                "cabins":        {c: {} for c in cabins_str.split(",")},
            }
            calcs_min = {"total_cc": total_cc}
            ok = m.send_approved_budget_email(data_min, calcs_min, pdf_path)
        if ok:
            st.success(f"PDF enviado para {ag_email}")
        else:
            st.error("Erro ao enviar — verifique configuração SMTP.")

    if col_b.button("✗ Recusar", use_container_width=True):
        st.warning("Orçamento recusado. Nenhum email enviado.")

    st.stop()

# ── Tela de aprovação de nova agência (Gustavo) ────────────────────────────────
_register_key = st.query_params.get("register")
if _register_key == "1":
    st.title("MAPU LODGE | NOVO USUÁRIO - Aprovar Acesso")
    st.divider()
    p = st.query_params
    ag_name    = p.get("ag_name", "—")
    ag_contact = p.get("ag_contact", "—")
    ag_email   = p.get("ag_email", "—")
    ag_phone   = p.get("ag_phone", "—")

    st.markdown(f"**Agência:** {ag_name}")
    st.markdown(f"**Contato:** {ag_contact}")
    st.markdown(f"**Email:** {ag_email}")
    st.markdown(f"**Telefone:** {ag_phone}")
    st.divider()

    if st.button("✓ Aprovar e Criar Acesso", type="primary", use_container_width=True):
        with st.spinner("Criando credenciais e enviando email..."):
            existing = m.load_agencies_dropbox()
            username, password = m.generate_agency_credentials(ag_name, list(existing.keys()))
            existing[username] = {
                "password":    password,
                "agency_name": ag_name,
                "contact":     ag_contact,
                "email":       ag_email,
                "phone":       ag_phone,
            }
            ok_save = m.save_agencies_dropbox(existing)
            agency_data = {
                "agency_name": ag_name,
                "contact":     ag_contact,
                "email":       ag_email,
                "phone":       ag_phone,
            }
            ok_mail = m.send_agency_credentials(agency_data, username, password)
        if ok_save and ok_mail:
            st.success(f"Acesso criado! Login: `{username}` · Email enviado para {ag_email}")
        elif ok_save:
            st.warning(f"Acesso criado (login: `{username}`) mas erro ao enviar email para {ag_email}.")
        else:
            st.error("Erro ao salvar no Dropbox.")

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

    st.divider()
    with st.expander("Ainda não tem acesso? Solicitar acesso"):
        with st.form("registro"):
            r1, r2 = st.columns(2)
            r_name    = r1.text_input("Nome da agência*")
            r_contact = r1.text_input("Contato*")
            r_email   = r2.text_input("Email*")
            r_phone   = r2.text_input("Telefone")
            enviar = st.form_submit_button("Enviar solicitação", use_container_width=True, type="primary")
        if enviar:
            if not r_name.strip() or not r_contact.strip() or not r_email.strip():
                st.error("Preencha nome, contato e email.")
            else:
                try:
                    _host = st.context.headers.get("host", "localhost:8502")
                    _scheme = "https" if "." in _host.split(":")[0] else "http"
                    _app_url = f"{_scheme}://{_host}"
                except Exception:
                    _app_url = "http://localhost:8502"
                ok = m.send_registration_request(
                    {"agency_name": r_name.strip(), "contact": r_contact.strip(),
                     "email": r_email.strip(), "phone": r_phone.strip()},
                    _app_url,
                )
                if ok:
                    st.success("Solicitação enviada! Você receberá suas credenciais por email em breve.")
                else:
                    st.error("Erro ao enviar. Tente novamente ou contate hola@mapuchile.com")
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

# ── Painel Admin ───────────────────────────────────────────────────────────────
if agency_info.get("is_admin"):
    st.markdown("## Gerenciar Agências")
    agencies_now = m.load_agencies_dropbox()

    if agencies_now:
        for u, info in list(agencies_now.items()):
            with st.expander(f"**{info.get('agency_name','?')}** · `{u}`"):
                with st.form(f"edit_{u}"):
                    c1, c2 = st.columns(2)
                    new_name    = c1.text_input("Nome da agência", info.get("agency_name",""))
                    new_contact = c1.text_input("Contato", info.get("contact",""))
                    new_email   = c2.text_input("Email", info.get("email",""))
                    new_phone   = c2.text_input("Telefone", info.get("phone",""))
                    new_pass    = c2.text_input("Senha (deixe em branco para manter)", "")
                    col_save, col_del = st.columns(2)
                    salvar  = col_save.form_submit_button("💾 Salvar", use_container_width=True)
                    excluir = col_del.form_submit_button("🗑 Excluir", use_container_width=True)
                if salvar:
                    agencies_now[u]["agency_name"] = new_name
                    agencies_now[u]["contact"]     = new_contact
                    agencies_now[u]["email"]        = new_email
                    agencies_now[u]["phone"]        = new_phone
                    if new_pass.strip():
                        agencies_now[u]["password"] = new_pass.strip()
                    if m.save_agencies_dropbox(agencies_now):
                        st.success("Salvo!")
                    else:
                        st.error("Erro ao salvar no Dropbox.")
                if excluir:
                    del agencies_now[u]
                    if m.save_agencies_dropbox(agencies_now):
                        st.success(f"Agência `{u}` removida.")
                        st.rerun()
    else:
        st.info("Nenhuma agência cadastrada ainda.")

    st.divider()
    st.markdown("### Nova Agência")
    with st.form("nova_agencia"):
        c1, c2 = st.columns(2)
        n_user    = c1.text_input("Usuário (login)").strip().lower()
        n_pass    = c1.text_input("Senha")
        n_name    = c1.text_input("Nome da agência")
        n_contact = c1.text_input("Contato")
        n_email   = c2.text_input("Email")
        n_phone   = c2.text_input("Telefone")
        criar = st.form_submit_button("➕ Criar Agência", use_container_width=True, type="primary")

    if criar:
        if not n_user or not n_pass or not n_name:
            st.error("Preencha usuário, senha e nome.")
        elif n_user in agencies_now:
            st.error(f"Usuário `{n_user}` já existe.")
        else:
            agencies_now[n_user] = {
                "password":    n_pass,
                "agency_name": n_name,
                "contact":     n_contact,
                "email":       n_email,
                "phone":       n_phone,
            }
            try:
                ok = m.save_agencies_dropbox(agencies_now)
                if ok:
                    st.success(f"Agência `{n_name}` criada! Login: `{n_user}`")
                else:
                    st.error("save_agencies_dropbox retornou False")
            except Exception as ex:
                st.exception(ex)
    st.stop()

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
