#!/usr/bin/env python3
"""Gera PDF de guia de uso do app para agências — PT / EN / ES."""

from fpdf import FPDF
from pathlib import Path

LOGO  = Path(__file__).parent / "assets/logos/MAPU_logo_BADGEwhitesml.png"
FONTS = Path(__file__).parent / "assets/fonts"
OUT   = Path(__file__).parent / "MAPU_AgencyGuide.pdf"

APP_URL = "mapu-orcamentos.streamlit.app"

CABINS_PT = "COIGUE (2 ad.) · NIRE (2 ad.) · CHAITEN (2 ad.+1 cr.) · CORCOVADO (4 ad. ou 2 ad.+3 cr.)"
CABINS_EN = "COIGUE (2 ad.) · NIRE (2 ad.) · CHAITEN (2 ad.+1 ch.) · CORCOVADO (4 ad. or 2 ad.+3 ch.)"
CABINS_ES = "COIGUE (2 ad.) · NIRE (2 ad.) · CHAITEN (2 ad.+1 ni.) · CORCOVADO (4 ad. o 2 ad.+3 ni.)"

CONTENT = {
    "PT": {
        "lang_label": "PORTUGUÊS",
        "title":   "Como usar o Sistema de Orçamentos",
        "access":  "Acesso",
        "steps": [
            ("Acesse o app",       f"Abra {APP_URL} no seu navegador."),
            ("Faça login",         "Use o usuário e senha enviados pela equipe MAPU. Primeira vez? Clique em 'Solicitar acesso', preencha o formulário e aguarde suas credenciais por email."),
            ("Selecione o idioma", "Escolha PT / ES / EN no topo. O PDF do cliente será gerado no idioma selecionado."),
            ("Datas",              "Informe check-in e check-out. Mínimo de 3 noites."),
            ("Dados do cliente",   "Preencha nome, email e telefone."),
            ("Cabanas",            f"Informe adultos e crianças por cabana (Adultos = 0 exclui a cabana).\n{CABINS_PT}"),
            ("Refeições",          "Automaticas por temporada: Dez-Mar = meia pensao (cafe + jantar)  ·  Abr-Nov = cafe da manha"),
            ("Gere o orcamento",   "Clique em 'Gerar Orcamento'. O sistema calcula e envia para aprovacao MAPU."),
            ("Aguarde aprovacao",  "Apos aprovacao voce recebe o PDF finalizado por email para encaminhar ao cliente."),
        ],
        "notes_title": "Informacoes importantes",
        "notes": [
            "Comissao de agencia (15%) aplicada automaticamente.",
            "Estadas >= 7 noites: desconto de 5% na hospedagem.",
            "Total inclui taxa de cartao de credito (4%).",
            "Duvidas: hola@mapuchile.com  ·  +569 58642354",
        ],
    },
    "EN": {
        "lang_label": "ENGLISH",
        "title":   "How to Use the Estimates System",
        "access":  "Access",
        "steps": [
            ("Access the app",        f"Open {APP_URL} in your browser."),
            ("Log in",                "Use the username and password provided by the MAPU team. First time? Click 'Solicitar acesso', fill out the form and wait for your credentials by email."),
            ("Select the language",   "Choose PT / ES / EN at the top. The client PDF will be generated in the selected language."),
            ("Dates",                 "Enter check-in and check-out. Minimum stay: 3 nights."),
            ("Client details",        "Fill in the client's name, email and phone number."),
            ("Cabins",                f"Enter adults and children per cabin (Adults = 0 excludes the cabin).\n{CABINS_EN}"),
            ("Meals",                 "Automatic by season: Dec-Mar = half board (breakfast + dinner)  ·  Apr-Nov = breakfast only"),
            ("Generate the estimate", "Click 'Gerar Orcamento'. The system calculates and sends for MAPU approval."),
            ("Wait for approval",     "Once approved you receive the final PDF by email to forward to your client."),
        ],
        "notes_title": "Important information",
        "notes": [
            "Agency commission (15%) applied automatically.",
            "Stays >= 7 nights: 5% cabin discount.",
            "Total includes credit card fee (4%).",
            "Questions: hola@mapuchile.com  ·  +569 58642354",
        ],
    },
    "ES": {
        "lang_label": "ESPAÑOL",
        "title":   "Como usar el Sistema de Presupuestos",
        "access":  "Acceso",
        "steps": [
            ("Accede al app",          f"Abre {APP_URL} en tu navegador."),
            ("Inicia sesion",          "Usa el usuario y contrasena proporcionados por el equipo MAPU. Primera vez? Haz clic en 'Solicitar acesso', completa el formulario y espera tus credenciales por email."),
            ("Selecciona el idioma",   "Elige PT / ES / EN en la parte superior. El PDF del cliente se generara en el idioma seleccionado."),
            ("Fechas",                 "Ingresa check-in y check-out. Minimo 3 noches."),
            ("Datos del cliente",      "Completa nombre, email y telefono."),
            ("Cabanas",                f"Ingresa adultos y ninos por cabana (Adultos = 0 excluye la cabana).\n{CABINS_ES}"),
            ("Comidas",                "Automaticas por temporada: Dic-Mar = media pension (desayuno + cena)  ·  Abr-Nov = solo desayuno"),
            ("Genera el presupuesto",  "Haz clic en 'Gerar Orcamento'. El sistema calcula y envia para aprobacion MAPU."),
            ("Espera la aprobacion",   "Tras la aprobacion recibiras el PDF final por email para reenviar a tu cliente."),
        ],
        "notes_title": "Informacion importante",
        "notes": [
            "Comision de agencia (15%) aplicada automaticamente.",
            "Estadias >= 7 noches: 5% de descuento en hospedaje.",
            "El total incluye tarifa de tarjeta de credito (4%).",
            "Consultas: hola@mapuchile.com  ·  +569 58642354",
        ],
    },
}

# ── Layout constants ────────────────────────────────────────────────────────────
HDR_H     = 28     # header height
STEP_NUM  = 7      # step number box size
STEP_GAP  = 1.5    # gap after each step
LH        = 4.2    # line height body
FS_TITLE  = 11     # step title font size
FS_BODY   = 9      # step body font size
FS_NOTES  = 8.5    # notes font size
LH_NOTES  = 4.0    # notes line height


class GuidePDF(FPDF):
    def __init__(self):
        super().__init__()
        self.add_font("DIN", "",  str(FONTS / "FF_DIN_Condensed_Regular.otf"))
        self.add_font("DIN", "B", str(FONTS / "FF_DIN_Condensed_Bold.otf"))
        self.set_auto_page_break(auto=True, margin=14)

    def header_bar(self, lang_label):
        self.set_fill_color(26, 26, 26)
        self.rect(0, 0, 210, HDR_H, "F")
        if LOGO.exists():
            self.image(str(LOGO), x=9, y=4, h=20)
        self.set_xy(38, 6)
        self.set_font("DIN", "B", 14)
        self.set_text_color(255, 255, 255)
        self.cell(160, 6, "MAPU LODGE  |  PATAGONIA  |  CHILE")
        self.set_xy(38, 13)
        self.set_font("DIN", "", 10)
        self.set_text_color(160, 160, 160)
        self.cell(100, 5, "AGENCY GUIDE")
        # pill idioma
        self.set_xy(155, 9)
        self.set_fill_color(55, 55, 55)
        self.set_text_color(200, 200, 200)
        self.set_font("DIN", "B", 8)
        self.cell(44, 7, lang_label, align="C", fill=True)
        self.set_text_color(0)
        self.set_y(HDR_H + 4)

    def footer(self):
        self.set_y(-12)
        self.set_font("DIN", "", 7)
        self.set_text_color(190, 190, 190)
        self.cell(0, 4, f"{APP_URL}  ·  hola@mapuchile.com  ·  +569 58642354", align="C", link=f"https://{APP_URL}")

    def draw_step(self, num, title, body):
        y0 = self.get_y()
        # número
        self.set_fill_color(26, 26, 26)
        self.set_text_color(255, 255, 255)
        self.set_font("DIN", "B", 9)
        self.set_xy(10, y0)
        self.cell(STEP_NUM, STEP_NUM, str(num), align="C", fill=True)
        # título
        self.set_xy(20, y0 + 0.5)
        self.set_text_color(26, 26, 26)
        self.set_font("DIN", "B", FS_TITLE)
        self.cell(0, STEP_NUM - 1, title)
        # corpo
        self.set_xy(20, y0 + STEP_NUM)
        self.set_font("DIN", "", FS_BODY)
        self.set_text_color(75, 75, 75)
        self.multi_cell(177, LH, body)
        self.ln(STEP_GAP)

    def draw_notes(self, title, notes):
        self.set_fill_color(245, 245, 245)
        self.set_draw_color(230, 230, 230)
        y = self.get_y() + 8
        self.set_xy(10, y)
        # linha topo
        self.set_draw_color(200, 200, 200)
        self.line(10, y, 200, y)
        self.ln(4)
        self.set_x(10)
        self.set_font("DIN", "B", 9)
        self.set_text_color(26, 26, 26)
        self.cell(0, 5, title)
        self.ln(7)
        for note in notes:
            self.set_x(12)
            self.set_font("DIN", "", FS_NOTES)
            self.set_text_color(80, 80, 80)
            self.multi_cell(184, LH_NOTES, f"·  {note}")


def build():
    pdf = GuidePDF()
    pdf.set_margins(10, 10, 10)

    for lang in ("PT", "EN", "ES"):
        c = CONTENT[lang]
        pdf.add_page()
        pdf.header_bar(c["lang_label"])

        # Título + URL
        pdf.set_font("DIN", "B", 16)
        pdf.set_text_color(26, 26, 26)
        pdf.cell(0, 8, c["title"], ln=True)

        pdf.set_fill_color(26, 26, 26)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("DIN", "B", 9)
        pdf.set_x(10)
        pdf.cell(187, 7, f"  {c['access']}: https://{APP_URL}", fill=True, ln=True, link=f"https://{APP_URL}")
        pdf.ln(4)

        # Passos
        for i, (step_title, step_body) in enumerate(c["steps"], 1):
            pdf.draw_step(i, step_title, step_body)

        # Notas
        pdf.draw_notes(c["notes_title"], c["notes"])

    pdf.output(str(OUT))
    print(f"✓ PDF gerado: {OUT}")


if __name__ == "__main__":
    build()
