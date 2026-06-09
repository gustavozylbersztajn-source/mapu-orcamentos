#!/usr/bin/env python3
"""Gera PDF de guia admin MAPU — só PT."""

from fpdf import FPDF
from pathlib import Path

LOGO  = Path(__file__).parent / "assets/logos/MAPU_logo_BADGEwhitesml.png"
FONTS = Path(__file__).parent / "assets/fonts"
OUT   = Path(__file__).parent / "MAPU_AdminGuide.pdf"

APP_URL = "mapu-orcamentos.streamlit.app"

HDR_H    = 26
STEP_NUM = 6
STEP_GAP = 0.8
LH       = 3.8
FS_TITLE = 10
FS_BODY  = 8.5
FS_NOTES = 8
LH_NOTES = 3.6

SECTIONS = [
    {
        "title": "Acesso Admin",
        "steps": [
            ("Acesse o app",        f"Abra https://{APP_URL} no navegador."),
            ("Faça login (admin)",  "Usuário: mapu  ·  Senha: Lonco2026@\nO painel de administração abre automaticamente após o login."),
            ("Gere seu orçamento",  "No final do painel admin, clique em 'Gerar Orcamento' para acessar o formulário.\nUse 'Voltar ao Painel Admin' para retornar ao gerenciamento."),
        ],
    },
    {
        "title": "Acesso Interno Rápido (Gustavo)",
        "steps": [
            ("Login sem senha",     "Usuário: guga  ·  Senha: deixar em branco e clicar Entrar\nVai direto ao formulário de orçamento, sem passar pelo painel admin."),
            ("Sem mínimo de noites", "O usuário guga não tem restrição de noites mínimas — útil para orçamentos especiais de 1 ou 2 noites.\nAgências continuam com mínimo de 3 noites."),
        ],
    },
    {
        "title": "Aprovar Orçamento de Agência",
        "steps": [
            ("Receba o email",   "Quando uma agência gera um orçamento, você recebe um email em hola@mapuchile.com com o assunto '[MAPU] Aprovação — Cliente · Agência'."),
            ("Revise os dados", "O email mostra: cliente, check-in, check-out, cabanas, número de hóspedes e total (c/ CC 4%)."),
            ("Aprove",          "Clique em 'APROVAR E ENVIAR PDF'. A agência recebe automaticamente o PDF aprovado por email."),
            ("Recuse",          "Se necessário, clique em 'Recusar'. Nenhum email é enviado à agência — entre em contato manualmente."),
        ],
    },
    {
        "title": "Aprovar Nova Agência (Auto-registro)",
        "steps": [
            ("Receba a solicitação", "Quando uma agência preenche o formulário 'Solicitar acesso', você recebe um email com assunto 'NOVO USUÁRIO | ACESSO APP ORÇAMENTOS MAPU LODGE'."),
            ("Revise os dados",      "O email mostra: nome da agência, contato, email e telefone."),
            ("Aprove",               "Clique em 'Aprovar e Criar Acesso'. O sistema gera login e senha automaticamente e envia as credenciais para a agência por email."),
        ],
    },
    {
        "title": "Gerenciar Agências (Painel Admin)",
        "steps": [
            ("Acesse o painel", "Após o login como admin (mapu), o painel de gerenciamento aparece automaticamente."),
            ("Editar agência",  "Clique em uma agência para expandir. Altere nome, contato, email, telefone ou senha e clique 'Salvar'."),
            ("Excluir agência", "Dentro do card da agência, clique 'Excluir'. A agência perde o acesso imediatamente."),
            ("Criar agência",   "Use o formulário 'Nova Agência' no final do painel. Defina usuário, senha e dados manualmente."),
        ],
    },
]

NOTES_TITLE = "Informações do sistema"
NOTES = [
    f"App: https://{APP_URL}",
    "Agências armazenadas no Dropbox: /config/agencies.json",
    "Emails enviados via hola@mapuchile.com (Hostinger SMTP)",
    "Orçamentos salvos no Dropbox: /ANO/MES/DATA_CLIENTE_AGENCIA/",
    "Credenciais e configurações: Streamlit Cloud > Manage app > Settings > Secrets",
]


class AdminPDF(FPDF):
    def __init__(self):
        super().__init__()
        self.add_font("DIN", "",  str(FONTS / "FF_DIN_Condensed_Regular.otf"))
        self.add_font("DIN", "B", str(FONTS / "FF_DIN_Condensed_Bold.otf"))
        self.set_auto_page_break(auto=True, margin=14)

    def header_bar(self):
        self.set_fill_color(26, 26, 26)
        self.rect(0, 0, 210, HDR_H, "F")
        if LOGO.exists():
            self.image(str(LOGO), x=9, y=4, h=20)
        self.set_xy(38, 6)
        self.set_font("DIN", "B", 14)
        self.set_text_color(255, 255, 255)
        self.cell(120, 6, "MAPU LODGE  |  PATAGONIA  |  CHILE")
        self.set_xy(38, 13)
        self.set_font("DIN", "", 10)
        self.set_text_color(160, 160, 160)
        self.cell(100, 5, "ADMIN GUIDE")
        self.set_xy(155, 9)
        self.set_fill_color(55, 55, 55)
        self.set_text_color(200, 200, 200)
        self.set_font("DIN", "B", 8)
        self.cell(44, 7, "PORTUGUÊS", align="C", fill=True)
        self.set_text_color(0)
        self.set_y(HDR_H + 4)

    def footer(self):
        self.set_y(-12)
        self.set_font("DIN", "", 7)
        self.set_text_color(190, 190, 190)
        self.cell(0, 4, f"{APP_URL}  ·  hola@mapuchile.com  ·  +569 58642354", align="C")

    def section_title(self, title):
        self.ln(1.5)
        self.set_fill_color(26, 26, 26)
        self.set_text_color(255, 255, 255)
        self.set_font("DIN", "B", 9)
        self.set_x(10)
        self.cell(187, 6, f"  {title}", fill=True, ln=True)
        self.ln(2)

    def draw_step(self, num, title, body):
        y0 = self.get_y()
        self.set_fill_color(80, 80, 80)
        self.set_text_color(255, 255, 255)
        self.set_font("DIN", "B", 9)
        self.set_xy(10, y0)
        self.cell(STEP_NUM, STEP_NUM, str(num), align="C", fill=True)
        self.set_xy(20, y0 + 0.5)
        self.set_text_color(26, 26, 26)
        self.set_font("DIN", "B", FS_TITLE)
        self.cell(0, STEP_NUM - 1, title)
        self.set_xy(20, y0 + STEP_NUM)
        self.set_font("DIN", "", FS_BODY)
        self.set_text_color(75, 75, 75)
        self.multi_cell(177, LH, body)
        self.ln(STEP_GAP)

    def draw_notes(self, title, notes):
        y = self.get_y() + 3
        self.set_xy(10, y)
        self.set_draw_color(200, 200, 200)
        self.line(10, y, 200, y)
        self.ln(4)
        self.set_x(10)
        self.set_font("DIN", "B", 9)
        self.set_text_color(26, 26, 26)
        self.cell(0, 5, title)
        self.ln(5)
        for note in notes:
            self.set_x(12)
            self.set_font("DIN", "", FS_NOTES)
            self.set_text_color(80, 80, 80)
            self.multi_cell(184, LH_NOTES, f"·  {note}")


def build():
    pdf = AdminPDF()
    pdf.set_margins(10, 10, 10)
    pdf.add_page()
    pdf.header_bar()

    # Título
    pdf.set_font("DIN", "B", 14)
    pdf.set_text_color(26, 26, 26)
    pdf.cell(0, 6, "Guia do Administrador", ln=True)
    pdf.ln(1)

    step_counter = 1
    for section in SECTIONS:
        pdf.section_title(section["title"])
        for title, body in section["steps"]:
            pdf.draw_step(step_counter, title, body)
            step_counter += 1

    pdf.draw_notes(NOTES_TITLE, NOTES)
    pdf.output(str(OUT))
    print(f"✓ PDF gerado: {OUT}")


if __name__ == "__main__":
    build()
