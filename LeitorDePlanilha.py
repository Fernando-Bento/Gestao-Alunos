import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
from reportlab.platypus import Paragraph, Spacer, Image, Table, TableStyle, PageBreak, BaseDocTemplate, Frame, PageTemplate, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from datetime import datetime, timedelta
import os, sys
from PIL import Image as PILImage

COR = "#FF6600"

def get_path(rel):
    base = getattr(sys, "frozen", False) and sys._MEIPASS or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, rel)

def ultima_semana_sexta_a_sexta(df):
    ultima = df["Data"].max()
    ultima_sexta = ultima - timedelta(days=(ultima.weekday() - 4) % 7)
    inicio = ultima_sexta - timedelta(days=6)
    return inicio, ultima_sexta

def safe_mode(s):
    try:
        return s.mode().iloc[0]
    except Exception:
        return "N/A"

def gerar_relatorio():
    arq = entry_arquivo.get()
    if not arq:
        messagebox.showwarning("Erro", "Selecione a planilha")
        return
    try:
        df = pd.read_excel(arq)
        df["Data"] = pd.to_datetime(df["Data"])

        periodo = var_periodo.get()
        if periodo == "Mensal":
            m = df["Data"].dt.month.max(); y = df["Data"].dt.year.max()
            dfp = df[(df["Data"].dt.month==m) & (df["Data"].dt.year==y)].copy()
            titulo = f"{m}/{y}"; slug = f"{m}_{y}"
        else:
            inicio, fim = ultima_semana_sexta_a_sexta(df)
            dfp = df[(df["Data"]>=inicio) & (df["Data"]<=fim)].copy()
            titulo = f"Semana de {inicio.strftime('%d/%m')} a {fim.strftime('%d/%m/%Y')}"
            slug = f"{inicio.strftime('%Y%m%d')}_{fim.strftime('%Y%m%d')}"

        if dfp.empty:
            messagebox.showwarning("Aviso", "Período sem dados.")
            return

        dfp["Dia"] = dfp["Data"].dt.day
        matriculados = int(dfp[dfp["Inscrito"]==1].shape[0])
        experimentais = int(dfp["Experimental"].sum())
        exp_conv = int(dfp[(dfp["Experimental"]==1) & (dfp["Inscrito"]==1)].shape[0])
        media_ins = dfp.groupby("Dia")["Inscrito"].sum().mean()
        pagamento_pred = safe_mode(dfp[dfp["Inscrito"]==1]["Pagamento"])
        plano_pred = safe_mode(dfp[dfp["Inscrito"]==1]["Plano"])
        try:
            renov = int(entry_renov.get() or 0)
        except Exception:
            renov = 0

        os.makedirs(get_path("graficos"), exist_ok=True)
        os.makedirs(get_path("Relatorios"), exist_ok=True)

        inscritos_dia = dfp.groupby("Dia")["Inscrito"].sum().astype(int)
        plt.figure(figsize=(7,2.6))
        ax = inscritos_dia.plot(kind="bar", color=COR)
        ax.yaxis.set_major_locator(MaxNLocator(integer=True))
        plt.xlabel("Dia"); plt.ylabel("Inscritos"); plt.tight_layout()
        plt.savefig(get_path("graficos/inscritos_dia.png"), bbox_inches="tight"); plt.close()

        pagamento = dfp.groupby("Pagamento")["Inscrito"].sum()
        plt.figure(figsize=(3.2,3.2))
        pagamento.plot(kind="pie", autopct='%1.1f%%', startangle=140,
                       colors=[COR,"#FFA500","#FFB266","#FFCC99"])
        plt.ylabel(""); plt.tight_layout()
        plt.savefig(get_path("graficos/pagamento.png"), bbox_inches="tight"); plt.close()

        def rodape(canvas, doc):
            if doc.page == 2:
                canvas.saveState()
                txt = f"{datetime.now().strftime('%d/%m/%Y %H:%M')} - Raquel Camille da Silva"
                canvas.setFont("Helvetica", 8)
                canvas.drawRightString(A4[0]-40, 20, txt)
                canvas.restoreState()

        pdf_path = get_path(f"Relatorios/relatorio_{slug}.pdf")
        doc = BaseDocTemplate(pdf_path, pagesize=A4)
        frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="f")
        doc.addPageTemplates([PageTemplate(id="p", frames=frame, onPage=rodape)])

        styles = getSampleStyleSheet()
        title_s = ParagraphStyle("title", parent=styles["Title"], alignment=1, textColor=colors.HexColor(COR))
        h_s = ParagraphStyle("h", parent=styles["Heading2"], textColor=colors.HexColor(COR))
        normal = styles["Normal"]

        story = []
        story.append(Paragraph(f"Relatório ({titulo})", title_s))
        story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor(COR)))
        story.append(Spacer(1,10))

        resumo = (f"No período {titulo}, tivemos <b>{matriculados} alunos que assinaram o plano</b>. "
                  f"Foram realizadas <b>{experimentais} aulas experimentais</b>, das quais <b>{exp_conv} resultaram em matrícula</b>. "
                  f"A forma de pagamento predominante foi <b>{pagamento_pred}</b> e o plano mais escolhido foi <b>{plano_pred}</b>.")
        story.append(Paragraph("Resumo", h_s)); story.append(Paragraph(resumo, normal)); story.append(Spacer(1,8))

        story.append(Paragraph("Métricas Principais", h_s))
        for txt in [
            f"Total de alunos matriculados: {matriculados}",
            f"Aulas experimentais realizadas: {experimentais}",
            f"Aulas experimentais com conversão: {exp_conv}",
            f"Renovações de contratos: {renov}",
            f"Média diária de inscritos: {media_ins:.0f}",
            f"Forma de pagamento predominante: {pagamento_pred}",
            f"Plano mais escolhido: {plano_pred}"
        ]:
            story.append(Paragraph(txt, normal))
        story.append(Spacer(1,8))

        def add_img(path, titulo, w=360):
            img = PILImage.open(path); r = img.height / img.width
            story.append(Paragraph(titulo, h_s))
            story.append(Image(path, width=w, height=w * r))
            story.append(Spacer(1,6))

        add_img(get_path("graficos/inscritos_dia.png"), "Inscrições por Dia", w=360)
        add_img(get_path("graficos/pagamento.png"), "Distribuição por Pagamento", w=170)

        story.append(PageBreak())
        story.append(Paragraph("Lista de novos alunos", ParagraphStyle("c", parent=h_s, alignment=1)))
        tabela = dfp[dfp["Inscrito"]==1][["Aluno","Plano","Pagamento"]]
        dados = [tabela.columns.tolist()] + tabela.values.tolist()
        t = Table(dados, hAlign="CENTER")
        t.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.5,colors.HexColor(COR)),
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#FFCC99")),
            ("ALIGN",(0,1),(-1,-1),"CENTER")
        ]))
        story.append(t)

        doc.build(story)
        messagebox.showinfo("Sucesso", f"Relatório gerado: {pdf_path}")

    except Exception as e:
        messagebox.showerror("Erro", str(e))

# Interface
root = tk.Tk(); root.title("Gerador de Relatório"); root.geometry("520x260")
tk.Label(root, text="Planilha (.xlsx):").pack(pady=6)
entry_arquivo = tk.Entry(root, width=66); entry_arquivo.pack()
def selecionar_arquivo():
    f = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])
    if f:
        entry_arquivo.delete(0, tk.END); entry_arquivo.insert(0, f)
tk.Button(root, text="Selecionar arquivo", command=selecionar_arquivo).pack(pady=6)
tk.Label(root, text="Período:").pack()
var_periodo = tk.StringVar(value="Mensal")
tk.Radiobutton(root, text="Mensal", variable=var_periodo, value="Mensal").pack()
tk.Radiobutton(root, text="Semanal (sexta a sexta)", variable=var_periodo, value="Semanal").pack()
tk.Label(root, text="Renovações (manual):").pack(pady=4)
entry_renov = tk.Entry(root, width=10); entry_renov.insert(0,"0"); entry_renov.pack()
tk.Button(root, text="Gerar Relatório", width=20, command=gerar_relatorio).pack(pady=12)
root.mainloop()
