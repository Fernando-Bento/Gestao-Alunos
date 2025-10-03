import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
from reportlab.platypus import Paragraph, Spacer, Image, Table, TableStyle, PageBreak, BaseDocTemplate, Frame, PageTemplate, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from datetime import datetime, timedelta
from pathlib import Path
from PIL import Image as PILImage
import os, sys

COR = "#FF6600"

REQUIRED_COLS = {"Data", "Aluno", "Inscrito", "Experimental", "Plano", "Pagamento"}

def safe_mode(s):
    try:
        return s.mode().iloc[0]
    except Exception:
        return "N/A"

def ultima_semana_sexta_a_sexta(df):
    ultima = df["Data"].max()
    ultima_sexta = ultima - timedelta(days=(ultima.weekday() - 4) % 7)
    inicio = ultima_sexta - timedelta(days=6)
    return inicio, ultima_sexta

def gerar_relatorio():
    arquivo = entry_arquivo.get().strip()
    if not arquivo:
        messagebox.showwarning("Erro", "Selecione a planilha (.xlsx).")
        return

    try:
        # --------- Ler arquivo selecionado (sempre) ----------
        df = pd.read_excel(arquivo)
        if df.empty:
            messagebox.showwarning("Erro", "Planilha vazia.")
            return
        # valida colunas
        faltantes = REQUIRED_COLS - set(df.columns)
        if faltantes:
            messagebox.showerror("Erro", f"Colunas faltando na planilha: {', '.join(faltantes)}")
            return

        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
        if df["Data"].isna().all():
            messagebox.showerror("Erro", "Coluna 'Data' não contém datas válidas.")
            return

        # --------- escolher período ----------
        periodo = var_periodo.get()
        if periodo == "Mensal":
            m = df["Data"].dt.month.max()
            y = df["Data"].dt.year.max()
            dfp = df[(df["Data"].dt.month == m) & (df["Data"].dt.year == y)].copy()
            titulo = f"{m:02d}/{y}"
            slug = f"{y}{m:02d}"
        else:  # semanal sexta->sexta com base na última data do arquivo atual
            inicio, fim = ultima_semana_sexta_a_sexta(df)
            dfp = df[(df["Data"] >= inicio) & (df["Data"] <= fim)].copy()
            titulo = f"Semana de {inicio.strftime('%d/%m')} a {fim.strftime('%d/%m/%Y')}"
            slug = f"{inicio.strftime('%Y%m%d')}_{fim.strftime('%Y%m%d')}"

        if dfp.empty:
            messagebox.showwarning("Aviso", "Período selecionado não contém dados.")
            return

        # --------- métricas ----------
        dfp["Dia"] = dfp["Data"].dt.day
        total_matriculados = int(dfp[dfp["Inscrito"] == 1].shape[0])
        total_experimentais = int(dfp["Experimental"].sum())
        exp_com_conv = int(dfp[(dfp["Experimental"] == 1) & (dfp["Inscrito"] == 1)].shape[0])
        media_diaria = dfp.groupby("Dia")["Inscrito"].sum().mean()
        pagamento_pred = safe_mode(dfp[dfp["Inscrito"] == 1]["Pagamento"])
        plano_pred = safe_mode(dfp[dfp["Inscrito"] == 1]["Plano"])
        try:
            renov = int(entry_renov.get() or 0)
        except Exception:
            renov = 0

        # --------- criar pastas: Desktop/Relatorios/<slug>/Graficos ----------
        base = Path.home() / "Desktop" / "Relatorios" / slug
        graf_dir = base / "Graficos"
        base.mkdir(parents=True, exist_ok=True)
        graf_dir.mkdir(parents=True, exist_ok=True)

        # --------- gráficos com nomes únicos (por slug) ----------
        graf_inscr = graf_dir / f"inscritos_dia_{slug}.png"
        inscritos_dia = dfp.groupby("Dia")["Inscrito"].sum().astype(int)
        plt.figure(figsize=(7, 2.6))
        ax = inscritos_dia.plot(kind="bar", color=COR)
        ax.yaxis.set_major_locator(MaxNLocator(integer=True))
        plt.xlabel("Dia"); plt.ylabel("Inscritos"); plt.tight_layout()
        plt.savefig(str(graf_inscr), bbox_inches="tight"); plt.close()

        graf_pag = graf_dir / f"pagamento_{slug}.png"
        pagamento = dfp.groupby("Pagamento")["Inscrito"].sum()
        plt.figure(figsize=(3.2, 3.2))
        pagamento.plot(kind="pie", autopct='%1.1f%%', startangle=140,
                       colors=[COR, "#FFA500", "#FFB266", "#FFCC99"])
        plt.ylabel(""); plt.tight_layout()
        plt.savefig(str(graf_pag), bbox_inches="tight"); plt.close()

        # --------- PDF ----------
        pdf_path = base / f"relatorio_{slug}.pdf"
        pdf_path_str = str(pdf_path)

        def rodape(canvas, doc):
            canvas.saveState()
            txt = f"{datetime.now().strftime('%d/%m/%Y %H:%M')} - Raquel Camille da Silva"
            canvas.setFont("Helvetica", 8)
            canvas.drawRightString(A4[0] - 20 * mm, 10 * mm, txt)
            canvas.restoreState()

        doc = BaseDocTemplate(pdf_path_str, pagesize=A4)
        frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="f")
        doc.addPageTemplates([PageTemplate(id="p", frames=frame, onPage=rodape)])

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle("title", parent=styles["Title"], alignment=1, textColor=colors.HexColor(COR))
        hstyle = ParagraphStyle("h", parent=styles["Heading2"], textColor=colors.HexColor(COR))
        normal = styles["Normal"]

        story = []
        story.append(Paragraph(f"Relatório ({titulo})", title_style))
        story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor(COR)))
        story.append(Spacer(1, 8))

        resumo = (
            f"No período {titulo}, tivemos <b>{total_matriculados} alunos que assinaram o plano</b>. "
            f"Foram realizadas <b>{total_experimentais} aulas experimentais</b>, das quais <b>{exp_com_conv} resultaram em matrícula</b>. "
            f"Renovações (manual): <b>{renov}</b>. Forma de pagamento predominante: <b>{pagamento_pred}</b>. Plano mais escolhido: <b>{plano_pred}</b>."
        )
        story.append(Paragraph("Resumo do Período", hstyle))
        story.append(Paragraph(resumo, normal))
        story.append(Spacer(1, 8))

        story.append(Paragraph("Métricas Principais", hstyle))
        for linha in [
            f"Total de alunos matriculados: {total_matriculados}",
            f"Aulas experimentais realizadas: {total_experimentais}",
            f"Aulas experimentais com conversão: {exp_com_conv}",
            f"Renovações de contratos (manual): {renov}",
            f"Média diária de inscritos: {media_diaria:.0f}"
        ]:
            story.append(Paragraph(linha, normal))
        story.append(Spacer(1, 8))

        def add_img(path_obj, titulo, width=360):
            path_s = str(path_obj)
            img = PILImage.open(path_s)
            ratio = img.height / img.width
            story.append(Paragraph(titulo, hstyle))
            story.append(Image(path_s, width=width, height=width * ratio))
            story.append(Spacer(1, 6))

        add_img(graf_inscr, "Inscrições por Dia", width=360)
        add_img(graf_pag, "Distribuição por Pagamento", width=170)

        story.append(PageBreak())
        story.append(Paragraph("Lista de novos alunos", ParagraphStyle("c", parent=hstyle, alignment=1)))
        tabela = dfp[dfp["Inscrito"] == 1][["Aluno", "Plano", "Pagamento"]]
        dados = [tabela.columns.tolist()] + tabela.values.tolist()
        tabela_report = Table(dados, hAlign="CENTER")
        tabela_report.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor(COR)),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#FFCC99")),
            ("ALIGN", (0, 1), (-1, -1), "CENTER"),
        ]))
        story.append(tabela_report)

        doc.build(story)

        messagebox.showinfo("Sucesso", f"Relatório gerado:\n{pdf_path_str}")

    except Exception as ex:
        messagebox.showerror("Erro", str(ex))


# ---------- Interface ----------
root = tk.Tk()
root.title("Gerador de Relatório")
root.geometry("560x300")

tk.Label(root, text="Planilha (.xlsx):").pack(pady=6)
entry_arquivo = tk.Entry(root, width=72)
entry_arquivo.pack()

def selecionar_arquivo():
    f = filedialog.askopenfilename(filetypes=[("Planilha Excel", "*.xlsx;*.xls")])
    if f:
        entry_arquivo.delete(0, tk.END)
        entry_arquivo.insert(0, f)

tk.Button(root, text="Selecionar arquivo", command=selecionar_arquivo).pack(pady=6)

tk.Label(root, text="Período:").pack()
var_periodo = tk.StringVar(value="Mensal")
tk.Radiobutton(root, text="Mensal", variable=var_periodo, value="Mensal").pack()
tk.Radiobutton(root, text="Semanal (sexta a sexta)", variable=var_periodo, value="Semanal").pack()

tk.Label(root, text="Renovações (manual):").pack(pady=4)
entry_renov = tk.Entry(root, width=10)
entry_renov.insert(0, "0")
entry_renov.pack()

tk.Button(root, text="Gerar Relatório", width=20, command=gerar_relatorio).pack(pady=12)

root.mainloop()
