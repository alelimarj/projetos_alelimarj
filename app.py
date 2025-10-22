# ============================================================
# app.py — Protocolo Prisma ver. 0.7i
# ============================================================

import streamlit as st
import pandas as pd
import csv
import re
import io
import os
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
import plotly.express as px

st.set_page_config(page_title="Protocolo Prisma ver. 0.7i", layout="wide")
st.title("🧾 Protocolo Prisma — ver. 0.7i")
st.caption("Protocolo para conversão de arquivo .txt em Excel, com gráfico de barras 3D exportado, tabela-resumo com totais e formatação condicional, e gráfico de pizza interativo.")

# ----------------------- Funções auxiliares -----------------------


def parse_number_us(s):
    if s is None:
        return None
    s = str(s).strip().strip('"')
    if s == "":
        return None
    s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def br_format(n):
    if n is None or n == "":
        return ""
    s = f"{n:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def extract_between(text, start_label, end_label):
    p = text.find(start_label)
    if p == -1:
        return ""
    p += len(start_label)
    q = text.find(end_label, p)
    return text[p:q].strip() if q != -1 else text[p:].strip()


def extract_plano(text):
    m = re.search(r'Plano\s*:\s*"?([^"]*?)"?$', text.strip())
    if m and m.group(1):
        return m.group(1).strip()
    if "Plano:" in text:
        return text.split("Plano:")[-1].strip().strip('"')
    return ""


def detect_periodo_first_lines(raw_text: str) -> str:
    lines = raw_text.splitlines()[:3]
    header = " ".join(lines)
    m = re.search(
        r'[",]*Período\s*:[" ,]*([0-3]\d/[0-1]\d/\d{4}\s*a\s*[0-3]\d/[0-1]\d/\d{4})', header)
    return m.group(1).strip() if m else ""

# ----------------------- Núcleo de processamento -----------------------


def process_txt_content(txt: str) -> pd.DataFrame:
    raw = txt.replace(",Setor:,", ",")
    lines_all = raw.splitlines()
    default_periodo = detect_periodo_first_lines(raw)
    current_setor = ""
    current_periodo = default_periodo
    records = []
    i = 0

    m_data = re.search(r"Data:\s*,?\s*([0-3]?\d/[0-1]?\d/\d{4})", txt)
    data_extracao = m_data.group(1) if m_data else "DATA_NAO_ENCONTRADA"

    while i < len(lines_all):
        line = lines_all[i]
        if ("AMERICAS MEDICAL CITY" in line) or ("ALCLIMA" in line):
            i += 1
            continue

        if "Período" in line:
            window = ",".join(lines_all[max(0, i-2):min(len(lines_all), i+5)])
            m = re.search(
                r'[",]*Período\s*:[" ,]*([0-3]\d/[0-1]\d/\d{4}\s*a\s*[0-3]\d/[0-1]\d/\d{4})', window)
            current_periodo = m.group(1).strip() if m else default_periodo

        if line.startswith("Setor:"):
            parts = next(csv.reader([line]))
            if len(parts) >= 2:
                current_setor = parts[1].strip().strip('"')

        if line.startswith("Paciente:"):
            fields = next(csv.reader([line]))
            payload = fields[1].strip().strip('"') if len(fields) >= 2 else ""
            split_marker = "  Entrada: "
            p_ent = payload.find(split_marker)
            id_nome = payload[:p_ent].strip() if p_ent != - \
                1 else payload.strip()
            entrada = extract_between(payload, "  Entrada: ", "  Alta: ")
            alta = extract_between(payload, "  Alta: ", "  Convênio: ")
            convenio = extract_between(payload, "  Convênio: ", "  Plano: ")
            plano = extract_plano(payload)

            j = i + 1
            while j < len(lines_all):
                l2 = lines_all[j]
                if l2.startswith("Paciente:") or l2.startswith("Setor:"):
                    break
                if ("AMERICAS MEDICAL CITY" in l2) or ("ALCLIMA" in l2):
                    j += 1
                    continue
                if l2.startswith('"Tipo de Produto:"'):
                    prod_fields = next(csv.reader([l2]))
                    tipo_produto = prod_fields[1].strip().strip(
                        '"') if len(prod_fields) > 1 else ""

                    k = j + 1
                    while k < len(lines_all):
                        l3 = lines_all[k]
                        if ("AMERICAS MEDICAL CITY" in l3) or ("ALCLIMA" in l3):
                            k += 1
                            continue
                        if "Total do Tipo de Produto:" in l3:
                            tot_fields = next(csv.reader([l3]))
                            qtd_total = parse_number_us(tot_fields[1]) if len(
                                tot_fields) > 1 else None
                            custo_atual = parse_number_us(
                                tot_fields[2]) if len(tot_fields) > 2 else None
                            consumo_total = parse_number_us(
                                tot_fields[3]) if len(tot_fields) > 3 else None
                            periodo_final = current_periodo or default_periodo
                            records.append({
                                "Extração Sishop": data_extracao,
                                "Período": periodo_final,
                                "Setor": current_setor,
                                "Paciente": id_nome,
                                "Entrada": entrada,
                                "Alta": alta,
                                "Convênio": convenio,
                                "Plano": plano,
                                "Tipo de Produto": tipo_produto,
                                "Qtd. Total": qtd_total,
                                "Custo Atual": custo_atual,
                                "Consumo Total": consumo_total
                            })
                            j = k
                            break
                        if l3.startswith('"Tipo de Produto:"') or l3.startswith("Paciente:") or l3.startswith("Setor:"):
                            break
                        k += 1
                j += 1
            i = j - 1
        i += 1

    df = pd.DataFrame(records, columns=[
        "Extração Sishop", "Período", "Setor", "Paciente", "Entrada", "Alta",
        "Convênio", "Plano", "Tipo de Produto", "Qtd. Total", "Custo Atual", "Consumo Total"
    ])
    return df

# ----------------------- Interface Streamlit -----------------------


try:
    st.image("image001 (1).png", use_container_width=False, width=760)
except TypeError:
    st.image("image001 (1).png", use_column_width=False, width=760)

st.markdown("### 1️⃣ Envie o arquivo .txt para processar")
uploaded = st.file_uploader("Selecionar arquivo", type=["txt"])

if uploaded:
    text = uploaded.read().decode("utf-8", errors="ignore")
    df = process_txt_content(text)

    # DE-PARA Setor
    depara_path = os.path.join(os.getcwd(), "DE PARA SETOR.xlsx")
    if os.path.exists(depara_path):
        df_depara = pd.read_excel(depara_path, header=0)
        df_depara.columns = df_depara.columns.str.strip()
        if len(df_depara.columns) >= 2:
            col_setor = df_depara.columns[0]
            col_correlata = df_depara.columns[1]
            df = df.merge(df_depara[[col_setor, col_correlata]],
                          how="left", left_on="Setor", right_on=col_setor)
            df.insert(df.columns.get_loc("Setor") + 1,
                      "Setor Agrupado", df[col_correlata])
            df["Setor Agrupado"].fillna(
                "*SOLICITAR ASSOCIAÇÃO DE SETOR*", inplace=True)
            df.drop(columns=[col_setor, col_correlata], inplace=True)

    # Ordenação + contagem única Paciente x Setor
    if "Setor Agrupado" in df.columns:
        df.sort_values(by=["Paciente", "Setor Agrupado"],
                       inplace=True, ignore_index=True)
        df["Cont. Pac.&Setor Unico"] = (~df.duplicated(
            subset=["Paciente", "Setor Agrupado"], keep="first")).astype(int)
    else:
        df.sort_values(by=["Paciente", "Setor"],
                       inplace=True, ignore_index=True)
        df["Cont. Pac.&Setor Unico"] = (~df.duplicated(
            subset=["Paciente", "Setor"], keep="first")).astype(int)

    # Conversão para numérico nas colunas de valores
    df_export = df.copy()
    for c in ["Qtd. Total", "Custo Atual", "Consumo Total"]:
        df_export[c] = pd.to_numeric(df_export[c], errors="coerce")

    # ------------------- Visualizações -------------------
    st.markdown("### 3️⃣ VOLUME DE ATENDIMENTO POR SETOR AGRUPADO")

    coluna_setor = "Setor Agrupado" if "Setor Agrupado" in df_export.columns else "Setor"
    agrupamento = (
        df_export.groupby(coluna_setor)["Cont. Pac.&Setor Unico"]
        .sum()
        .sort_values(ascending=False)
    )

    # ---------- Gráfico de Barras 3D (efeito) ----------
    fig, ax = plt.subplots(figsize=(6, 2.3))
    bars = ax.bar(agrupamento.index, agrupamento.values,
                  color=plt.cm.viridis(np.linspace(
                      0.3, 0.9, len(agrupamento))),
                  edgecolor="none")

    # Sombreamento 3D suave
    for bar in bars:
        bar.set_zorder(2)
        bar.set_alpha(0.92)
        ax.add_patch(Rectangle(
            (bar.get_x(), 0), bar.get_width(), bar.get_height(),
            facecolor="k", alpha=0.08, zorder=1))

    ax.set_title("VOLUME DE ATENDIMENTO POR SETOR AGRUPADO",
                 fontsize=8.4, fontweight="bold", color="#333")
    ax.set_xlabel("")
    ax.set_ylabel("")
    ax.set_yticks([])
    plt.xticks(rotation=45, ha="right", fontsize=6.6, color="#222")
    ax.set_ylim(0, max(agrupamento) * 1.18)
    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)
    ax.spines["bottom"].set_color("#aaa")

    # Valores no topo (formato BR sem decimais)
    for i, v in enumerate(agrupamento):
        ax.text(i, v + max(agrupamento) * 0.015, f"{int(v):,}".replace(",", "."),
                ha="center", va="bottom", fontsize=7, fontweight="bold", color="#222")

    st.pyplot(fig)

    # ---------- Tabela Resumo com totais + formatação condicional ----------
    st.markdown("#### 📊 RESUMO DE ATENDIMENTO POR SETOR AGRUPADO")

    # Base numérica para cálculos/plots
    df_resumo_base = agrupamento.reset_index()
    df_resumo_base.columns = ["Setor Agrupado", "Volume de Atendimentos"]
    df_resumo_base["% do Total"] = (
        df_resumo_base["Volume de Atendimentos"] /
        df_resumo_base["Volume de Atendimentos"].sum() * 100
    ).round(2)

    total_row = pd.DataFrame({
        "Setor Agrupado": ["TOTAL GERAL"],
        "Volume de Atendimentos": [df_resumo_base["Volume de Atendimentos"].sum()],
        "% do Total": [100.00],
    })
    df_resumo_base = pd.concat([df_resumo_base, total_row], ignore_index=True)

    # DataFrame exibido (formatado BR)
    df_resumo_disp = df_resumo_base.copy()
    df_resumo_disp["Volume de Atendimentos"] = df_resumo_disp["Volume de Atendimentos"].apply(
        lambda x: f"{int(x):,}".replace(",", "."))
    df_resumo_disp["% do Total"] = df_resumo_disp["% do Total"].apply(
        lambda x: f"{x:.2f}%")

    # Formatação condicional (gradiente de verde) apenas nas linhas não-totais
    df_cond = df_resumo_base.copy()
    df_cond["__is_total__"] = df_cond["Setor Agrupado"].eq("TOTAL GERAL")
    max_val = df_cond.loc[~df_cond["__is_total__"],
                          "Volume de Atendimentos"].max()
    min_val = df_cond.loc[~df_cond["__is_total__"],
                          "Volume de Atendimentos"].min()

    def color_scale(val, is_total):
        if is_total or pd.isna(val):
            return ""
        ratio = (val - min_val) / (max_val -
                                   min_val) if max_val != min_val else 0
        r, g, b = int(210 - 120*ratio), int(255 -
                                            110*ratio), int(210 - 200*ratio)
        return f"background-color: rgb({r},{g},{b});"

    def style_row(row):
        is_total = row["Setor Agrupado"] == "TOTAL GERAL"
        styles = []
        for col in df_resumo_disp.columns:
            if col == "Volume de Atendimentos":
                raw_val = df_cond.loc[row.name, "Volume de Atendimentos"]
                styles.append(color_scale(raw_val, is_total))
            else:
                styles.append("")
        if is_total:
            styles = [
                "font-weight: bold; background-color: #e8e8e8;"] * len(styles)
        return styles

    styled = (df_resumo_disp.style
              .set_table_styles([
                  {"selector": "th", "props": [
                      ("padding", "3px 6px"), ("font-size", "12px")]},
                  {"selector": "td", "props": [
                      ("padding", "1px 4px"), ("font-size", "11px")]}
              ])
              .apply(style_row, axis=1)
              )

    st.dataframe(styled, use_container_width=True)

    # ---------- Gráfico de Pizza (interativo) ----------
    st.markdown("#### 🥧 Distribuição Percentual por Setor Agrupado (interativo)")

    df_pie = df_resumo_base[df_resumo_base["Setor Agrupado"]
                            != "TOTAL GERAL"].copy()
    fig_pie = px.pie(
        df_pie,
        names="Setor Agrupado",
        values="Volume de Atendimentos",
        hole=0.35
    )
    fig_pie.update_traces(textposition="inside",
                          texttemplate="%{label}<br>%{percent}")
    fig_pie.update_layout(margin=dict(l=0, r=0, t=20, b=0), height=320)
    st.plotly_chart(fig_pie, use_container_width=True,
                    config={"displaylogo": False})

    # ------------------- Exportação Excel -------------------
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        # Dados completos
        df_export.to_excel(writer, index=False,
                           sheet_name="Protocolo Prisma ver. 0.7i")
        # Resumo (formatado para leitura)
        df_resumo_disp.to_excel(writer, index=False,
                                sheet_name="Resumo por Setor Agrupado")

        # Exporta imagem do gráfico de barras (requer Pillow para embed)
        img_buffer = io.BytesIO()
        fig.savefig(img_buffer, format="png", dpi=200, bbox_inches="tight")
        img_buffer.seek(0)
        from openpyxl.drawing.image import Image
        ws = writer.book.create_sheet("Gráfico Barras 3D")
        ws.add_image(Image(img_buffer), "A1")

    buffer.seek(0)

    # Nome do arquivo de saída
    periodo_valor = (str(df_export.iloc[0]["Período"]).replace("/", "-").replace(" ", "_")
                     if not df_export.empty else "sem_periodo")
    nome_arquivo = f"Prot_Prisma_{periodo_valor}_Sishop.xlsx"

    st.success("✅ Processamento completo! Barras 3D, tabela com totais e formatação condicional, e pizza interativa incluídas.")
    st.download_button(label="📥 Baixar Excel Gerado", data=buffer, file_name=nome_arquivo,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("Envie um arquivo .txt para iniciar o processamento.")
