# ============================================================
# app.py — Protocolo Prisma ver. 0.7.6 — MULTI-ARQUIVOS + GRÁFICOS + HISTÓRICO
# (Ajustes: exclusão automática de meses duplicados no upload + "último do mês" + subtítulos globais + memória histórica única por mês)
# Base REGRA DE OURO (v0.7) preservada e estendida
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
from math import ceil
from datetime import datetime, date

st.set_page_config(page_title="PROTOCOLO PRISMA VER. 0.7.6", layout="wide")
st.title("🧾 PROTOCOLO PRISMA VER. 0.7.6")

st.caption(
    "Protocolo para conversão de arquivo .txt em Excel — seleção múltipla, consolidação por mês."
)

st.markdown(
    "**Para perfeita execução do Protocolo Prisma, extraia o relatório de Consumo Normal - Sishop, estritamente nas configurações da figura abaixo e salve os .txt.**"
)

st.markdown(
    "**IMPORTANTE: A VOLUMETRIA APURADA REMETE AOS PACIENTES COM CONSUMO NO MÊS EM QUESTÃO E NÃO NA DATA DE ENTRADA / ATENDIMENTO.**"
)

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
    lines = raw_text.splitlines()[:5]
    header = " ".join(lines)

    m = re.search(
        r'[",]*Per[ií]odo\s*:[" ,]*([0-3]?\d/[0-1]?\d/\d{4})\s*a\s*([0-3]?\d/[0-1]?\d/\d{4})',
        header
    )

    return f"{m.group(1)} a {m.group(2)}" if m else ""


# ============================================================
# ALTERAÇÃO 1 — Período Label agora retorna DATA
# ============================================================

def periodo_label_br(periodo_str: str):
    """
    Converte 'dd/mm/yyyy a dd/mm/yyyy'
    para data do primeiro dia do mês (01/mm/aa)
    """

    if not isinstance(periodo_str, str) or "a" not in periodo_str:
        return None

    inicio = periodo_str.split("a")[0].strip()

    try:

        d = pd.to_datetime(inicio, format="%d/%m/%Y", errors="coerce")

        if pd.isna(d):
            return None

        return datetime(d.year, d.month, 1)

    except Exception:
        return None


def parse_periodo_to_dates(periodo_str: str):

    if not isinstance(periodo_str, str) or 'a' not in periodo_str:
        return (None, None)

    partes = periodo_str.split('a')

    if len(partes) != 2:
        return (None, None)

    ini = partes[0].strip()
    fim = partes[1].strip()

    try:

        dt_ini = datetime.strptime(ini, "%d/%m/%Y").date()
        dt_fim = datetime.strptime(fim, "%d/%m/%Y").date()

        return (dt_ini, dt_fim)

    except Exception:
        return (None, None)


def periodo_key(lbl):

    if pd.isna(lbl):
        return (9999, 99)

    try:

        d = pd.to_datetime(lbl)

        return (d.year, d.month)

    except Exception:
        return (9999, 99)


def fmt_de_ate(dt_ini: date, dt_fim: date) -> str:

    if dt_ini is None or dt_fim is None:
        return ""

    return f"de {dt_ini.strftime('%d/%m/%Y')} a {dt_fim.strftime('%d/%m/%Y')}"
# ----------------------- Núcleo de processamento -----------------------


def process_txt_content(txt: str, origem_nome: str = "", upload_seq: int = 0) -> pd.DataFrame:
    """
    Processa o conteúdo de um .txt (Sishop) e retorna o DataFrame com linhas por (Paciente x Tipo de Produto).
    """

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

        if "Período" in line or "Período" in ",".join(lines_all[max(0, i-2):i+5]):

            window = ",".join(lines_all[max(0, i-2):min(len(lines_all), i+5)])

            m = re.search(
                r'[",]*Per[ií]odo\s*:[" ,]*([0-3]?\d/[0-1]?\d/\d{4})\s*a\s*([0-3]?\d/[0-1]?\d/\d{4})',
                window
            )

            if m:
                current_periodo = f"{m.group(1)} a {m.group(2)}"
            else:
                current_periodo = default_periodo

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

                                "Arquivo Origem": origem_nome,
                                "Upload Seq": upload_seq,
                                "Extração Sishop": data_extracao,

                                "Período": periodo_final,

                                "Período Label": periodo_label_br(periodo_final),

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

        "Arquivo Origem",
        "Upload Seq",
        "Extração Sishop",

        "Período",
        "Período Label",

        "Setor",

        "Paciente",

        "Entrada",
        "Alta",
        "Convênio",
        "Plano",

        "Tipo de Produto",

        "Qtd. Total",
        "Custo Atual",
        "Consumo Total"

    ])

    # ============================================================
    # ALTERAÇÃO 2 — SPLIT DA COLUNA PACIENTE
    # ============================================================

    if not df.empty and "Paciente" in df.columns:

        split_cols = df["Paciente"].str.split(" - ", n=1, expand=True)

        df["Registro"] = split_cols[0]
        df["Nome do Paciente"] = split_cols[1]

        df.drop(columns=["Paciente"], inplace=True)

        # reorganizar posição das colunas (mantendo lógica do arquivo)
        cols = df.columns.tolist()

        pos = cols.index("Registro")

        cols.insert(pos + 1, cols.pop(cols.index("Nome do Paciente")))

        df = df[cols]

    return df


def process_multiple_texts(file_infos) -> pd.DataFrame:
    """
    Recebe lista de dicts {'name','text','upload_seq'} já filtrada (sem meses duplicados).
    """

    frames = []

    for info in file_infos:

        frames.append(
            process_txt_content(
                info["text"],
                origem_nome=info["name"],
                upload_seq=info["upload_seq"]
            )
        )

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)
# ----------------------- Interface -----------------------


try:
    st.image("image001 (1).png", use_container_width=False, width=1000)
except TypeError:
    st.image("image001 (1).png", use_column_width=False, width=1000)

st.markdown("### 1️⃣ Envie os arquivos .txt para processar (seleção múltipla)")

uploaded_files = st.file_uploader(
    "Selecionar arquivos",
    type=["txt"],
    accept_multiple_files=True
)

# Caminho da memória histórica
HIST_PATH = os.path.join(os.getcwd(), "prisma_historico.parquet")

# ------------------- Pré-filtragem upload -------------------

file_infos_all = []
kept_infos = []
discarded_infos = []

if uploaded_files:

    for idx, f in enumerate(uploaded_files, start=1):

        try:
            content = f.read().decode("utf-8", errors="ignore")
        except Exception:
            f.seek(0)
            content = f.read().decode("latin-1", errors="ignore")

        periodo = detect_periodo_first_lines(content)

        label = periodo_label_br(periodo)

        dt_ini, dt_fim = parse_periodo_to_dates(periodo)

        file_infos_all.append({

            "name": getattr(f, "name", "arquivo.txt"),
            "text": content,

            "period": periodo,
            "label": label,

            "ini": dt_ini,
            "fim": dt_fim,

            "upload_seq": idx

        })

    df_files = pd.DataFrame(file_infos_all)

    if not df_files.empty:

        df_files["_fim_ord"] = pd.to_datetime(df_files["fim"], errors="coerce")

        df_files = df_files.sort_values(
            by=["label", "_fim_ord", "upload_seq"],
            ascending=[True, False, False]
        )

        mask_dup = df_files.duplicated(subset=["label"], keep="first")

        kept_df = df_files[~mask_dup].copy()
        disc_df = df_files[mask_dup].copy()

        kept_infos = kept_df.to_dict(orient="records")
        discarded_infos = disc_df.to_dict(orient="records")

        st.markdown("#### 📄 Arquivos válidos para processamento")

        for r in kept_infos:

            lbl = r["label"]

            lbl_txt = (
                pd.to_datetime(lbl).strftime("%d/%b/%y")
                if pd.notna(lbl) else ""
            )

            st.markdown(
                f"✅ **{r['name']}** — Período: {r['period']} • Label: {lbl_txt}"
            )

        if len(discarded_infos) > 0:

            st.markdown("---")
            st.markdown("#### 🗑️ Arquivos descartados")

            for r in discarded_infos:

                lbl = r["label"]

                lbl_txt = (
                    pd.to_datetime(lbl).strftime("%d/%b/%y")
                    if pd.notna(lbl) else ""
                )

                st.markdown(
                    f"🗑️ {r['name']} — Período: {r['period']} • Label: {lbl_txt}"
                )

        if len(kept_infos) > 0:

            dt_ini_global = min(
                [x["ini"] for x in kept_infos if x["ini"] is not None]
            )

            dt_fim_global = max(
                [x["fim"] for x in kept_infos if x["fim"] is not None]
            )

            if dt_ini_global and dt_fim_global:

                st.caption(
                    f"📅 Período global considerado: **de {dt_ini_global.strftime('%d/%m/%Y')} a {dt_fim_global.strftime('%d/%m/%Y')}**"
                )

    # ------------------- Processamento -------------------

    df = process_multiple_texts(kept_infos)

    # ------------------- DE PARA SETOR -------------------

    depara_path = os.path.join(os.getcwd(), "DE PARA SETOR.xlsx")

    if os.path.exists(depara_path) and not df.empty:

        df_depara = pd.read_excel(depara_path, header=0)

        df_depara.columns = df_depara.columns.str.strip()

        if len(df_depara.columns) >= 2:

            col_setor = df_depara.columns[0]
            col_correlata = df_depara.columns[1]

            df = df.merge(
                df_depara[[col_setor, col_correlata]],
                how="left",
                left_on="Setor",
                right_on=col_setor
            )

            df.insert(
                df.columns.get_loc("Setor") + 1,
                "Setor Agrupado",
                df[col_correlata]
            )

            df["Setor Agrupado"].fillna(
                "*SOLICITAR ASSOCIAÇÃO DE SETOR*",
                inplace=True
            )

            df.drop(columns=[col_setor, col_correlata], inplace=True)

    # ------------------- Ordenação -------------------

    if not df.empty:

        if "Setor Agrupado" in df.columns:

            df.sort_values(
                by=["Registro", "Setor Agrupado"],
                inplace=True,
                ignore_index=True
            )

            df["Cont. Pac.&Setor Unico"] = (

                ~df.duplicated(
                    subset=["Registro", "Setor Agrupado", "Período Label"],
                    keep="first"
                )

            ).astype(int)

        else:

            df.sort_values(
                by=["Registro", "Setor"],
                inplace=True,
                ignore_index=True
            )

            df["Cont. Pac.&Setor Unico"] = (

                ~df.duplicated(
                    subset=["Registro", "Setor", "Período Label"],
                    keep="first"
                )

            ).astype(int)

    # ------------------- Conversões numéricas -------------------

    df_export = df.copy()

    for c in ["Qtd. Total", "Custo Atual", "Consumo Total"]:

        df_export[c] = pd.to_numeric(df_export[c], errors="coerce")

    # ------------------- Prévia -------------------

    st.markdown(
        "### 2️⃣ Prévia de conversão do Protocolo Prisma (consolidada)"
    )

    df_preview = df.copy()

    # formatar números BR

    for c in ["Qtd. Total", "Custo Atual", "Consumo Total"]:

        df_preview[c] = df_preview[c].apply(br_format)

    # formatar DATA visualmente

    if "Período Label" in df_preview.columns:

        df_preview["Período Label"] = pd.to_datetime(
            df_preview["Período Label"],
            errors="coerce"
        ).dt.strftime("%d/%b/%y")

    st.dataframe(
        df_preview.head(15),
        use_container_width=True
    )
    # ------------------- Escolher mês mais recente -------------------

    df_export[["Per_Inicio", "Per_Fim"]] = df_export["Período"].apply(
        lambda s: pd.Series(parse_periodo_to_dates(s))
    )

    if not df_export.empty:

        df_label_maxfim = (
            df_export.groupby("Período Label")["Per_Fim"]
            .max()
            .reset_index()
            .dropna(subset=["Per_Fim"])
        )

        if not df_label_maxfim.empty:

            label_mais_recente = df_label_maxfim.loc[
                df_label_maxfim["Per_Fim"].idxmax(),
                "Período Label"
            ]

            df_consolidado = df_export[
                df_export["Período Label"] == label_mais_recente
            ].copy()

            periodo_mais_recente = df_consolidado.iloc[0]["Período"]

        else:

            df_consolidado = df_export.copy()
            periodo_mais_recente = ""

    else:

        df_consolidado = df_export.copy()
        periodo_mais_recente = ""

    # ------------------- Gráfico principal -------------------

    st.markdown(
        "### 3️⃣ VOLUME DE ATENDIMENTO, COM CONSUMO MENSAL, POR SETOR AGRUPADO (CONSOLIDADO)"
    )

    coluna_setor = (
        "Setor Agrupado"
        if "Setor Agrupado" in df_consolidado.columns
        else "Setor"
    )

    agrupamento = (

        df_consolidado
        .groupby(coluna_setor)["Cont. Pac.&Setor Unico"]
        .sum()
        .sort_values(ascending=False)

    )

    cmap = plt.cm.get_cmap("tab20", len(agrupamento))
    cores = [cmap(i) for i in range(len(agrupamento))]

    fig, ax = plt.subplots(figsize=(8, 3))

    bars = ax.bar(
        agrupamento.index,
        agrupamento.values,
        color=cores
    )

    for bar, cor in zip(bars, cores):

        bar.set_zorder(2)

        ax.add_patch(

            Rectangle(
                (bar.get_x(), 0),
                bar.get_width(),
                bar.get_height(),
                facecolor=cor,
                alpha=0.15,
                zorder=1
            )

        )

    ax.set_title(
        "Volume de Atendimento por Setor Agrupado",
        fontsize=7,
        fontweight="bold"
    )

    ax.set_yticks([])

    plt.xticks(rotation=45, ha="right", fontsize=6)

    for i, v in enumerate(agrupamento):

        ax.text(
            i,
            v,
            f"{int(v):,}".replace(",", "."),
            ha="center",
            va="bottom",
            fontsize=7,
            fontweight="bold"
        )

    st.pyplot(fig)

    # ------------------- Tabela Resumo -------------------

    st.markdown(
        "#### 📊 RESUMO DE ATENDIMENTO POR SETOR AGRUPADO"
    )

    df_resumo = agrupamento.reset_index()

    df_resumo.columns = [
        "Setor Agrupado",
        "Volume de Atendimentos"
    ]

    if not df_resumo.empty:

        df_resumo["% do Total"] = (

            df_resumo["Volume de Atendimentos"]
            / df_resumo["Volume de Atendimentos"].sum()
            * 100

        ).round(2)

    st.dataframe(df_resumo, use_container_width=True)

    # ------------------- Pizza -------------------

    st.markdown(
        "#### 🥧 Distribuição Percentual por Setor Agrupado"
    )

    if not df_resumo.empty:

        fig_pie = px.pie(
            df_resumo,
            names="Setor Agrupado",
            values="Volume de Atendimentos",
            hole=0.35
        )

        st.plotly_chart(
            fig_pie,
            use_container_width=True
        )

    # ------------------- Exportação Excel -------------------

    st.markdown(
        "### 5️⃣ Exportação consolidada para Excel"
    )

    if not df_consolidado.empty:

        buffer = io.BytesIO()

        with pd.ExcelWriter(
            buffer,
            engine="openpyxl"
        ) as writer:

            df_consolidado.to_excel(
                writer,
                index=False,
                sheet_name="Protocolo Prisma"
            )

            df_resumo.to_excel(
                writer,
                index=False,
                sheet_name="Resumo"
            )

            # -------------------
            # FORMATAÇÃO DATA EXCEL
            # -------------------

            ws = writer.book["Protocolo Prisma"]

            col_index = list(df_consolidado.columns).index(
                "Período Label"
            ) + 1

            from openpyxl.styles import numbers

            for row in ws.iter_rows(
                min_row=2,
                min_col=col_index,
                max_col=col_index
            ):

                for cell in row:

                    cell.number_format = "DD/MMM/YY"

        buffer.seek(0)

        periodo_nome = ""

        if not df_consolidado.empty:

            d = pd.to_datetime(
                df_consolidado.iloc[0]["Período Label"]
            )

            periodo_nome = d.strftime("%Y_%m")

        st.download_button(

            label="📥 Baixar Excel Gerado",

            data=buffer,

            file_name=f"Prot_Prisma_{periodo_nome}.xlsx",

            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )

    # ------------------- Limpar memória histórica -------------------

    st.divider()

    if st.button("🗑️ Limpar memória histórica"):

        try:

            if os.path.exists(HIST_PATH):

                os.remove(HIST_PATH)

                st.success(
                    "Memória histórica apagada com sucesso."
                )

            else:

                st.info(
                    "Não há memória histórica para limpar."
                )

        except Exception as e:

            st.error(
                f"Falha ao limpar memória histórica: {e}"
            )

    st.success(
        "✅ Processamento completo!"
    )

else:

    st.info(
        "Envie um ou mais arquivos .txt para iniciar o processamento."
    )
