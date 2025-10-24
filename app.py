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
st.caption("Protocolo para conversão de arquivo .txt em Excel — seleção múltipla, consolidação por mês (sempre o último upload), gráficos e memória histórica (.parquet).")
st.markdown("**Para perfeita execução do Protocolo Prisma, extraia o relatório de Consumo Normal - Sishop, estritamente nas configurações da figura abaixo e salve os .txt.**")
st.markdown("**IMPORTANTE: A VOLUMETRIA APURADA REMETE AOS PACIENTES COM CONSUMO NO MÊS EM QUESTÃO E NÃO NA DATA DE ENTRADA / ATENDIMENTO.**")

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
        r'[",]*Per[ií]odo\s*:[" ,]*([0-3]?\d/[0-1]?\d/\d{4})\s*a\s*([0-3]?\d/[0-1]?\d/\d{4})', header)
    return f"{m.group(1)} a {m.group(2)}" if m else ""


def periodo_label_br(periodo_str: str) -> str:
    """Converte 'dd/mm/yyyy a dd/mm/yyyy' -> 'mmm/aa' em pt-br (ex.: 'jan/25')."""
    if not isinstance(periodo_str, str) or 'a' not in periodo_str:
        return ""
    inicio = periodo_str.split('a')[0].strip()
    try:
        d = pd.to_datetime(inicio, format="%d/%m/%Y", errors="coerce")
        if pd.isna(d):
            return ""
        meses = {1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
                 7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez"}
        return f"{meses[d.month]}/{str(d.year)[-2:]}"
    except Exception:
        return ""


def parse_periodo_to_dates(periodo_str: str):
    """
    Recebe 'dd/mm/yyyy a dd/mm/yyyy' e retorna (dt_inicio, dt_fim) em datetime.date.
    Retorna (None, None) se não parseável.
    """
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
    # usado para ordenar 'jan/25', 'fev/25', ...
    if not isinstance(lbl, str) or "/" not in lbl:
        return (9999, 99)
    mes_map = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
               "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}
    try:
        m, a = lbl.split("/")
        return (2000 + int(a), mes_map.get(m, 99))
    except Exception:
        return (9999, 99)


def fmt_de_ate(dt_ini: date, dt_fim: date) -> str:
    if dt_ini is None or dt_fim is None or pd.isna(dt_ini) or pd.isna(dt_fim):
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
                r'[",]*Per[ií]odo\s*:[" ,]*([0-3]?\d/[0-1]?\d/\d{4})\s*a\s*([0-3]?\d/[0-1]?\d/\d{4})', window)
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
        "Arquivo Origem", "Upload Seq", "Extração Sishop", "Período", "Período Label", "Setor", "Paciente", "Entrada", "Alta",
        "Convênio", "Plano", "Tipo de Produto", "Qtd. Total", "Custo Atual", "Consumo Total"
    ])
    return df


def process_multiple_texts(file_infos) -> pd.DataFrame:
    """
    Recebe lista de dicts {'name','text','upload_seq'} já filtrada (sem meses duplicados).
    """
    frames = []
    for info in file_infos:
        frames.append(process_txt_content(
            info["text"], origem_nome=info["name"], upload_seq=info["upload_seq"]))
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)

# ----------------------- Interface -----------------------


try:
    st.image("image001 (1).png", use_container_width=False, width=1000)
except TypeError:
    st.image("image001 (1).png", use_column_width=False, width=1000)

st.markdown("### 1️⃣ Envie os arquivos .txt para processar (seleção múltipla)")
uploaded_files = st.file_uploader("Selecionar arquivos", type=[
                                  "txt"], accept_multiple_files=True)

# Caminho da memória histórica (persistente em disco) — PARQUET
HIST_PATH = os.path.join(os.getcwd(), "prisma_historico.parquet")

# ------------------- Pré-filtragem no upload: remover meses duplicados -------------------
file_infos_all = []
kept_infos = []
discarded_infos = []

if uploaded_files:
    # 1) Ler conteúdo, detectar períodos e montar lista
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

    # 2) Deduplicar por label: manter o arquivo com MAIOR FIM; em caso de empate, manter o de maior upload_seq (último)
    df_files = pd.DataFrame(file_infos_all)
    if not df_files.empty:
        # ordenar para manter primeiro a melhor versão por Label
        # (Fim desc; UploadSeq desc)
        df_files["_fim_ord"] = pd.to_datetime(df_files["fim"], errors="coerce")
        df_files = df_files.sort_values(
            by=["label", "_fim_ord", "upload_seq"], ascending=[True, False, False])

        # marcar duplicados
        mask_dup = df_files.duplicated(subset=["label"], keep="first")

        kept_df = df_files[~mask_dup].copy()
        disc_df = df_files[mask_dup].copy()

        kept_infos = kept_df.to_dict(orient="records")
        discarded_infos = disc_df.to_dict(orient="records")

        # Exibir lista com período ao lado
        st.markdown(
            "#### 📄 Arquivos **válidos** para processamento (após filtragem por mês):")
        for r in kept_infos:
            st.markdown(
                f"✅ **{r['name']}** — Período: {r['period']} • Label: {r['label']}")

        if len(discarded_infos) > 0:
            st.markdown("---")
            st.markdown("#### 🗑️ Arquivos **descartados** (mês duplicado):")
            for r in discarded_infos:
                st.markdown(
                    f"🗑️ {r['name']} — Período: {r['period']} • Label: {r['label']}")

        # Faixa global (pelos mantidos)
        if len(kept_infos) > 0:
            dt_ini_global = min([x["ini"]
                                for x in kept_infos if x["ini"] is not None])
            dt_fim_global = max([x["fim"]
                                for x in kept_infos if x["fim"] is not None])
            if dt_ini_global and dt_fim_global:
                st.caption(
                    f"📅 Período global considerado nesta carga: **de {dt_ini_global.strftime('%d/%m/%Y')} a {dt_fim_global.strftime('%d/%m/%Y')}**")

    # 3) Processar somente os mantidos
    df = process_multiple_texts(kept_infos)

    # DE-PARA Setor (opcional)
    depara_path = os.path.join(os.getcwd(), "DE PARA SETOR.xlsx")
    if os.path.exists(depara_path) and not df.empty:
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
    if not df.empty:
        if "Setor Agrupado" in df.columns:
            df.sort_values(by=["Paciente", "Setor Agrupado"],
                           inplace=True, ignore_index=True)
            df["Cont. Pac.&Setor Unico"] = (~df.duplicated(
                subset=["Paciente", "Setor Agrupado", "Período Label"], keep="first")).astype(int)
        else:
            df.sort_values(by=["Paciente", "Setor"],
                           inplace=True, ignore_index=True)
            df["Cont. Pac.&Setor Unico"] = (~df.duplicated(
                subset=["Paciente", "Setor", "Período Label"], keep="first")).astype(int)

    # Conversão para numérico (cópia para exportação)
    df_export = df.copy()
    for c in ["Qtd. Total", "Custo Atual", "Consumo Total"]:
        df_export[c] = pd.to_numeric(df_export[c], errors="coerce")

    # ------------------- Escolher o mês mais recente para os CONSOLIDADOS (3️⃣, 📊, 🥧) -------------------
    # cria colunas Inicio/Fim como datas para cada linha
    df_export[["Per_Inicio", "Per_Fim"]] = df_export["Período"].apply(
        lambda s: pd.Series(parse_periodo_to_dates(s))
    )

    # Seleciona o label mais recente (maior Per_Fim) — df_export já está deduplicado por etiqueta de mês
    if not df_export.empty:
        df_label_maxfim = (
            df_export.groupby("Período Label")["Per_Fim"]
            .max()
            .reset_index()
            .dropna(subset=["Per_Fim"])
        )
        if not df_label_maxfim.empty:
            label_mais_recente = df_label_maxfim.loc[df_label_maxfim["Per_Fim"].idxmax(
            ), "Período Label"]
            df_consolidado = df_export[df_export["Período Label"]
                                       == label_mais_recente].copy()
            periodo_mais_recente = df_consolidado.iloc[0]["Período"]
            periodo_mais_recente_label = label_mais_recente
        else:
            df_consolidado = df_export.copy()
            periodo_mais_recente = ""
            periodo_mais_recente_label = ""
    else:
        df_consolidado = df_export.copy()
        periodo_mais_recente = ""
        periodo_mais_recente_label = ""

    # ------------------- Prévia tradicional -------------------
    st.markdown("### 2️⃣ Prévia de conversão do Protocolo Prisma (consolidada)")
    df_preview = df.copy()
    for c in ["Qtd. Total", "Custo Atual", "Consumo Total"]:
        df_preview[c] = df_preview[c].apply(br_format)
    st.dataframe(df_preview.head(15), use_container_width=True)

    # ------------------- Gráfico principal (consolidado — mês mais recente) -------------------
    st.markdown(
        "### 3️⃣ VOLUME DE ATENDIMENTO, COM CONSUMO MENSAL, POR SETOR AGRUPADO (CONSOLIDADO)")
    if periodo_mais_recente:
        # Subtítulo com faixa global do consolidado
        dt_ini_consol = df_consolidado["Per_Inicio"].min()
        dt_fim_consol = df_consolidado["Per_Fim"].max()
        subtitulo_3 = fmt_de_ate(dt_ini_consol, dt_fim_consol)
        if subtitulo_3:
            st.caption(
                f"**Período considerado:** {subtitulo_3} • Label: **{periodo_mais_recente_label}**")

    coluna_setor = "Setor Agrupado" if "Setor Agrupado" in df_consolidado.columns else "Setor"
    agrupamento = (
        df_consolidado.groupby(coluna_setor)["Cont. Pac.&Setor Unico"]
        .sum().sort_values(ascending=False)
    )

    # Paleta única por setor (ajuste visual permitido)
    cmap = plt.cm.get_cmap("tab20", len(agrupamento))
    cores = [cmap(i) for i in range(len(agrupamento))]

    fig, ax = plt.subplots(figsize=(8, 3))
    bars = ax.bar(agrupamento.index, agrupamento.values,
                  color=cores, edgecolor="none")
    for bar, cor in zip(bars, cores):
        bar.set_zorder(2)
        bar.set_alpha(0.92)
        ax.add_patch(Rectangle((bar.get_x(), 0), bar.get_width(),
                     bar.get_height(), facecolor=cor, alpha=0.15, zorder=1))
    ax.set_title(
        "VOLUME DE ATENDIMENTO, COM CONSUMO MENSAL, POR SETOR AGRUPADO (Consolidado)",
        fontsize=6.4, fontweight="bold", color="#333",
    )
    ax.set_xlabel("")
    ax.set_ylabel("")
    ax.set_yticks([])
    plt.xticks(rotation=45, ha="right", fontsize=4.6, color="#222")
    if len(agrupamento) > 0:
        ax.set_ylim(0, max(agrupamento) * 1.18)
        for i, v in enumerate(agrupamento):
            ax.text(i, v + max(agrupamento) * 0.015, f"{int(v):,}".replace(",", "."),
                    ha="center", va="bottom", fontsize=7, fontweight="bold", color="#222")
    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)
    ax.spines["bottom"].set_color("#aaa")
    st.pyplot(fig)
    # Subtítulo logo abaixo do gráfico 3 (faixa global)
    if periodo_mais_recente and subtitulo_3:
        st.caption(f"**Período do gráfico 3️⃣:** {subtitulo_3}")

    # ---------- Tabela Resumo com totais (consolidado — mês mais recente) ----------
    st.markdown(
        "#### 📊 RESUMO DE ATENDIMENTO, COM CONSUMO MENSAL, POR SETOR AGRUPADO (CONSOLIDADO)")
    if periodo_mais_recente and subtitulo_3:
        st.caption(f"Resumo referente ao período {subtitulo_3}")

    df_resumo_base = agrupamento.reset_index()
    df_resumo_base.columns = ["Setor Agrupado", "Volume de Atendimentos"]
    if not df_resumo_base.empty:
        df_resumo_base["% do Total"] = (
            df_resumo_base["Volume de Atendimentos"] /
            df_resumo_base["Volume de Atendimentos"].sum() * 100
        ).round(2)
    total_row = pd.DataFrame({
        "Setor Agrupado": ["TOTAL GERAL"],
        "Volume de Atendimentos": [df_resumo_base["Volume de Atendimentos"].sum() if not df_resumo_base.empty else 0],
        "% do Total": [100.00 if not df_resumo_base.empty else 0.0],
    })
    df_resumo_base = pd.concat([df_resumo_base, total_row], ignore_index=True)

    df_resumo_disp = df_resumo_base.copy()
    if not df_resumo_disp.empty:
        df_resumo_disp["Volume de Atendimentos"] = df_resumo_disp["Volume de Atendimentos"].apply(
            lambda x: f"{int(x):,}".replace(",", "."))
        df_resumo_disp["% do Total"] = df_resumo_disp["% do Total"].apply(
            lambda x: f"{x:.2f}%")
    st.dataframe(df_resumo_disp, use_container_width=True)

    # ---------- Gráfico de Pizza (consolidado — mês mais recente) ----------
    st.markdown(
        "#### 🥧 Distribuição Percentual por Setor Agrupado (Consolidado)")
    if periodo_mais_recente and subtitulo_3:
        st.caption(f"Distribuição referente ao período {subtitulo_3}")

    df_pie = df_resumo_base[df_resumo_base["Setor Agrupado"]
                            != "TOTAL GERAL"].copy()
    if not df_pie.empty:
        fig_pie = px.pie(df_pie, names="Setor Agrupado",
                         values="Volume de Atendimentos", hole=0.35)
        fig_pie.update_traces(textposition="inside",
                              texttemplate="%{label}<br>%{percent}")
        fig_pie.update_layout(margin=dict(l=0, r=0, t=20, b=0), height=320)
        st.plotly_chart(fig_pie, use_container_width=True,
                        config={"displaylogo": False})
    else:
        st.info("Sem dados para o gráfico de pizza.")

    # ------------------- Memória histórica para 4️⃣ (PARQUET) -------------------
    # Cria tabela consolidada do upload atual por Setor Exibição x Período Label
    col_set_exib = "Setor Agrupado" if "Setor Agrupado" in df_export.columns else "Setor"

    # Para cada label do lote atual, usamos o próprio mês consolidado (sem duplicidade no df_export)
    per_last = (
        df_export.groupby("Período Label")
        .apply(lambda d: d.loc[d["Per_Fim"].idxmax(), ["Período", "Per_Inicio", "Per_Fim"]])
        .reset_index()
        .rename(columns={"Período Label": "Per_Label"})
    )
    per_info = {row["Per_Label"]: (
        row["Período"], row["Per_Inicio"], row["Per_Fim"]) for _, row in per_last.iterrows()}

    g_atual = (
        df_export.groupby([col_set_exib, "Período Label"])[
            "Cont. Pac.&Setor Unico"]
        .sum().reset_index()
        .rename(columns={col_set_exib: "Setor Exibicao", "Período Label": "Per_Label", "Cont. Pac.&Setor Unico": "Volume"})
    )
    # Acopla Inicio/Fim/Periodo resolvidos
    g_atual["Periodo"] = g_atual["Per_Label"].map(
        lambda k: per_info.get(k, ("", None, None))[0])
    g_atual["Inicio"] = g_atual["Per_Label"].map(
        lambda k: per_info.get(k, ("", None, None))[1])
    g_atual["Fim"] = g_atual["Per_Label"].map(
        lambda k: per_info.get(k, ("", None, None))[2])

    # Carrega histórico se existir
    if os.path.exists(HIST_PATH):
        try:
            hist = pd.read_parquet(HIST_PATH)
            # Garante tipos de data (date)
            for c in ["Inicio", "Fim"]:
                if c in hist.columns:
                    hist[c] = pd.to_datetime(hist[c], errors="coerce").dt.date
        except Exception:
            hist = pd.DataFrame(
                columns=["Setor Exibicao", "Per_Label", "Periodo", "Inicio", "Fim", "Volume"])
    else:
        hist = pd.DataFrame(
            columns=["Setor Exibicao", "Per_Label", "Periodo", "Inicio", "Fim", "Volume"])

    # Atualiza histórico garantindo unicidade de mês:
    # - Remove do histórico todos os meses presentes nesta carga
    # - Concatena os dados atuais (já sem duplicidade por mês)
    if not g_atual.empty:
        if not hist.empty:
            hist = hist[~hist["Per_Label"].isin(
                g_atual["Per_Label"].unique())].copy()
        hist = pd.concat([hist, g_atual[["Setor Exibicao", "Per_Label",
                         "Periodo", "Inicio", "Fim", "Volume"]]], ignore_index=True)

        # Persistência (normalizando datas para datetime64)
        hist_to_save = hist.copy()
        for c in ["Inicio", "Fim"]:
            if c in hist_to_save.columns:
                hist_to_save[c] = pd.to_datetime(
                    hist_to_save[c], errors="coerce")
        hist_to_save.to_parquet(HIST_PATH, index=False)

    # ------------------- Gráficos por Setor (segregados por Período) — usando memória histórica -------------------
    st.markdown(
        "### 4️⃣ Gráficos de barras por **Setor** agrupados por **Período** (independentes)")

    if not hist.empty:
        df_hist_show = hist.copy()
        df_hist_show["ord_key"] = df_hist_show["Per_Label"].map(periodo_key)
        setores = sorted(df_hist_show["Setor Exibicao"].unique().tolist())

        # Subtítulo geral com a faixa global presente na memória histórica
        dt_ini_hist = pd.to_datetime(
            df_hist_show["Inicio"], errors="coerce").min()
        dt_fim_hist = pd.to_datetime(
            df_hist_show["Fim"], errors="coerce").max()
        subtitulo_hist_global = fmt_de_ate(dt_ini_hist.date() if pd.notna(dt_ini_hist) else None,
                                           dt_fim_hist.date() if pd.notna(dt_fim_hist) else None)
        if subtitulo_hist_global:
            st.caption(
                f"Período global presente na memória histórica: **{subtitulo_hist_global}**")

        for idx, setor in enumerate(setores):
            data_setor = df_hist_show[df_hist_show["Setor Exibicao"] == setor].copy(
            )
            data_setor.sort_values(
                by="Per_Label", key=lambda s: s.map(periodo_key), inplace=True)
            x = data_setor["Per_Label"].tolist()
            y = data_setor["Volume"].tolist()

            fig_s, ax_s = plt.subplots(figsize=(9, 4.5))
            cor = plt.cm.tab20(idx % 20)
            ax_s.bar(x, y, color=cor)
            ax_s.set_title(f"{setor} — Volume por período",
                           fontsize=7, fontweight="bold", color="#333")
            ax_s.set_xlabel("")
            ax_s.set_ylabel("")
            ax_s.set_yticks([])
            plt.xticks(rotation=0, fontsize=7, color="#222")
            if len(y) > 0:
                ax_s.set_ylim(0, max(y) * 1.18)
                for i, v in enumerate(y):
                    ax_s.text(i, v + (max(y) * 0.03 if max(y) > 0 else 0.05),
                              f"{int(v):,}".replace(",", "."), ha="center", va="bottom",
                              fontsize=7, fontweight="bold", color="#222")
            for spine in ["top", "right", "left"]:
                ax_s.spines[spine].set_visible(False)
            ax_s.spines["bottom"].set_color("#aaa")

            st.pyplot(fig_s)

            # Subtítulo por gráfico: faixa global do setor (mais antiga -> mais nova)
            if not data_setor.empty:
                dt_i = pd.to_datetime(
                    data_setor["Inicio"], errors="coerce").min()
                dt_f = pd.to_datetime(data_setor["Fim"], errors="coerce").max()
                sub_setor = fmt_de_ate(dt_i.date() if pd.notna(dt_i) else None,
                                       dt_f.date() if pd.notna(dt_f) else None)
                if sub_setor:
                    st.caption(
                        f"**Período considerado — {setor}:** {sub_setor}")
    else:
        st.info(
            "Memória histórica vazia. Esta será a primeira carga após exportar/processar arquivos.")

    # ------------------- Exportação Excel Consolidada (+ split se exceder) -------------------
    st.markdown("### 5️⃣ Exportação consolidada para Excel")
    if df_consolidado.empty:
        st.warning("Nada a exportar (após o filtro do período consolidado).")
    else:
        df_resumo_export = df_resumo_base.copy()

        def build_excel_bytes(df_main: pd.DataFrame, df_resumo: pd.DataFrame, fig_barras) -> bytes:
            buffer_local = io.BytesIO()
            with pd.ExcelWriter(buffer_local, engine="openpyxl") as writer:
                df_main.to_excel(writer, index=False,
                                 sheet_name="Protocolo Prisma ver. 0.7.6")
                df_resumo.to_excel(writer, index=False,
                                   sheet_name="Resumo por Setor Agrupado")
                img_buffer = io.BytesIO()
                fig_barras.savefig(img_buffer, format="png",
                                   dpi=200, bbox_inches="tight")
                img_buffer.seek(0)
                from openpyxl.drawing.image import Image as XLImage
                ws = writer.book.create_sheet("Gráfico Barras Consolidado")
                ws.add_image(XLImage(img_buffer), "A1")
            buffer_local.seek(0)
            return buffer_local.getvalue()

        EXCEL_MAX = 1_048_576
        linhas = len(df_consolidado)
        periodo_valor = (str(df_consolidado.iloc[0]["Período"]).replace("/", "-").replace(" ", "_")
                         if not df_consolidado.empty else "sem_periodo")
        base_name = f"Prot_Prisma_{periodo_valor}_Sishop"

        if linhas <= EXCEL_MAX:
            bytes_one = build_excel_bytes(
                df_consolidado, df_resumo_export, fig)
            st.download_button(
                label="📥 Baixar Excel Gerado (Consolidado)",
                data=bytes_one,
                file_name=f"{base_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            meio = ceil(linhas / 2)
            parte1 = df_consolidado.iloc[:meio].copy()
            parte2 = df_consolidado.iloc[meio:].copy()
            bytes_p1 = build_excel_bytes(parte1, df_resumo_export, fig)
            bytes_p2 = build_excel_bytes(parte2, df_resumo_export, fig)
            total_partes = 2
            st.info(
                f"O conjunto consolidado excede {EXCEL_MAX:,} linhas — arquivos divididos em 2 partes.")
            col_a, col_b = st.columns(2)
            with col_a:
                st.download_button(
                    label="📥 Baixar Parte 1 (_p1de2)",
                    data=bytes_p1,
                    file_name=f"{base_name}_p1de{total_partes}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with col_b:
                st.download_button(
                    label="📥 Baixar Parte 2 (_p2de2)",
                    data=bytes_p2,
                    file_name=f"{base_name}_p2de{total_partes}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    # ------------------- Botão para limpar memória histórica -------------------
    st.divider()
    if st.button("🗑️ Limpar memória histórica (gráficos 4️⃣)"):
        try:
            if os.path.exists(HIST_PATH):
                os.remove(HIST_PATH)
                st.success("Memória histórica apagada com sucesso.")
            else:
                st.info("Não há memória histórica para limpar.")
        except Exception as e:
            st.error(f"Falha ao limpar memória histórica: {e}")

    st.success(
        "✅ Processamento completo! Conversão consolidada, gráficos e exportação prontos.")
else:
    st.info("Envie um ou mais arquivos .txt para iniciar o processamento.")
