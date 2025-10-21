# ============================================================
# app.py — Protocolo Prisma ver. 0.7
# Execução: streamlit run app.py
# ============================================================

import streamlit as st
import pandas as pd
import csv
import re
import io
import os

st.set_page_config(page_title="Protocolo Prisma ver. 0.7", layout="wide")
st.title("🧾 Protocolo Prisma — ver. 0.7")
st.caption("Processamento ver. 0.3 + DE PARA SETOR aplicado antes da prévia e do download final.")

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

def detect_default_periodo(raw_text: str) -> str:
    m = re.search(r'Período\s*:.*?"\s*([0-3]\d/[0-1]\d/\d{4}\s*a\s*[0-3]\d/[0-1]\d/\d{4})\s*"', raw_text, flags=re.DOTALL)
    if m: return m.group(1).strip()
    m = re.search(r'Período\s*:\s*([0-3]\d/[0-1]\d/\d{4}\s*a\s*[0-3]\d/[0-1]\d/\d{4})', raw_text, flags=re.DOTALL)
    if m: return m.group(1).strip()
    m = re.search(r'([0-3]\d/[0-1]\d/\d{4})\s*(?:\n|\s)*a\s*(?:\n|\s)*([0-3]\d/[0-1]\d/\d{4})', raw_text)
    if m: return f"{m.group(1)} a {m.group(2)}"
    return ""

# ----------------------- Núcleo ver. 0.3 -----------------------
def process_txt_content(txt: str) -> pd.DataFrame:
    raw = txt.replace(",Setor:,", ",")
    lines_all = raw.splitlines()
    default_periodo = detect_default_periodo(raw)
    current_setor = ""
    current_periodo = default_periodo
    records = []
    i = 0

    # Captura da data de extração (Extração Sishop)
    m_data = re.search(r"Data:\s*,?\s*([0-3]?\d/[0-1]?\d/\d{4})", txt)
    data_extracao = m_data.group(1) if m_data else "DATA_NAO_ENCONTRADA"

    while i < len(lines_all):
        line = lines_all[i]
        # Ignorar cabeçalhos/rodapés do relatório
        if ("AMERICAS MEDICAL CITY" in line) or ("ALCLIMA" in line):
            i += 1
            continue

        # Atualizar período se aparecer no contexto próximo
        if "Período" in line:
            window = ",".join(lines_all[max(0, i-2):min(len(lines_all), i+5)])
            m = re.search(r'Período\s*:.*?"([^"]+)"', window)
            if not m:
                m = re.search(r'Período\s*:\s*([0-3]\d/[0-1]\d/\d{4}\s*a\s*[0-3]\d/[0-1]\d/\d{4})', window)
            current_periodo = m.group(1).strip() if m else default_periodo

        # Captura de Setor
        if line.startswith("Setor:"):
            parts = next(csv.reader([line]))
            if len(parts) >= 2:
                current_setor = parts[1].strip().strip('"')

        # Bloco Paciente
        if line.startswith("Paciente:"):
            fields = next(csv.reader([line]))
            payload = fields[1].strip().strip('"') if len(fields) >= 2 else ""
            split_marker = "  Entrada: "
            p_ent = payload.find(split_marker)
            id_nome = payload[:p_ent].strip() if p_ent != -1 else payload.strip()
            entrada = extract_between(payload, "  Entrada: ", "  Alta: ")
            alta = extract_between(payload, "  Alta: ", "  Convênio: ")
            convenio = extract_between(payload, "  Convênio: ", "  Plano: ")
            plano = extract_plano(payload)

            # Varre tipos de produto e consolida no "Total do Tipo de Produto:"
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
                    tipo_produto = prod_fields[1].strip().strip('"') if len(prod_fields) > 1 else ""

                    k = j + 1
                    while k < len(lines_all):
                        l3 = lines_all[k]
                        if ("AMERICAS MEDICAL CITY" in l3) or ("ALCLIMA" in l3):
                            k += 1
                            continue
                        if "Total do Tipo de Produto:" in l3:
                            tot_fields = next(csv.reader([l3]))
                            qtd_total = parse_number_us(tot_fields[1]) if len(tot_fields) > 1 else None
                            custo_atual = parse_number_us(tot_fields[2]) if len(tot_fields) > 2 else None
                            consumo_total = parse_number_us(tot_fields[3]) if len(tot_fields) > 3 else None
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

# 🔹 Imagem acima da seção 1️⃣
try:
    st.image("image001 (1).png", use_container_width=False, width=760)
except TypeError:
    st.image("image001 (1).png", use_column_width=False, width=760)

st.markdown("### 1️⃣ Envie o arquivo .txt para processar")
uploaded = st.file_uploader("Selecionar arquivo", type=["txt"])

if uploaded:
    text = uploaded.read().decode("utf-8", errors="ignore")

    # --- Etapa 1: Processamento base (ver. 0.3) ---
    df = process_txt_content(text)

    # --- Etapa 2: Aplicar DE PARA ANTES da prévia ---
    depara_path = os.path.join(os.getcwd(), "DE PARA SETOR.xlsx")
    if os.path.exists(depara_path):
        try:
            df_depara = pd.read_excel(depara_path, header=0)
            df_depara.columns = df_depara.columns.str.strip()
            if len(df_depara.columns) >= 2:
                col_setor = df_depara.columns[0]
                col_correlata = df_depara.columns[1]
                df = df.merge(df_depara[[col_setor, col_correlata]],
                              how="left", left_on="Setor", right_on=col_setor)
                df.insert(df.columns.get_loc("Setor") + 1, "Setor Corrigido", df[col_correlata])
                df.drop(columns=[col_setor, col_correlata], inplace=True)
                st.success("🧩 Correlação DE PARA SETOR aplicada com sucesso!")
            else:
                st.warning("⚠️ Arquivo 'DE PARA SETOR.xlsx' não possui colunas suficientes (mínimo: A e B).")
        except Exception as e:
            st.warning(f"⚠️ Falha ao aplicar correlação DE PARA SETOR: {e}")
    else:
        st.info("ℹ️ Arquivo 'DE PARA SETOR.xlsx' não encontrado. Prévia exibida sem correlação.")

    # --- Etapa 3: Prévia (já com DE PARA) ---
    df_preview = df.copy()
    for c in ["Qtd. Total", "Custo Atual", "Consumo Total"]:
        df_preview[c] = df_preview[c].apply(br_format)

    st.markdown("### 2️⃣ Prévia do Protocolo Prisma (com DE PARA aplicado)")
    st.dataframe(df_preview.head(10), use_container_width=True)

    # --- Etapa 4: Exportação ---
    df_export = df.copy()
    for c in ["Qtd. Total", "Custo Atual", "Consumo Total"]:
        df_export[c] = pd.to_numeric(df_export[c], errors="coerce")

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Protocolo Prisma ver. 0.7")
    buffer.seek(0)

    periodo_valor = str(df_export.iloc[0]["Período"]).replace("/", "-").replace(" ", "_") if not df_export.empty else "sem_periodo"
    nome_arquivo = f"Prot_Prisma_{periodo_valor}_Sishop.xlsx"

    st.success("✅ Processamento completo! DE PARA SETOR já incluído na prévia e no arquivo final.")
    st.download_button(label="📥 Baixar Excel Gerado", data=buffer, file_name=nome_arquivo,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Envie um arquivo .txt para iniciar o processamento.")
