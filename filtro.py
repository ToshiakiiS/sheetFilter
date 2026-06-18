import streamlit as st
import pandas as pd
import io
import zipfile
import csv
import xml.etree.ElementTree as ET

EXCEL_ROW_LIMIT = 1_048_576

def parse_excel_smart_engine(file_bytes, selected_columns=None, headers_only=False):
    """
    Motor híbrido: Tenta a leitura nativa do Pandas primeiro para manter 100% de 
    compatibilidade com arquivos normais. Se falhar por corrupção, ativa a mineração XML.
    """
    in_buffer = io.BytesIO(file_bytes)
    
    # Se for CSV puro, processa direto via Pandas
    if not zipfile.is_zipfile(in_buffer):
        in_buffer.seek(0)
        if headers_only:
            df = pd.read_csv(in_buffer, nrows=0, sep=None, engine='python')
            return df.columns.tolist(), {}
        return pd.read_csv(in_buffer, usecols=selected_columns, sep=None, engine='python')

    # --- ESTRATÉGIA 1: Leitura Nativa Tradicional (Para 99% das planilhas normais do time) ---
    try:
        in_buffer.seek(0)
        if headers_only:
            df = pd.read_excel(in_buffer, nrows=1, engine='openpyxl').dropna(how='all')
            return df.columns.tolist(), {}
        else:
            return pd.read_excel(in_buffer, usecols=selected_columns, engine='openpyxl')
    except Exception:
        # Se der erro de XML estrutural ou travamento de ponteiro, avança para a Estratégia 2
        pass

    # --- ESTRATÉGIA 2: Mineração XML de Emergência (Para o arquivo corrompido de 130MB) ---
    shared_strings = []
    
    # Remove lixo eletrônico injetado por ERPs se necessário
    start_idx = file_bytes.find(b'PK\x03\x04')
    if start_idx > 0:
        file_bytes = file_bytes[start_idx:]

    with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as z:
        namelist = z.namelist()
        norm_map = {n.replace('\\', '/').lower(): name for name in namelist}
        
        # Extrai dicionário global de textos (Shared Strings)
        ss_key = next((orig for norm, orig in norm_map.items() if 'sharedstrings.xml' in norm), None)
        if ss_key:
            try:
                ss_content = z.read(ss_key)
                for event, elem in ET.iterparse(io.BytesIO(ss_content), events=('end',)):
                    if elem.tag.split('}')[-1] == 't':
                        shared_strings.append(elem.text or "")
                        elem.clear()
            except Exception:
                pass
        
        # OTIMIZAÇÃO: Filtra todas as abas internas e seleciona AUTOMATICAMENTE a maior em bytes
        sheet_entries = [item for item in z.infolist() if 'worksheets/sheet' in item.filename.replace('\\', '/').lower()]
        if not sheet_entries:
            raise ValueError("Nenhuma aba de dados localizada dentro do arquivo Excel.")
        
        largest_sheet = max(sheet_entries, key=lambda x: x.file_size)
        sheet_bytes = z.read(largest_sheet.filename)

    # Parse estruturado do XML da aba gigante
    context = ET.iterparse(io.BytesIO(sheet_bytes), events=('start', 'end'))
    context = iter(context)
    event, root = next(context)
    
    rows = []
    current_row = {}
    row_has_cells = False
    
    for event, elem in context:
        tag = elem.tag.split('}')[-1]
        if event == 'start' and tag == 'row':
            current_row = {}
            row_has_cells = False
        elif event == 'end' and tag == 'c':
            r_attr = elem.get('r', '')
            col_letter = ''.join([c for c in r_attr if c.isalpha()])
            t_attr = elem.get('t', '')
            
            val_elem = elem.find('{*}v') or elem.find('v')
            val = val_elem.text if val_elem is not None else ""
            
            if t_attr == 's' and val.isdigit():
                idx = int(val)
                if idx < len(shared_strings):
                    val = shared_strings[idx]
            elif t_attr == 'inlineStr':
                t_elem = elem.find('{*}t') or elem.find('t')
                if t_elem is not None:
                    val = t_elem.text or ""
                    
            if col_letter:
                current_row[col_letter] = val
                row_has_cells = True
        elif event == 'end' and tag == 'row':
            if not row_has_cells:
                root.clear()
                continue
            if current_row:
                rows.append(current_row)
                if headers_only and len(rows) >= 1:
                    # Confere se a linha realmente tem conteúdo textual para evitar cabeçalhos fantasmas
                    if any(str(v).strip() for v in current_row.values()):
                        break
            root.clear()

    if not rows:
        return ([], {}) if headers_only else pd.DataFrame()

    df = pd.DataFrame(rows)
    sorted_cols = sorted(df.columns, key=lambda x: (len(x), x))
    df = df[sorted_cols]
    
    raw_headers = df.iloc[0].fillna('').astype(str).tolist()
    headers = [h.strip() for h in raw_headers if h.strip()]
    
    if headers_only:
        header_to_letter = {str(df.iloc[0][col]).strip(): col for col in df.columns}
        return headers, header_to_letter
        
    df = df[1:]
    df.columns = raw_headers
    
    if selected_columns:
        existing_cols = [c for c in selected_columns if c in df.columns]
        df = df[existing_cols]
        
    return df

# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="Universal Column Selector", page_icon="📊", layout="wide")

st.title("📊 Universal Spreadsheet Column Filter")
st.write(
    "Motor inteligente adaptativo. Esse sistema gerencia de forma autônoma a leitura de arquivos normais "
    "e reconstrói em tempo de execução planilhas corrompidas por sistemas legados."
)

uploaded_file = st.file_uploader("Upload your spreadsheet (.csv, .xlsx)", type=["csv", "xlsx"])

if uploaded_file is not None:
    current_file_key = f"{uploaded_file.name}_{uploaded_file.size}"
    
    if "file_key" not in st.session_state or st.session_state.file_key != current_file_key:
        st.session_state.file_key = current_file_key
        st.session_state.pop("output_key", None)
        st.session_state.pop("output_bytes", None)
        
        with st.spinner("⚡ Analisando e mapeando colunas do arquivo..."):
            uploaded_file.seek(0)
            raw_bytes = uploaded_file.read()
            st.session_state.raw_bytes = raw_bytes
            
            try:
                columns, header_to_letter = parse_excel_smart_engine(raw_bytes, headers_only=True)
                st.session_state.columns = columns
            except Exception as e:
                st.error(f"❌ Não foi possível mapear este arquivo: {e}")
                st.stop()

    columns = st.session_state.columns
    raw_bytes = st.session_state.raw_bytes
    
    st.write("### 🗂️ Select columns to include in the output:")
    selected_columns = st.multiselect(
        "Click below to add/remove columns:",
        options=columns,
        default=columns
    )
    
    if not selected_columns:
        st.warning("⚠️ Please select at least one column to generate an output.")
    else:
        st.write("### ⚙️ Export Settings")
        output_format = st.radio("Choose output format:", ["CSV", "Excel (.xlsx)"], index=1)
        output_key = (current_file_key, output_format, tuple(selected_columns))
        
        if st.button("🚀 Process & Generate Download Link"):
            if st.session_state.get("output_key") != output_key:
                with st.spinner(f"Processando registros e gerando arquivo limpo..."):
                    try:
                        buffer = io.BytesIO()
                        final_df = parse_excel_smart_engine(raw_bytes, selected_columns=selected_columns, headers_only=False)
                        row_count, col_count = final_df.shape

                        if output_format == "CSV":
                            final_df.to_csv(buffer, index=False, encoding='utf-8')
                            mime_type = "text/csv"
                            out_filename = f"filtered_{uploaded_file.name.split('.')[0]}.csv"
                        else:
                            if row_count > EXCEL_ROW_LIMIT:
                                st.error("❌ O formato Excel aceita no máximo 1.048.576 linhas. Use a saída CSV.")
                                st.stop()
                            final_df.to_excel(buffer, index=False, engine="openpyxl")
                            mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            out_filename = f"filtered_{uploaded_file.name.split('.')[0]}.xlsx"

                        st.session_state.output_key = output_key
                        st.session_state.output_bytes = buffer.getvalue()
                        st.session_state.output_filename = out_filename
                        st.session_state.output_mime = mime_type
                        st.session_state.output_rows = row_count

                    except Exception as e:
                        st.error(f"❌ Falha durante o processamento do arquivo: {e}")
                        st.stop()

        if st.session_state.get("output_key") == output_key and st.session_state.get("output_bytes"):
            st.success(f"Pronto! Arquivo processado com sucesso contendo {st.session_state.output_rows:,} linhas.")
            st.download_button(
                label=f"📥 Download Filtered {output_format}",
                data=st.session_state.output_bytes,
                file_name=st.session_state.output_filename,
                mime=st.session_state.output_mime
            )