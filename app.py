from __future__ import annotations

from io import BytesIO
from pathlib import Path
import zipfile

import pandas as pd
import streamlit as st

from core import (
    CaseMapping,
    detect_basic_case_ids,
    parse_combinations_csv,
    parse_member_results_xlsx,
    parse_reactions_csv,
    process_data,
)
from exporters import build_excel, build_word


APP_VERSION = "1.1.0"
BASE_DIR = Path(__file__).resolve().parent
ASSET_EXAMPLE = BASE_DIR / "assets" / "exemplo_resultados_elu_els.png"

st.set_page_config(
    page_title="Gerador de Tabelas - Robot",
    page_icon="📐",
    layout="wide",
)

st.markdown(
    """
    <style>
      .block-container {padding-top: 1.7rem; padding-bottom: 2rem; max-width: 1500px;}
      h1, h2, h3 {font-family: 'Times New Roman', serif;}
      .small-note {color:#666; font-size:0.9rem;}
      div[data-testid="stFileUploader"] {border: 1px solid #dedede; border-radius: 8px; padding: .35rem .55rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Gerador de Tabelas - Robot Structural Analysis")
st.caption(f"Robot → Memória de Cálculo Blossom | v{APP_VERSION}")

with st.expander("Como preparar os 3 arquivos de entrada", expanded=False):
    st.markdown(
        """
        **1. CSV de combinações**  
        Exporte diretamente do Robot em `.csv`. O arquivo deve conter, entre outras, as colunas **Nome**, **Tipo de combinação** e **Definição**.

        **2. CSV de reações**  
        Exporte diretamente da tabela de reações do Robot em `.csv`, contendo **Nó/Caso, FX, FY, FZ, MX, MY e MZ**.

        **3. Excel ELU/ELS**  
        Na janela **Verificação de membro (ELS; ELU)** do Robot, selecione todas as linhas, copie e cole a partir da célula **A1** de um Excel em branco.  
        O cabeçalho não precisa ser copiado. Salve como `.xlsx`.
        """
    )
    if ASSET_EXAMPLE.exists():
        st.image(str(ASSET_EXAMPLE), caption="Exemplo do cabeçalho da janela Verificação de membro (ELS; ELU)", use_container_width=True)

st.subheader("1. Arquivos de entrada")
c1, c2, c3 = st.columns(3)
with c1:
    file_comb = st.file_uploader(
        "CSV de combinações",
        type=["csv"],
        help="Exportação direta do Robot. Normalmente UTF-16 e separado por ponto e vírgula.",
        key="comb",
    )
with c2:
    file_reac = st.file_uploader(
        "CSV de reações",
        type=["csv"],
        help="Tabela de reações exportada do Robot, incluindo Nó/Caso e seis componentes.",
        key="reac",
    )
with c3:
    file_members = st.file_uploader(
        "Excel de resultados ELU/ELS",
        type=["xlsx", "xlsm"],
        help="Excel criado em branco com os resultados da janela Verificação de membro (ELS; ELU) copiados para A1.",
        key="members",
    )

all_files = file_comb is not None and file_reac is not None and file_members is not None

parsed = None
if all_files:
    try:
        combinations = parse_combinations_csv(file_comb.getvalue())
        reactions = parse_reactions_csv(file_reac.getvalue())
        members = parse_member_results_xlsx(file_members.getvalue())
        case_ids = detect_basic_case_ids(combinations, reactions)
        parsed = (combinations, reactions, members, case_ids)
        st.success(
            f"Arquivos lidos: {len(combinations)} combinações, {len(reactions)} linhas de reações e {len(members)} membros."
        )
    except Exception as exc:
        st.error(f"Falha ao interpretar os arquivos: {exc}")

if parsed:
    combinations, reactions, members, case_ids = parsed

    st.subheader("2. Identificação e mapeamento dos casos")
    left, right = st.columns([1, 2])
    with left:
        structure_name = st.text_input("Nome da estrutura", value="Plataforma")
        st.markdown(
            '<div class="small-note">O nome é usado nos títulos das tabelas do Excel e do Word.</div>',
            unsafe_allow_html=True,
        )
    with right:
        st.info(
            "Os números dos casos não são fixos. Confirme o nome e a abreviação de cada caso detectado. "
            "É esse mapeamento que será usado para traduzir as combinações e nomear as reações."
        )

    default_map = pd.DataFrame(
        [
            {
                "Caso Robot": cid,
                "Carregamento": f"CASO {cid}",
                "Abreviação": f"C{cid}",
                "Incluir nas reações": True,
            }
            for cid in case_ids
        ]
    )
    edited_map = st.data_editor(
        default_map,
        use_container_width=True,
        hide_index=True,
        disabled=["Caso Robot"],
        column_config={
            "Caso Robot": st.column_config.NumberColumn("Caso Robot", format="%d"),
            "Carregamento": st.column_config.TextColumn("Carregamento", help="Ex.: PESO PRÓPRIO, TEMPERATURA +20°C"),
            "Abreviação": st.column_config.TextColumn("Abreviação", help="Ex.: PP, CP, SC, T+°C"),
            "Incluir nas reações": st.column_config.CheckboxColumn("Incluir nas reações"),
        },
        key="case_mapping_editor",
    )

    mappings = [
        CaseMapping(
            case_id=int(row["Caso Robot"]),
            load_name=str(row["Carregamento"] or "").strip(),
            abbreviation=str(row["Abreviação"] or "").strip(),
            include_reactions=bool(row["Incluir nas reações"]),
        )
        for _, row in edited_map.iterrows()
    ]

    with st.expander("Configurações avançadas", expanded=False):
        a1, a2, a3 = st.columns(3)
        with a1:
            lambda_limit = st.number_input("λ máximo padrão", min_value=1.0, value=200.0, step=10.0)
        with a2:
            force_factor = st.number_input(
                "Fator kgf → kN",
                min_value=0.000001,
                value=0.01,
                step=0.001,
                format="%.5f",
                help="Padrão adotado: 0,01, correspondente à aproximação g ≈ 10 m/s².",
            )
        with a3:
            include_raw = st.checkbox("Incluir dados brutos no Excel", value=True, help="As abas RAW são mantidas ocultas para rastreabilidade.")
        moment_factor = st.number_input(
            "Fator kgfm → kN·m",
            min_value=0.000001,
            value=float(force_factor),
            step=0.001,
            format="%.5f",
        )

    confirmed = st.checkbox(
        "Confirmo que revisei o mapeamento dos números dos casos de carregamento.",
        value=False,
    )

    if st.button("Processar arquivos", type="primary", use_container_width=True, disabled=not confirmed):
        processed = process_data(
            combinations=combinations,
            reactions=reactions,
            members=members,
            case_mappings=mappings,
            structure_name=structure_name.strip() or "Estrutura",
            lambda_limit=float(lambda_limit),
            force_conversion_factor=float(force_factor),
            moment_conversion_factor=float(moment_factor),
        )

        input_names = {
            "combinations": file_comb.name,
            "reactions": file_reac.name,
            "members": file_members.name,
        }
        excel_bytes = build_excel(processed, mappings, input_names, include_raw_data=include_raw)
        word_bytes = build_word(processed, mappings, input_names)

        zip_io = BytesIO()
        safe = re_safe = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in (structure_name.strip() or "Estrutura"))
        with zipfile.ZipFile(zip_io, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(f"Tabelas_Robot_{safe}.xlsx", excel_bytes)
            zf.writestr(f"Tabelas_Robot_{safe}.docx", word_bytes)
        zip_bytes = zip_io.getvalue()

        st.session_state["processed"] = processed
        st.session_state["downloads"] = (excel_bytes, word_bytes, zip_bytes, safe)
        st.session_state["mappings"] = mappings

if "processed" in st.session_state:
    processed = st.session_state["processed"]
    st.subheader("3. Validação")
    for item in processed.validations:
        if item.level == "success":
            st.success(item.message)
        elif item.level == "error":
            st.error(item.message)
        elif item.level == "warning":
            st.warning(item.message)
        else:
            st.info(item.message)

    st.subheader("4. Pré-visualização")
    tabs = st.tabs(["ELU", "ELS", "Tabela 8", "Tabela 9", "Reações"])
    with tabs[0]:
        st.dataframe(pd.DataFrame(processed.combinations_elu, columns=["Combinação", "Expressão para MC", "Definição Robot"]), use_container_width=True, hide_index=True)
    with tabs[1]:
        st.dataframe(pd.DataFrame(processed.combinations_els, columns=["Combinação", "Expressão para MC", "Definição Robot"]), use_container_width=True, hide_index=True)
    with tabs[2]:
        st.dataframe(pd.DataFrame(processed.table8, columns=["Membro", "Perfil", "Material", "Lay", "Laz", "Índice ELU", "Caso", "Status Tensão", "Status Esbeltez"]), use_container_width=True, hide_index=True)
    with tabs[3]:
        st.dataframe(pd.DataFrame(processed.table9, columns=["Membro", "Perfil", "Ratio (uy)", "Caso (uy)", "Ratio (uz)", "Caso (uz)", "Ratio (vx)", "Caso (vx)", "Ratio (vy)", "Caso (vy)", "Status Flecha"]), use_container_width=True, hide_index=True)
    with tabs[4]:
        st.dataframe(pd.DataFrame(processed.table10, columns=["Base/nó", "Grupo", "Carregamento", "Fx (kN)", "Fy (kN)", "Fz (kN)", "Mx (kN·m)", "My (kN·m)", "Mz (kN·m)"]), use_container_width=True, hide_index=True)

    st.subheader("5. Exportação")
    excel_bytes, word_bytes, zip_bytes, safe = st.session_state["downloads"]
    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button(
            "Baixar Excel",
            data=excel_bytes,
            file_name=f"Tabelas_Robot_{safe}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with d2:
        st.download_button(
            "Baixar Word",
            data=word_bytes,
            file_name=f"Tabelas_Robot_{safe}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with d3:
        st.download_button(
            "Baixar pacote ZIP",
            data=zip_bytes,
            file_name=f"Tabelas_Robot_{safe}.zip",
            mime="application/zip",
            use_container_width=True,
        )

st.divider()
st.caption("Os dados são processados em memória durante a sessão do Streamlit; o aplicativo não precisa armazenar os arquivos enviados.")
