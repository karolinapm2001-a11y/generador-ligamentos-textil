
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import zipfile
from matplotlib.patches import Circle, Polygon
from io import BytesIO
from xlsxwriter import Workbook

st.set_page_config(
    page_title="Generador Automático de Ligamentos y Levas",
    page_icon="🧶",
    layout="wide"
)

# =========================================================
# ESTILO
# =========================================================
st.markdown("""
<style>
html, body, [class*="css"] {
    background-color:#ffffff !important;
    color:#111111 !important;
}
[data-testid="stAppViewContainer"],
[data-testid="stHeader"],
[data-testid="stSidebar"] {
    background:#ffffff !important;
    color:#111111 !important;
}
.block-container {
    padding-top:1rem;
    padding-bottom:2rem;
}
h1,h2,h3,h4,h5,h6,p,label,span,div {
    color:#111111 !important;
}
.section-title {
    font-weight:700;
    color:#17365D !important;
    font-size:1.08rem;
    margin:.4rem 0 .6rem 0;
    padding-bottom:.28rem;
    border-bottom:2px solid #D9E2F3;
}
input, textarea {
    background:#ffffff !important;
    color:#111111 !important;
}
[data-baseweb="select"] > div,
[data-baseweb="input"] > div {
    background:#ffffff !important;
    color:#111111 !important;
}
[data-baseweb="select"] *,
[data-testid="stDataFrame"] * {
    color:#111111 !important;
}
[data-baseweb="popover"],
[data-baseweb="popover"] > div,
[data-baseweb="menu"],
[data-baseweb="menu"] ul,
[data-baseweb="menu"] li,
[role="listbox"],
[role="option"] {
    background:#ffffff !important;
    color:#111111 !important;
}
[role="option"] *,
[data-baseweb="menu"] li * {
    color:#111111 !important;
}
[role="option"]:hover,
[data-baseweb="menu"] li:hover {
    background:#f2f2f2 !important;
    color:#111111 !important;
}
.stDownloadButton > button {
    background:#2E7D32 !important;
    color:white !important;
    border:none !important;
    border-radius:6px !important;
    font-weight:600 !important;
}
.note {
    background:#F7F9FC;
    border:1px solid #D9E2F3;
    padding:10px 12px;
    border-radius:6px;
}
</style>
""", unsafe_allow_html=True)

st.title("GENERADOR AUTOMÁTICO DE LIGAMENTOS Y LEVAS – TEJIDO CIRCULAR")
st.caption("Versión 5.9 · Plato/Dial descendente · Cilindro ascendente")

# =========================================================
# LIGAMENTO
# =========================================================
def parse_sequence(txt):
    vals = []
    clean = str(txt).replace(",", "-").replace(" ", "-")
    for p in clean.split("-"):
        p = p.strip()
        if not p:
            continue
        try:
            v = int(p)
            if 1 <= v <= 6:
                vals.append(v)
        except:
            pass
    return vals if vals else [1]

def draw_ligament_symbol(ax, sym, x0, y0, width=1.0, height=0.55, lw=2.0):
    # Cada símbolo ocupa exactamente un módulo y enlaza con el siguiente
    x1 = x0 + width
    xm = x0 + width/2
    r = min(width*0.20, height*0.34)

    if sym == 1:
        # círculo debajo de línea, tocando la línea
        ax.plot([x0, x1], [y0, y0], color="black", lw=lw)
        ax.add_patch(Circle((xm, y0-r), r, fill=False, color="black", lw=lw))

    elif sym == 2:
        # círculo encima de línea, tocando la línea
        ax.plot([x0, x1], [y0, y0], color="black", lw=lw)
        ax.add_patch(Circle((xm, y0+r), r, fill=False, color="black", lw=lw))

    elif sym == 3:
        # V hacia abajo: extremos exactamente sobre la línea base
        ax.plot(
            [x0, xm, x1],
            [y0, y0-height*0.58, y0],
            color="black",
            lw=lw
        )

    elif sym == 4:
        # ∧ hacia arriba: extremos exactamente sobre la línea base
        ax.plot(
            [x0, xm, x1],
            [y0, y0+height*0.58, y0],
            color="black",
            lw=lw
        )

    elif sym == 5:
        # línea horizontal con punto debajo
        ax.plot([x0, x1], [y0, y0], color="black", lw=lw)
        ax.plot(
            xm,
            y0-height*0.34,
            marker="o",
            markersize=4.2,
            color="black"
        )

    elif sym == 6:
        # V invertida/cruce hacia abajo con círculo debajo de la línea
        # línea horizontal
        ax.plot([x0, x1], [y0, y0], color="black", lw=lw)

        # dos diagonales hacia el centro superior/inferior según referencia
        ax.plot(
            [x0+width*0.18, xm, x0+width*0.82],
            [y0+height*0.46, y0-height*0.18, y0+height*0.46],
            color="black",
            lw=lw
        )

        # círculo debajo de la línea, tocando la línea
        ax.add_patch(
            Circle(
                (xm, y0-r),
                r,
                fill=False,
                color="black",
                lw=lw
            )
        )

def draw_catalog():
    fig, ax = plt.subplots(figsize=(5.2, 4.5), facecolor="white")
    ax.set_facecolor("white")

    for i, sym in enumerate([1,2,3,4,5,6]):
        y = 5-i
        ax.text(
            0.15, y, str(sym),
            fontsize=13, fontweight="bold",
            va="center", color="black"
        )
        draw_ligament_symbol(
            ax, sym,
            x0=0.95,
            y0=y,
            width=1.05,
            height=0.75,
            lw=2.2
        )

    ax.set_xlim(0,2.35)
    ax.set_ylim(-0.6,5.8)
    ax.axis("off")
    fig.tight_layout()
    return fig

def draw_system_ligaments(system_sequences, visible_slots=12, labels=None, yarns=None,
                          tipo_fontura="Monofontura"):
    n = len(system_sequences)
    labels = labels or [""] * n
    yarns = yarns or [""] * n

    row_gap = 1.15
    fig_h = max(4.4, n*0.90 + 1.4)
    fig_w = max(14.8, visible_slots*0.75 + 8.2)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h), facecolor="white")
    ax.set_facecolor("white")

    lig_width = visible_slots * 1.0
    x_text1 = lig_width + 1.05
    x_text2 = lig_width + 3.10

    for i, seq in enumerate(system_sequences):
        y = (n - 1 - i) * row_gap

        # Repetición continua del ligamento
        for pos in range(visible_slots):
            sym = seq[pos % len(seq)]
            draw_ligament_symbol(
                ax,
                sym,
                x0=pos*1.0,
                y0=y,
                width=1.0,
                height=0.62,
                lw=1.85
            )


        # Texto a la derecha
        ax.text(
            x_text1,
            y,
            str(labels[i]) if i < len(labels) else "",
            ha="center",
            va="center",
            fontsize=9.5,
            fontweight="bold",
            color="#111111"
        )

        ax.text(
            x_text2,
            y,
            str(yarns[i]) if i < len(yarns) else "",
            ha="center",
            va="center",
            fontsize=9.2,
            color="#111111"
        )

        # Separador entre sistemas
        ax.plot(
            [-0.10, lig_width+4.25],
            [y - 0.58, y - 0.58],
            color="#E6E6E6",
            lw=0.7
        )

    top_y = (n - 1) * row_gap

    ax.text(
        x_text1,
        top_y + 0.72,
        "TIPO / DESCRIPCIÓN",
        ha="center",
        va="bottom",
        fontsize=9,
        fontweight="bold",
        color="#17365D"
    )

    ax.text(
        x_text2,
        top_y + 0.72,
        "HILO / MATERIAL",
        ha="center",
        va="bottom",
        fontsize=9,
        fontweight="bold",
        color="#17365D"
    )

    ax.set_xlim(-0.4, lig_width + 4.35)
    ax.set_ylim(-1.05, top_y + 1.15)

    # Mantener numeración inferior del eje X
    ax.set_xticks([])
    ax.set_yticks([])

    ax.set_title(
        "LIGAMENTO GENERADO – SECUENCIA CONTINUA POR SISTEMA",
        fontweight="bold",
        color="#17365D"
    )

    for sp in ax.spines.values():
        sp.set_visible(False)

    fig.tight_layout()
    return fig

# =========================================================
# SESIÓN PARA MATRICES EDITABLES
# =========================================================

def ensure_leva_state(n_agujas, n_sistemas, fontura_name="Mono"):
    key = f"leva_matrix_{fontura_name}_{n_agujas}_{n_sistemas}"
    if key not in st.session_state:
        df = pd.DataFrame(
            "Anulado",
            index=[f"Aguja {a}" for a in range(1,n_agujas+1)],
            columns=[f"Sistema {s}" for s in range(1,n_sistemas+1)]
        )
        st.session_state[key] = df
    return key

def ensure_needle_state(n_agujas, visible_slots, fontura_name="Mono"):
    key = f"needle_matrix_{fontura_name}_{n_agujas}_{visible_slots}"
    if key not in st.session_state:
        df = pd.DataFrame(
            False,
            index=[f"Aguja {a}" for a in range(1,n_agujas+1)],
            columns=[f"Rep. {p}" for p in range(1,visible_slots+1)]
        )
        st.session_state[key] = df
    return key

def _orden_visual_agujas(n_agujas, fontura_name):
    """Orden visual: Plato/Dial N→1; Cilindro y Mono 1→N."""
    numeros = range(n_agujas, 0, -1) if fontura_name == "Plato" else range(1, n_agujas + 1)
    return [f"Aguja {a}" for a in numeros]


def make_leva_editor(title, n_agujas, n_sistemas, fontura_name):
    st.markdown(f"**{title}**")

    leva_key = ensure_leva_state(n_agujas, n_sistemas, fontura_name)
    # Solo cambia el orden visual de las filas; no cambia Malla/Retención/Anulado/Vacío.
    orden_filas = _orden_visual_agujas(n_agujas, fontura_name)
    df_estado = st.session_state[leva_key].reindex(orden_filas)

    leva_columns = {
        f"Sistema {s}": st.column_config.SelectboxColumn(
            f"Sistema {s}",
            options=["Malla", "Retención", "Anulado", "Vacío"],
            required=True
        )
        for s in range(1,n_sistemas+1)
    }

    df_edit = st.data_editor(
        df_estado,
        column_config=leva_columns,
        use_container_width=True,
        num_rows="fixed",
        key=f"leva_editor_{fontura_name}_{n_agujas}_{n_sistemas}"
    )

    st.session_state[leva_key] = df_edit.copy()

    # Orientación de las levas según fontura.
    # Cilindro: triángulo y trapecio hacia arriba.
    # Plato/Dial: triángulo y trapecio hacia abajo.
    if fontura_name == "Plato":
        symbol_map = {
            "Malla": "▼",
            "Retención": "TRAP",
            "Anulado": "—",
            "Vacío": ""
        }
        caption = "▼ = Malla   TRAP = Retención   — = Anulado / Sin tejido   (Vacío = cuadro sin símbolo)"
    else:
        symbol_map = {
            "Malla": "▲",
            "Retención": "TRAP",
            "Anulado": "—",
            "Vacío": ""
        }
        caption = "▲ = Malla   TRAP = Retención   — = Anulado / Sin tejido   (Vacío = cuadro sin símbolo)"

    df_symbols = df_edit.copy()
    for col in df_symbols.columns:
        df_symbols[col] = df_symbols[col].map(lambda v: symbol_map.get(v, ""))

    st.caption(caption)
    return df_edit, df_symbols

def make_needle_editor(title, n_agujas, visible_slots, fontura_name):
    st.markdown(f"**{title}**")

    needle_key = ensure_needle_state(n_agujas, visible_slots, fontura_name)
    # Mantiene la selección asociada a cada aguja y solo cambia su posición visual.
    orden_filas = _orden_visual_agujas(n_agujas, fontura_name)
    df_estado = st.session_state[needle_key].reindex(orden_filas)

    needle_columns = {
        f"Rep. {p}": st.column_config.CheckboxColumn(
            f"Rep. {p}",
            default=False
        )
        for p in range(1,visible_slots+1)
    }

    df_edit = st.data_editor(
        df_estado,
        column_config=needle_columns,
        use_container_width=True,
        num_rows="fixed",
        key=f"needle_editor_{fontura_name}_{n_agujas}_{visible_slots}"
    )

    st.session_state[needle_key] = df_edit.copy()

    df_symbols = df_edit.copy()
    for col in df_symbols.columns:
        df_symbols[col] = df_symbols[col].map(
            lambda v: "I" if bool(v) else ""
        )

    st.caption("I = Aguja seleccionada")
    return df_edit, df_symbols

# =========================================================
# INTERFAZ
# =========================================================
left, right = st.columns([0.95, 2.05], gap="large")

with left:
    st.markdown('<div class="section-title">1. DATOS GENERALES</div>', unsafe_allow_html=True)

    n_sistemas = st.number_input(
        "N° de sistemas",
        min_value=1,
        max_value=32,
        value=8,
        step=1
    )

    numero_ficha = st.text_input(
        "N° de ficha",
        value="22239",
        placeholder="Ej. 22239"
    )

    tipo_fontura = st.selectbox(
        "Tipo de fontura",
        ["Monofontura", "Doblefontura"],
        index=0
    )

    if tipo_fontura == "Monofontura":
        n_agujas = st.number_input(
            "N° de agujas",
            min_value=1,
            max_value=64,
            value=8,
            step=1
        )
    else:
        st.markdown("**Cantidad de agujas por fontura**")
        col_cil_datos, col_plato_datos = st.columns(2)
        with col_cil_datos:
            n_agujas_cil = st.number_input(
                "N° agujas Cilindro",
                min_value=1,
                max_value=64,
                value=4,
                step=1,
                key="n_agujas_cilindro"
            )
        with col_plato_datos:
            n_agujas_plato = st.number_input(
                "N° agujas Plato / Dial",
                min_value=1,
                max_value=64,
                value=2,
                step=1,
                key="n_agujas_plato"
            )

    visible_slots = st.number_input(
        "Cantidad visible de repeticiones del ligamento",
        min_value=1,
        max_value=40,
        value=15,
        step=1,
        help="Este valor solo controla cuántas repeticiones se muestran en el dibujo del ligamento. No modifica levas ni selección de agujas."
    )

    st.markdown('<div class="section-title">2. LIGAMENTO POR SISTEMA</div>', unsafe_allow_html=True)

    st.markdown(
        '<div class="note">'
        'Escribe una secuencia para cada sistema. Ejemplos: '
        '<b>1-3-5</b>, <b>2-2-4-5</b>, <b>1-1-3-4-5</b>. '
        'El dibujo se genera automáticamente.'
        '</div>',
        unsafe_allow_html=True
    )

    system_sequences = []
    system_labels = []
    system_yarns = []

    for s in range(1, n_sistemas+1):
        st.markdown(f"**Sistema {s}**")

        txt = st.text_input(
            "Secuencia",
            value="1",
            key=f"seq_sistema_{s}",
            placeholder="Ej. 1-3-5"
        )
        system_sequences.append(parse_sequence(txt))

        label = st.text_input(
            "Tipo / descripción",
            value="",
            key=f"label_sistema_{s}",
            placeholder="Ej. JERSEY / VANIZADO / FLOTE"
        )
        system_labels.append(label)

        yarn = st.text_input(
            "Hilo / material",
            value="",
            key=f"yarn_sistema_{s}",
            placeholder="Ej. 30/1 TANGUIS ORG."
        )
        system_yarns.append(yarn)

        st.markdown("---")

    st.markdown('<div class="section-title">3. CATÁLOGO DE DIBUJOS</div>', unsafe_allow_html=True)
    st.pyplot(draw_catalog(), use_container_width=True)

with right:
    # =====================================================
    # LIGAMENTO RESULTANTE
    # =====================================================
    st.markdown('<div class="section-title">4. LIGAMENTO GENERADO</div>', unsafe_allow_html=True)

    st.pyplot(
        draw_system_ligaments(
            system_sequences,
            visible_slots,
            labels=system_labels,
            yarns=system_yarns,
            tipo_fontura=tipo_fontura
        ),
        use_container_width=True
    )

    # =====================================================
    # LEVAS Y AGUJAS SEGÚN FONTURA
    # =====================================================
    st.markdown('<div class="section-title">5. DISPOSICIÓN DE LEVAS – EDITABLE</div>', unsafe_allow_html=True)

    st.markdown(
        '<div class="note">'
        'Cada celda de levas es independiente por Aguja × Sistema. '
        'En doblefontura se trabaja por separado Cilindro y Plato/Dial.'
        '</div>',
        unsafe_allow_html=True
    )

    if tipo_fontura == "Monofontura":
        df_leva_edit, df_leva_symbols = make_leva_editor(
            "Levas – Monofontura",
            n_agujas,
            n_sistemas,
            "Mono"
        )

        st.markdown('<div class="section-title">6. SELECCIÓN DE AGUJAS – EDITABLE</div>', unsafe_allow_html=True)

        st.markdown(
            '<div class="note">'
            'La selección de agujas es independiente de las repeticiones visibles del ligamento. '
            'En monofontura se muestran tantas posiciones como agujas.'
            '</div>',
            unsafe_allow_html=True
        )

        df_needle_edit, df_needle_symbols = make_needle_editor(
            "Agujas – Monofontura",
            n_agujas,
            int(n_agujas),
            "Mono"
        )

        # Aliases para exportación
        leva_exports = [("LEVAS – MONOFONTURA", df_leva_symbols)]
        needle_exports = [("AGUJAS – MONOFONTURA", df_needle_symbols)]

    else:
        st.markdown(
            '<div class="note">'
            'En doblefontura, Cilindro y Plato/Dial usan cantidades independientes de agujas. '
            'Las columnas de levas dependen del N° de sistemas. La selección de agujas usa tantas posiciones '
            'como agujas tenga cada fontura. La cantidad visible de repeticiones solo afecta el ligamento. '
            'En levas, selecciona <b>Vacío</b> cuando necesites dejar el cuadro sin símbolo.'
            '</div>',
            unsafe_allow_html=True
        )

        tab_plato, tab_cil = st.tabs(["Plato / Dial", "Cilindro"])

        with tab_plato:
            df_leva_plato_edit, df_leva_plato_symbols = make_leva_editor(
                "Levas – Plato / Dial",
                int(n_agujas_plato),
                n_sistemas,
                "Plato"
            )

            df_needle_plato_edit, df_needle_plato_symbols = make_needle_editor(
                "Agujas / Dial",
                int(n_agujas_plato),
                int(n_agujas_plato),
                "Plato"
            )

        with tab_cil:
            df_leva_cil_edit, df_leva_cil_symbols = make_leva_editor(
                "Levas – Cilindro",
                int(n_agujas_cil),
                n_sistemas,
                "Cilindro"
            )

            df_needle_cil_edit, df_needle_cil_symbols = make_needle_editor(
                "Agujas / Cilindro",
                int(n_agujas_cil),
                int(n_agujas_cil),
                "Cilindro"
            )

        leva_exports = [
            ("LEVAS – PLATO / DIAL", df_leva_plato_symbols),
            ("LEVAS – CILINDRO", df_leva_cil_symbols)
        ]

        needle_exports = [
            ("AGUJAS / DIAL (PLATO)", df_needle_plato_symbols),
            ("AGUJAS / CILINDRO", df_needle_cil_symbols)
        ]

    # =====================================================
    # RESUMEN
    # =====================================================
    st.markdown('<div class="section-title">7. RESUMEN DE LIGAMENTO</div>', unsafe_allow_html=True)

    resumen = pd.DataFrame({
        "Sistema": list(range(1,n_sistemas+1)),
        "Secuencia": ["-".join(map(str,seq)) for seq in system_sequences],
        "Tipo / descripción": system_labels,
        "Hilo / material": system_yarns
    })

    st.dataframe(resumen, use_container_width=True, hide_index=True)


def leva_dataframe_to_png(df, title, orientation="up"):
    from matplotlib.patches import Polygon

    rows, cols = df.shape
    fig_w = max(7.5, cols * 1.15 + 2.2)
    fig_h = max(3.0, rows * 0.55 + 1.8)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h), facecolor="white")
    ax.set_facecolor("white")
    ax.set_xlim(0, cols + 1)
    ax.set_ylim(0, rows + 1)
    ax.axis("off")
    ax.set_title(title, fontsize=13, fontweight="bold", color="#17365D", pad=14)

    # grid
    for c in range(cols + 2):
        ax.plot([c, c], [0, rows + 1], color="#808080", lw=0.7)
    for r in range(rows + 2):
        ax.plot([0, cols + 1], [r, r], color="#808080", lw=0.7)

    # headers
    for j, col in enumerate(df.columns, start=1):
        ax.text(j + 0.5, rows + 0.5, str(col), ha="center", va="center",
                fontsize=8, fontweight="bold")
    for i, idx in enumerate(df.index):
        y = rows - i - 0.5
        ax.text(0.5, y, str(idx), ha="center", va="center",
                fontsize=8, fontweight="bold")

        for j, val in enumerate(df.iloc[i], start=1):
            x = j + 0.5
            sval = str(val).strip()

            if sval in ("▲", "▼"):
                if orientation == "down" or sval == "▼":
                    pts = [(x-0.16, y+0.12), (x+0.16, y+0.12), (x, y-0.16)]
                else:
                    pts = [(x-0.16, y-0.12), (x+0.16, y-0.12), (x, y+0.16)]
                ax.add_patch(Polygon(pts, closed=True, facecolor="black", edgecolor="black"))

            elif sval in ("TRAP", "⏢", "⏥"):
                if orientation == "down":
                    # Trapecio hacia abajo: ancho arriba, angosto abajo
                    pts = [(x-0.18, y+0.14), (x+0.18, y+0.14),
                           (x+0.11, y-0.14), (x-0.11, y-0.14)]
                else:
                    # Trapecio hacia arriba: angosto arriba, ancho abajo
                    pts = [(x-0.11, y+0.14), (x+0.11, y+0.14),
                           (x+0.18, y-0.14), (x-0.18, y-0.14)]
                ax.add_patch(Polygon(pts, closed=True, fill=False,
                                     edgecolor="black", linewidth=1.7))

            elif sval in ("—", "-", "–"):
                ax.plot([x-0.18, x+0.18], [y, y], color="black", lw=1.8)

            elif sval:
                ax.text(x, y, sval, ha="center", va="center", fontsize=8)

    fig.tight_layout()
    return fig

def dataframe_to_png(df, title, font_size=8):
    # Render de tabla a PNG usando matplotlib
    rows, cols = df.shape
    fig_w = max(7.5, cols * 1.15 + 2.2)
    fig_h = max(3.0, rows * 0.48 + 1.8)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h), facecolor="white")
    ax.axis("off")
    ax.set_title(title, fontsize=13, fontweight="bold", color="#17365D", pad=14)

    table = ax.table(
        cellText=df.values,
        rowLabels=df.index,
        colLabels=df.columns,
        cellLoc="center",
        rowLoc="center",
        loc="center"
    )

    table.auto_set_font_size(False)
    table.set_fontsize(font_size)
    table.scale(1.0, 1.45)

    # Cabeceras
    for (r, c), cell in table.get_celld().items():
        cell.set_edgecolor("#808080")
        cell.set_linewidth(0.6)
        if r == 0 or c == -1:
            cell.set_text_props(weight="bold")
            cell.set_facecolor("#F2F2F2")
        else:
            cell.set_facecolor("white")

    fig.tight_layout()
    return fig

def generar_zip_png():
    zip_buffer = BytesIO()

    # 1. Ligamento
    lig_buffer = BytesIO()
    fig_lig = draw_system_ligaments(
        system_sequences,
        visible_slots,
        labels=system_labels,
        yarns=system_yarns,
        tipo_fontura=tipo_fontura
    )
    fig_lig.savefig(
        lig_buffer,
        format="png",
        dpi=200,
        bbox_inches="tight",
        facecolor="white"
    )
    plt.close(fig_lig)
    lig_buffer.seek(0)

    # 2. Levas
    leva_buffer = BytesIO()

    if tipo_fontura == "Doblefontura":
        # Un PNG con Plato/Dial y Cilindro apilados
        fig_h = 4.0 + 0.5 * (len(df_leva_cil_symbols) + len(df_leva_plato_symbols))
        fig, axes = plt.subplots(
            2, 1,
            figsize=(max(8.5, n_sistemas*1.1 + 2.5), fig_h),
            facecolor="white"
        )

        for ax, df, title, orient in [
            (axes[0], df_leva_plato_symbols, "LEVAS - PLATO / DIAL", "down"),
            (axes[1], df_leva_cil_symbols, "LEVAS - CILINDRO", "up"),
        ]:
            ax.axis("off")
            ax.set_title(title, fontsize=12, fontweight="bold", color="#17365D", pad=10)

            rows, cols = df.shape
            ax.set_xlim(0, cols + 1)
            ax.set_ylim(0, rows + 1)

            for c in range(cols + 2):
                ax.plot([c, c], [0, rows + 1], color="#808080", lw=0.7)
            for r in range(rows + 2):
                ax.plot([0, cols + 1], [r, r], color="#808080", lw=0.7)

            for j, col in enumerate(df.columns, start=1):
                ax.text(j+0.5, rows+0.5, str(col), ha="center", va="center",
                        fontsize=8, fontweight="bold")
            for i, idx in enumerate(df.index):
                y = rows-i-0.5
                ax.text(0.5, y, str(idx), ha="center", va="center",
                        fontsize=8, fontweight="bold")
                for j, val in enumerate(df.iloc[i], start=1):
                    x = j+0.5
                    sval = str(val).strip()

                    if sval in ("▲","▼"):
                        if orient == "down":
                            pts = [(x-0.16,y+0.12),(x+0.16,y+0.12),(x,y-0.16)]
                        else:
                            pts = [(x-0.16,y-0.12),(x+0.16,y-0.12),(x,y+0.16)]
                        ax.add_patch(Polygon(pts, closed=True, facecolor="black", edgecolor="black"))

                    elif sval in ("TRAP","⏢","⏥"):
                        if orient == "down":
                            # Plato/Dial: trapecio hacia abajo
                            pts = [(x-0.18,y+0.14),(x+0.18,y+0.14),
                                   (x+0.11,y-0.14),(x-0.11,y-0.14)]
                        else:
                            # Cilindro: trapecio hacia arriba
                            pts = [(x-0.11,y+0.14),(x+0.11,y+0.14),
                                   (x+0.18,y-0.14),(x-0.18,y-0.14)]
                        ax.add_patch(Polygon(pts, closed=True, fill=False,
                                             edgecolor="black", linewidth=1.7))

                    elif sval in ("—","-","–"):
                        ax.plot([x-0.18,x+0.18],[y,y],color="black",lw=1.8)

        fig.tight_layout()
        fig.savefig(
            leva_buffer,
            format="png",
            dpi=200,
            bbox_inches="tight",
            facecolor="white"
        )
        plt.close(fig)
    else:
        fig = leva_dataframe_to_png(df_leva_symbols, "DISPOSICIÓN DE LEVAS", orientation="up")
        fig.savefig(
            leva_buffer,
            format="png",
            dpi=200,
            bbox_inches="tight",
            facecolor="white"
        )
        plt.close(fig)

    leva_buffer.seek(0)

    # 3. Agujas
    aguja_buffer = BytesIO()

    if tipo_fontura == "Doblefontura":
        # Dos tablas separadas en una sola imagen: Plato/Dial arriba y Cilindro abajo.
        tablas_ag = [
            ("AGUJAS / DIAL (PLATO)", df_needle_plato_symbols),
            ("AGUJAS / CILINDRO", df_needle_cil_symbols),
        ]
        fig_h = 4.0 + 0.48 * (len(df_needle_plato_symbols) + len(df_needle_cil_symbols))
        fig_w = max(8.5, max(df_needle_plato_symbols.shape[1], df_needle_cil_symbols.shape[1]) * 1.15 + 2.5)
        fig_ag, axes = plt.subplots(2, 1, figsize=(fig_w, fig_h), facecolor="white")

        for ax, (titulo_ag, df_ag) in zip(axes, tablas_ag):
            ax.axis("off")
            ax.set_title(titulo_ag, fontsize=12, fontweight="bold", color="#17365D", pad=10)
            table = ax.table(
                cellText=df_ag.fillna("").values,
                rowLabels=df_ag.index,
                colLabels=df_ag.columns,
                cellLoc="center",
                rowLoc="center",
                loc="center"
            )
            table.auto_set_font_size(False)
            table.set_fontsize(8)
            table.scale(1.0, 1.35)
            for (r, c), cell in table.get_celld().items():
                cell.set_edgecolor("#808080")
                cell.set_linewidth(0.6)
                if r == 0 or c == -1:
                    cell.set_text_props(weight="bold")
                    cell.set_facecolor("#F2F2F2")
                else:
                    cell.set_facecolor("white")

        fig_ag.tight_layout()
    else:
        titulo_ag, df_ag_export = needle_exports[0]
        fig_ag = dataframe_to_png(df_ag_export.fillna(""), titulo_ag)

    fig_ag.savefig(
        aguja_buffer,
        format="png",
        dpi=200,
        bbox_inches="tight",
        facecolor="white"
    )
    plt.close(fig_ag)
    aguja_buffer.seek(0)

    ficha = str(numero_ficha).strip() or "ficha"

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"{ficha}-1.png", lig_buffer.getvalue())
        zf.writestr(f"{ficha}-2.png", leva_buffer.getvalue())
        zf.writestr(f"{ficha}-3.png", aguja_buffer.getvalue())

    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# =========================================================
# EXPORTAR EXCEL
# =========================================================
def generar_excel():
    bio = BytesIO()
    wb = Workbook(bio, {"in_memory":True})
    ws = wb.add_worksheet("Resultado")

    fmt_title = wb.add_format({
        "bold":True,
        "font_size":15,
        "align":"center",
        "bg_color":"#D9EAD3",
        "font_color":"#17365D",
        "border":1
    })

    fmt_h = wb.add_format({
        "bold":True,
        "align":"center",
        "bg_color":"#F2F2F2",
        "border":1
    })

    fmt_c = wb.add_format({
        "align":"center",
        "border":1
    })

    # Ligamento como imagen
    img_lig = BytesIO()
    fig = draw_system_ligaments(
        system_sequences,
        visible_slots,
        labels=system_labels,
        yarns=system_yarns,
        tipo_fontura=tipo_fontura
    )
    fig.savefig(
        img_lig,
        format="png",
        dpi=180,
        bbox_inches="tight",
        facecolor="white"
    )
    plt.close(fig)
    img_lig.seek(0)

    ws.merge_range(
        "A1:J2",
        "GENERADOR AUTOMÁTICO DE LIGAMENTOS Y LEVAS – V5.9",
        fmt_title
    )

    ws.write("A3","N° de ficha",fmt_h)
    ws.write("B3",numero_ficha,fmt_c)

    ws.write("A4","N° sistemas",fmt_h)
    ws.write("B4",n_sistemas,fmt_c)

    ws.write("A5","Tipo de fontura",fmt_h)
    ws.write("B5",tipo_fontura,fmt_c)
    if tipo_fontura == "Doblefontura":
        ws.write("A6","N° agujas Cilindro",fmt_h)
        ws.write("B6",int(n_agujas_cil),fmt_c)
        ws.write("A7","N° agujas Plato / Dial",fmt_h)
        ws.write("B7",int(n_agujas_plato),fmt_c)
        ws.write("A8","Repeticiones visibles ligamento",fmt_h)
        ws.write("B8",int(visible_slots),fmt_c)
        sec_start = 9
    else:
        ws.write("A6","N° agujas",fmt_h)
        ws.write("B6",int(n_agujas),fmt_c)
        ws.write("A7","Repeticiones visibles ligamento",fmt_h)
        ws.write("B7",int(visible_slots),fmt_c)
        sec_start = 8

    # Secuencias
    ws.write(sec_start-1,0,"Sistema",fmt_h)
    ws.write(sec_start-1,1,"Secuencia",fmt_h)
    ws.write(sec_start-1,2,"Tipo / descripción",fmt_h)
    ws.write(sec_start-1,3,"Hilo / material",fmt_h)

    for i,seq in enumerate(system_sequences):
        ws.write(sec_start+i,0,i+1,fmt_c)
        ws.write(sec_start+i,1,"-".join(map(str,seq)),fmt_c)
        ws.write(sec_start+i,2,system_labels[i],fmt_c)
        ws.write(sec_start+i,3,system_yarns[i],fmt_c)

    img_row = sec_start + n_sistemas + 1

    ws.merge_range(img_row,0,img_row,9,"LIGAMENTO GENERADO",fmt_title)
    ws.insert_image(
        img_row+1,0,
        "ligamento.png",
        {
            "image_data":img_lig,
            "x_scale":0.66,
            "y_scale":0.66
        }
    )

    r = img_row+28

    # Levas
    for title, df_export in leva_exports:
        ws.write(r,0,title,fmt_h)
        r += 1
        ws.write(r,0,"Aguja",fmt_h)

        for j,col in enumerate(df_export.columns,1):
            ws.write(r,j,col,fmt_h)

        for i in range(len(df_export)):
            ws.write(r+i+1,0,df_export.index[i],fmt_h)
            for j in range(len(df_export.columns)):
                ws.write(
                    r+i+1,
                    j+1,
                    ("▱" if str(df_export.iloc[i,j]).strip() == "TRAP" else df_export.iloc[i,j]),
                    fmt_c
                )

        r += len(df_export)+3

    # Selección de agujas
    for title, df_export in needle_exports:
        ws.write(r,0,title,fmt_h)
        r += 1
        ws.write(r,0,"Aguja",fmt_h)

        for j,col in enumerate(df_export.columns,1):
            ws.write(r,j,col,fmt_h)

        for i in range(len(df_export)):
            ws.write(r+i+1,0,df_export.index[i],fmt_h)
            for j in range(len(df_export.columns)):
                ws.write(
                    r+i+1,
                    j+1,
                    df_export.iloc[i,j],
                    fmt_c
                )

        r += len(df_export)+3

    ws.set_column(0,0,18)
    ws.set_column(1,1,18)
    ws.set_column(2,2,22)
    ws.set_column(3,3,30)
    ws.set_column(4,max(4,n_sistemas),12)

    wb.close()
    bio.seek(0)
    return bio.getvalue()

st.divider()

st.download_button(
    "⬇️ EXPORTAR RESULTADO A EXCEL",
    data=generar_excel(),
    file_name=f"{str(numero_ficha).strip() or 'ficha'}-ligamento.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


st.download_button(
    "🗂️ EXPORTAR 3 PNG EN ZIP",
    data=generar_zip_png(),
    file_name=f"{str(numero_ficha).strip() or 'ficha'}_imagenes.zip",
    mime="application/zip"
)

st.info(
    "V5.9: Plato/Dial = agujas en orden descendente y levas hacia abajo; "
    "Cilindro = agujas en orden ascendente y levas hacia arriba. "
    "La imagen de agujas muestra ambas tablas por separado."
)
