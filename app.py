
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import zipfile
from matplotlib.patches import Circle, Polygon
from io import BytesIO
from PIL import Image, ImageChops
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
st.caption("Versión 5.19 · Minijacquard configurable por patrones · Agujas sin bloques · Bloques verticales")

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
            if 1 <= v <= 7:
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
        # Referencia de ficha: línea horizontal continua + zigzag.
        # La parte BAJA del zigzag coincide con el centro de cada círculo
        # sobre la línea base; los picos quedan ENTRE círculos.
        ax.plot([x0, x1], [y0, y0], color="black", lw=lw)

        # V con valle exactamente en la línea base (sobre el círculo).
        # Al repetirse, los extremos altos se unen y forman los picos
        # entre una aguja y la siguiente, como en la ficha de referencia.
        ax.plot(
            [x0, xm, x1],
            [y0 + height*0.62, y0, y0 + height*0.62],
            color="black",
            lw=lw
        )

        # Círculo con el tamaño normal del catálogo, tangente a la línea.
        ax.add_patch(
            Circle(
                (xm, y0-r),
                r,
                fill=False,
                color="black",
                lw=lw
            )
        )

    elif sym == 7:
        # Línea horizontal con flecha hacia la derecha.
        # Se mantiene dentro de un módulo para que pueda repetirse en el ligamento.
        ax.annotate(
            "",
            xy=(x1, y0),
            xytext=(x0, y0),
            arrowprops=dict(arrowstyle="->", color="black", lw=lw, shrinkA=0, shrinkB=0)
        )

def draw_catalog():
    fig, ax = plt.subplots(figsize=(5.2, 5.1), facecolor="white")
    ax.set_facecolor("white")

    for i, sym in enumerate([1,2,3,4,5,6,7]):
        y = 6-i
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
    ax.set_ylim(-0.6,6.8)
    ax.axis("off")
    fig.tight_layout()
    return fig

def draw_system_ligaments(system_sequences, visible_slots=12, labels=None, yarns=None,
                          tipo_fontura="Monofontura", max_sistemas_bloque=8):
    """Dibuja el ligamento en bloques de 8 sistemas.

    Excepción: si la ficha tiene entre 1 y 10 sistemas, mantiene un solo bloque.
    Desde 11 sistemas, crea bloques horizontales de máximo 8
    (1-8, 9-16, 17-24, ...), conservando la numeración real.
    """
    n = len(system_sequences)
    labels = labels or [""] * n
    yarns = yarns or [""] * n

    module_w = 0.84
    symbol_h = 0.60
    row_gap = 1.08
    lig_width = visible_slots * module_w

    # Si existen descripciones/materiales, se mantienen dentro de cada bloque.
    hay_texto = any(str(x).strip() for x in labels) or any(str(x).strip() for x in yarns)
    desc_col_w = 2.55 if hay_texto else 0.0
    material_col_w = 2.05 if hay_texto else 0.0
    text_gap = 0.72 if hay_texto else 0.0
    block_gap = 1.05
    block_inner_w = lig_width + text_gap + desc_col_w + material_col_w

    # Hasta 10 sistemas se conserva un único bloque. Desde 11, bloques de 8.
    tam_bloque = n if n <= 10 else max_sistemas_bloque
    n_bloques = max(1, (n + tam_bloque - 1) // tam_bloque)
    filas_max = min(tam_bloque, n)
    fig_w = max(9.2, n_bloques * block_inner_w + (n_bloques - 1) * block_gap + 0.9)
    fig_h = max(2.6, filas_max * 0.95 + 0.90)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h), facecolor="white")
    ax.set_facecolor("white")

    for b in range(n_bloques):
        ini = b * tam_bloque
        fin = min(ini + tam_bloque, n)
        cantidad = fin - ini
        x_base = b * (block_inner_w + block_gap)
        x_text1 = x_base + lig_width + text_gap
        x_text2 = x_text1 + desc_col_w + material_col_w / 2
        x_right = x_base + block_inner_w

        # Orden ascendente visual de arriba hacia abajo: 1, 2, 3... / 13, 14, 15...
        for local_i, global_i in enumerate(range(ini, fin)):
            y = (cantidad - 1 - local_i) * row_gap
            seq = system_sequences[global_i]

            ax.text(x_base - 0.34, y, str(global_i + 1), ha="center", va="center",
                    fontsize=18, fontweight="bold", color="#111111")

            for pos in range(visible_slots):
                sym = seq[pos % len(seq)]
                draw_ligament_symbol(ax, sym, x0=x_base + pos * module_w, y0=y,
                                     width=module_w, height=symbol_h, lw=2.6)

            if hay_texto:
                ax.text(x_text1, y, str(labels[global_i]) if global_i < len(labels) else "",
                        ha="left", va="center", fontsize=17.0, fontweight="bold",
                        linespacing=1.05, color="#111111")
                ax.text(x_text2, y, str(yarns[global_i]) if global_i < len(yarns) else "",
                        ha="center", va="center", fontsize=16.0, fontweight="bold", color="#111111")

            ax.plot([x_base - 0.05, x_right], [y - row_gap * 0.50, y - row_gap * 0.50],
                    color="#E9E9E9", lw=0.55)

        top_y = (cantidad - 1) * row_gap
        if hay_texto:
            ax.plot([x_base + lig_width + 0.38, x_base + lig_width + 0.38],
                    [-0.42, top_y + 0.62], color="#EFEFEF", lw=0.6)
            ax.plot([x_text1 + desc_col_w, x_text1 + desc_col_w],
                    [-0.42, top_y + 0.62], color="#EFEFEF", lw=0.6)
            ax.text(x_text1, top_y + 0.50, "TIPO / DESCRIPCIÓN", ha="left", va="bottom",
                    fontsize=14.0, fontweight="bold", color="#17365D")
            ax.text(x_text2, top_y + 0.50, "HILO / MATERIAL", ha="center", va="bottom",
                    fontsize=14.0, fontweight="bold", color="#17365D")

    total_w = n_bloques * block_inner_w + (n_bloques - 1) * block_gap
    ax.set_xlim(-0.62, total_w + 0.15)
    ax.set_ylim(-0.48, (filas_max - 1) * row_gap + 0.68)
    ax.set_xticks([])
    ax.set_yticks([])
    for sp in ax.spines.values():
        sp.set_visible(False)
    fig.tight_layout(pad=0.08)
    return fig

def draw_minijacquard_ligament(pattern_sequences, pattern_assignments, visible_slots=6):
    """Dibuja Minijacquard por patrones configurables, independientes del N° real de sistemas."""
    n_pat = max(1, len(pattern_sequences))
    module_w = 0.90
    symbol_h = 0.68
    slots = max(1, min(int(visible_slots), 12))
    lig_width = slots * module_w
    row_gap = 1.35
    y_top = (n_pat - 1) * row_gap + 1.25
    footer_h = max(0.85, n_pat * 0.30 + 0.35)
    fig_h = max(3.4, n_pat * 1.25 + footer_h)
    fig, ax = plt.subplots(figsize=(10.5, fig_h), facecolor="white")
    ax.set_facecolor("white")

    for i, seq in enumerate(pattern_sequences):
        num = i + 1
        y = y_top - i * row_gap
        seq = seq or [1]
        ax.text(-1.10, y, str(num), ha="center", va="center", fontsize=20, fontweight="bold", color="black")
        ax.text(-0.35, y, "{", ha="center", va="center", fontsize=56, fontweight="normal", color="black")
        for pos in range(slots):
            sym = seq[pos % len(seq)]
            draw_ligament_symbol(ax, sym, x0=pos*module_w, y0=y, width=module_w, height=symbol_h, lw=2.3)

    footer_top = 0.25
    for i, asignacion in enumerate(pattern_assignments):
        y = footer_top - i * 0.30
        texto = str(asignacion).strip() or f"PATRÓN {i+1}"
        ax.text(-0.10, y, texto.upper(), ha="left", va="center", fontsize=12.5, color="black")
        ax.text(lig_width+0.75, y, str(i+1), ha="center", va="center", fontsize=13, fontweight="bold")

    line_y = footer_top - max(1, n_pat) * 0.30 - 0.12
    ax.plot([-1.25, lig_width], [line_y, line_y], color="black", lw=1.5)
    ax.set_xlim(-1.55, lig_width+1.25)
    ax.set_ylim(line_y-0.18, y_top+0.70)
    ax.axis("off")
    fig.tight_layout(pad=0.15)
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
    """Orden visual ascendente normal para todas las fonturas: 1→N."""
    numeros = range(1, n_agujas + 1)
    return [f"Aguja {a}" for a in numeros]


def make_leva_editor(title, n_agujas, n_sistemas, fontura_name):
    st.markdown(f"**{title}**")

    leva_key = ensure_leva_state(n_agujas, n_sistemas, fontura_name)
    # Solo cambia el orden visual de las filas; no cambia Malla/Retención/Anulado/Vacío.
    orden_filas = _orden_visual_agujas(n_agujas, fontura_name)
    df_estado = st.session_state[leva_key].reindex(orden_filas)

    leva_columns = {
        f"Sistema {s}": st.column_config.SelectboxColumn(
            f"S{s}",
            options=["Malla", "Retención", "Anulado", "Vacío"],
            required=True
        )
        for s in range(1,n_sistemas+1)
    }

    # Editor con clave estable por fontura. Mantener una key que cambia con las
    # dimensiones puede hacer que React intente desmontar un componente que ya
    # fue reemplazado durante el rerun (removeChild / NotFoundError).
    editor_key = f"leva_editor_{fontura_name}"

    # Si cambió la forma de la matriz, limpiamos únicamente el estado interno
    # del widget antes de volver a dibujarlo. La matriz real sigue guardada en
    # leva_key y conserva la lógica Aguja × Sistema.
    shape_key = f"{editor_key}_shape"
    current_shape = (int(n_agujas), int(n_sistemas))
    if st.session_state.get(shape_key) != current_shape:
        st.session_state.pop(editor_key, None)
        st.session_state[shape_key] = current_shape

    df_edit = st.data_editor(
        df_estado,
        column_config=leva_columns,
        use_container_width=True,
        num_rows="fixed",
        key=editor_key
    )

    # Actualizar la matriz persistente sin sustituir el estado interno del
    # data_editor. Reindexamos al orden lógico original para que Plato/Dial
    # pueda mostrarse N→1 sin alterar los datos guardados.
    base_index = st.session_state[leva_key].index
    st.session_state[leva_key] = df_edit.reindex(base_index).copy()

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

    # Misma estrategia de estabilidad que en el editor de levas.
    editor_key = f"needle_editor_{fontura_name}"
    shape_key = f"{editor_key}_shape"
    current_shape = (int(n_agujas), int(visible_slots))
    if st.session_state.get(shape_key) != current_shape:
        st.session_state.pop(editor_key, None)
        st.session_state[shape_key] = current_shape

    df_edit = st.data_editor(
        df_estado,
        column_config=needle_columns,
        use_container_width=True,
        num_rows="fixed",
        key=editor_key
    )

    base_index = st.session_state[needle_key].index
    st.session_state[needle_key] = df_edit.reindex(base_index).copy()

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
        max_value=96,
        value=8,
        step=1
    )

    numero_ficha = st.text_input(
        "N° de ficha",
        value="22239",
        placeholder="Ej. 22239"
    )

    numero_item = st.number_input(
        "Ítem",
        min_value=1,
        max_value=99,
        value=1,
        step=1,
        help="Se agrega al nombre de los archivos exportados."
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

    needle_repetitions = st.number_input(
        "N° de agujas por rapport",
        min_value=1,
        max_value=40,
        value=12,
        step=1,
        help="Controla únicamente las columnas Rep. de la tabla de selección de agujas. Es independiente del N° de agujas y de las repeticiones visibles del ligamento."
    )

    tipo_seleccion_agujas = st.selectbox(
        "Tipo de selección de agujas",
        ["Normal", "Minijacquard"],
        help="Minijacquard permite marcar una matriz Agujas × Sistemas para formar dibujos como rombos."
    )
    if tipo_seleccion_agujas == "Minijacquard":
        agrupacion_minijacquard = st.number_input(
            "Agrupar agujas de", min_value=1, max_value=8, value=2, step=1,
            help="Ej.: con 48 agujas y agrupación 2 se muestran 1-2, 3-4, ... 47-48."
        )

        st.markdown("**Configuración del ligamento Minijacquard**")
        n_patrones_mj = st.number_input(
            "N° de dibujos / patrones de ligamento",
            min_value=1, max_value=12, value=2, step=1,
            help="No es el N° de sistemas. Indica cuántos dibujos distintos aparecen en la ficha."
        )
        minijacquard_pattern_sequences = []
        minijacquard_pattern_assignments = []
        for ptn in range(1, int(n_patrones_mj) + 1):
            st.markdown(f"**Dibujo {ptn}**")
            c1, c2 = st.columns([0.8, 1.2])
            with c1:
                seq_ptn = st.text_input(
                    "Secuencia del dibujo",
                    value="2-1" if ptn == 1 else ("1" if ptn == 2 else "1"),
                    key=f"mj_pattern_seq_{ptn}",
                    placeholder="Ej. 2-1-2-1"
                )
            with c2:
                default_asig = "Sistemas impares" if ptn == 1 else ("Sistemas pares" if ptn == 2 else f"Sistemas del dibujo {ptn}")
                asig_ptn = st.text_input(
                    "Asignación de sistemas",
                    value=default_asig,
                    key=f"mj_pattern_asig_{ptn}",
                    placeholder="Ej. Sistemas impares / S1-S5-S9-S13"
                )
            minijacquard_pattern_sequences.append(parse_sequence(seq_ptn))
            minijacquard_pattern_assignments.append(asig_ptn)


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

        label = st.text_area(
            "Tipo / descripción",
            value="",
            key=f"label_sistema_{s}",
            placeholder="Ej. Rizo: 30/1 COTTON\nVanizado: 75/72 PES",
            height=80
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

    fig_lig_preview = (
        draw_minijacquard_ligament(minijacquard_pattern_sequences, minijacquard_pattern_assignments, visible_slots)
        if tipo_seleccion_agujas == "Minijacquard"
        else draw_system_ligaments(
            system_sequences, visible_slots, labels=system_labels, yarns=system_yarns,
            tipo_fontura=tipo_fontura
        )
    )
    st.pyplot(fig_lig_preview, use_container_width=True)

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
            'La cantidad de columnas Rep. se define de forma independiente con “N° de agujas por rapport”.'
            '</div>',
            unsafe_allow_html=True
        )

        if tipo_seleccion_agujas == "Minijacquard":
            grupos = []
            g = int(agrupacion_minijacquard)
            for a in range(1, int(n_agujas) + 1, g):
                grupos.append(f"{a}-{min(a+g-1, int(n_agujas))}")
            mj_key = f"minijacquard_{int(n_agujas)}_{int(n_sistemas)}_{g}"
            if mj_key not in st.session_state:
                st.session_state[mj_key] = pd.DataFrame(False, index=grupos, columns=[f"S{s}" for s in range(1, int(n_sistemas)+1)])
            mj_cols = {c: st.column_config.CheckboxColumn(c, default=False) for c in st.session_state[mj_key].columns}
            df_needle_edit = st.data_editor(st.session_state[mj_key], column_config=mj_cols, use_container_width=True, num_rows="fixed", key=f"mj_editor_{int(n_agujas)}_{int(n_sistemas)}_{g}")
            st.session_state[mj_key] = df_needle_edit.copy()
            df_needle_symbols = df_needle_edit.copy()
            for _col in df_needle_symbols.columns:
                df_needle_symbols[_col] = df_needle_symbols[_col].map(lambda v: "X" if bool(v) else "")
            st.caption("X = grupo de agujas seleccionado en ese sistema")
        else:
            df_needle_edit, df_needle_symbols = make_needle_editor(
                "Agujas – Monofontura",
                n_agujas,
                int(needle_repetitions),
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
                int(needle_repetitions),
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
                int(needle_repetitions),
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


def _rangos_bloques_columnas(n_cols, max_bloque=8):
    """1-10 columnas: un bloque. Desde 11: bloques de máximo 8."""
    if n_cols <= 10:
        return [(0, n_cols)]
    return [(i, min(i + max_bloque, n_cols)) for i in range(0, n_cols, max_bloque)]


def _dibujar_leva_bloque(ax, df, orientation="up"):
    from matplotlib.patches import Polygon
    rows, cols = df.shape
    ax.set_xlim(0, cols + 1)
    ax.set_ylim(0, rows + 1)
    ax.axis("off")

    for c in range(cols + 2):
        ax.plot([c, c], [0, rows + 1], color="#808080", lw=1.0)
    for r in range(rows + 2):
        ax.plot([0, cols + 1], [r, r], color="#808080", lw=1.0)

    for j, col in enumerate(df.columns, start=1):
        ax.text(j + 0.5, rows + 0.5, str(col).replace("Sistema ", "S"),
                ha="center", va="center", fontsize=16, fontweight="bold")
    for i, idx in enumerate(df.index):
        y = rows - i - 0.5
        ax.text(0.5, y, str(idx), ha="center", va="center", fontsize=16, fontweight="bold")
        for j, val in enumerate(df.iloc[i], start=1):
            x = j + 0.5
            sval = str(val).strip()
            if sval in ("▲", "▼"):
                if orientation == "down" or sval == "▼":
                    pts = [(x-0.18,y+0.15),(x+0.18,y+0.15),(x,y-0.18)]
                else:
                    pts = [(x-0.18,y-0.15),(x+0.18,y-0.15),(x,y+0.18)]
                ax.add_patch(Polygon(pts, closed=True, facecolor="black", edgecolor="black"))
            elif sval in ("TRAP", "⏢", "⏥"):
                if orientation == "down":
                    pts = [(x-0.20,y+0.16),(x+0.20,y+0.16),(x+0.12,y-0.16),(x-0.12,y-0.16)]
                else:
                    pts = [(x-0.12,y+0.16),(x+0.12,y+0.16),(x+0.20,y-0.16),(x-0.20,y-0.16)]
                ax.add_patch(Polygon(pts, closed=True, fill=False, edgecolor="black", linewidth=2.0))
            elif sval in ("—", "-", "–"):
                ax.plot([x-0.20, x+0.20], [y, y], color="black", lw=2.0)


def leva_dataframe_to_png(df, title, orientation="up"):
    # Los bloques se apilan verticalmente: S1-S8 arriba, S9-S16 debajo, etc.
    rangos = _rangos_bloques_columnas(df.shape[1])
    n_b = len(rangos)
    rows = df.shape[0]
    max_cols = max((b-a) for a,b in rangos)
    fig_w = max(7.5, (max_cols + 1) * 1.10)
    fig_h = max(3.2, n_b * (rows * 0.78 + 1.05) + 0.8)
    fig, axes = plt.subplots(n_b, 1, figsize=(fig_w, fig_h), facecolor="white", squeeze=False)
    fig.suptitle(title, fontsize=20, fontweight="bold", color="#17365D", y=0.995)
    for r, (a,b) in enumerate(rangos):
        _dibujar_leva_bloque(axes[r][0], df.iloc[:, a:b], orientation)
    fig.subplots_adjust(left=0.03, right=0.99, bottom=0.025, top=0.94, hspace=0.28)
    return fig


def _dibujar_agujas_bloque(ax, df, font_size=16):
    ax.axis("off")
    table = ax.table(
        cellText=df.fillna("").values,
        rowLabels=df.index,
        colLabels=df.columns,
        cellLoc="center", rowLoc="center", loc="center",
        bbox=[0.0, 0.02, 1.0, 0.96]
    )
    table.auto_set_font_size(False)
    table.set_fontsize(font_size)
    for (r,c), cell in table.get_celld().items():
        cell.set_edgecolor("#808080")
        cell.set_linewidth(1.2)
        if r == 0 or c == -1:
            cell.set_text_props(weight="bold")
            cell.set_facecolor("#F2F2F2")
        else:
            cell.set_facecolor("white")
            if str(cell.get_text().get_text()).strip() == "I":
                cell.set_text_props(weight="bold", fontsize=font_size+2)


def minijacquard_dataframe_to_png(df, title, font_size=14):
    """Minijacquard siempre se exporta como una sola tabla continua, sin bloques."""
    rows, cols = df.shape
    fig_w = max(9.0, (cols + 1) * 0.62)
    fig_h = max(3.4, rows * 0.62 + 1.35)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), facecolor="white")
    fig.suptitle(title, fontsize=20, fontweight="bold", color="#17365D", y=0.99)
    _dibujar_agujas_bloque(ax, df, font_size=font_size)
    fig.subplots_adjust(left=0.04, right=0.995, bottom=0.04, top=0.90)
    return fig

def dataframe_to_png(df, title, font_size=16):
    # Igual que levas: cada bloque va debajo del anterior, nunca al costado.
    rangos = _rangos_bloques_columnas(df.shape[1])
    n_b = len(rangos)
    rows = df.shape[0]
    max_cols = max((b-a) for a,b in rangos)
    fig_w = max(7.5, (max_cols + 1) * 1.05)
    fig_h = max(3.2, n_b * (rows * 0.72 + 1.0) + 0.8)
    fig, axes = plt.subplots(n_b, 1, figsize=(fig_w, fig_h), facecolor="white", squeeze=False)
    fig.suptitle(title, fontsize=20, fontweight="bold", color="#17365D", y=0.995)
    for r, (a,b) in enumerate(rangos):
        _dibujar_agujas_bloque(axes[r][0], df.iloc[:, a:b], font_size=max(font_size,16))
    fig.subplots_adjust(left=0.03, right=0.99, bottom=0.025, top=0.94, hspace=0.28)
    return fig


def _levas_doble_to_png(tablas):
    # Plato/Dial y Cilindro, con sus bloques, todos uno debajo del otro.
    paneles = []
    for titulo, df, orient in tablas:
        for bi, (a,b) in enumerate(_rangos_bloques_columnas(df.shape[1])):
            paneles.append((titulo if bi == 0 else "", df.iloc[:, a:b], orient))
    max_cols = max(df.shape[1] for _,df,_ in paneles)
    fig_w = max(8.5, (max_cols + 1) * 1.10)
    fig_h = max(6.0, sum(df.shape[0] * 0.78 + 1.15 for _,df,_ in paneles))
    fig, axes = plt.subplots(len(paneles), 1, figsize=(fig_w, fig_h), facecolor="white", squeeze=False)
    for r,(titulo,df,orient) in enumerate(paneles):
        ax = axes[r][0]
        _dibujar_leva_bloque(ax, df, orient)
        if titulo:
            ax.set_title(titulo, fontsize=16, fontweight="bold", color="#17365D", pad=6)
    fig.subplots_adjust(left=0.03,right=0.99,bottom=0.02,top=0.98,hspace=0.32)
    return fig


def _agujas_doble_to_png(tablas):
    paneles = []
    for titulo, df in tablas:
        for bi, (a,b) in enumerate(_rangos_bloques_columnas(df.shape[1])):
            paneles.append((titulo if bi == 0 else "", df.iloc[:, a:b]))
    max_cols = max(df.shape[1] for _,df in paneles)
    fig_w = max(8.5, (max_cols + 1) * 1.05)
    fig_h = max(6.0, sum(df.shape[0] * 0.72 + 1.10 for _,df in paneles))
    fig, axes = plt.subplots(len(paneles), 1, figsize=(fig_w, fig_h), facecolor="white", squeeze=False)
    for r,(titulo,df) in enumerate(paneles):
        ax = axes[r][0]
        _dibujar_agujas_bloque(ax, df, font_size=16)
        if titulo:
            ax.set_title(titulo, fontsize=16, fontweight="bold", color="#17365D", pad=6)
    fig.subplots_adjust(left=0.03,right=0.99,bottom=0.02,top=0.98,hspace=0.32)
    return fig


def _guardar_png_recortado(fig, buffer, dpi=220, margen_px=8, ancho_objetivo=1200):
    """Guarda el gráfico como una captura: recorta todo el blanco exterior y deja un margen mínimo."""
    tmp = BytesIO()
    fig.savefig(
        tmp, format="png", dpi=dpi,
        bbox_inches="tight", pad_inches=0,
        facecolor="white"
    )
    tmp.seek(0)
    img = Image.open(tmp).convert("RGB")

    # Detectar contenido: cualquier píxel que no sea casi blanco.
    fondo = Image.new("RGB", img.size, (255, 255, 255))
    diff = ImageChops.difference(img, fondo).convert("L")
    # Ignora ruido muy tenue del antialiasing, pero conserva líneas grises de la tabla.
    mask = diff.point(lambda px: 255 if px > 8 else 0)
    bbox = mask.getbbox()
    if bbox:
        l, t, r, b = bbox
        l = max(0, l - margen_px); t = max(0, t - margen_px)
        r = min(img.width, r + margen_px); b = min(img.height, b + margen_px)
        img = img.crop((l, t, r, b))

    # Normaliza el tamaño del archivo sin volver a agregar lienzo blanco.
    if ancho_objetivo and img.width != ancho_objetivo:
        nuevo_alto = max(1, round(img.height * ancho_objetivo / img.width))
        img = img.resize((ancho_objetivo, nuevo_alto), Image.Resampling.LANCZOS)

    img.save(buffer, format="PNG", optimize=True)
    buffer.seek(0)

def generar_zip_png():
    zip_buffer = BytesIO()

    # 1. Ligamento
    lig_buffer = BytesIO()
    fig_lig = (
        draw_minijacquard_ligament(minijacquard_pattern_sequences, minijacquard_pattern_assignments, visible_slots)
        if tipo_seleccion_agujas == "Minijacquard"
        else draw_system_ligaments(
            system_sequences, visible_slots, labels=system_labels, yarns=system_yarns,
            tipo_fontura=tipo_fontura
        )
    )
    _guardar_png_recortado(fig_lig, lig_buffer, dpi=220, margen_px=6, ancho_objetivo=1200)
    plt.close(fig_lig)

    # 2. Levas
    leva_buffer = BytesIO()

    if tipo_fontura == "Doblefontura":
        fig = _levas_doble_to_png([
            ("LEVAS - PLATO / DIAL", df_leva_plato_symbols, "down"),
            ("LEVAS - CILINDRO", df_leva_cil_symbols, "up"),
        ])
        _guardar_png_recortado(fig, leva_buffer, dpi=220, margen_px=6, ancho_objetivo=1200)
        plt.close(fig)
    else:
        fig = leva_dataframe_to_png(df_leva_symbols, "DISPOSICIÓN DE LEVAS", orientation="up")
        _guardar_png_recortado(fig, leva_buffer, dpi=220, margen_px=6, ancho_objetivo=1200)
        plt.close(fig)

    leva_buffer.seek(0)

    # 3. Agujas
    aguja_buffer = BytesIO()

    if tipo_fontura == "Doblefontura":
        fig_ag = _agujas_doble_to_png([
            ("AGUJAS / DIAL (PLATO)", df_needle_plato_symbols),
            ("AGUJAS / CILINDRO", df_needle_cil_symbols),
        ])
    else:
        titulo_ag, df_ag_export = needle_exports[0]
        if tipo_seleccion_agujas == "Minijacquard":
            fig_ag = minijacquard_dataframe_to_png(df_ag_export.fillna(""), titulo_ag)
        else:
            fig_ag = dataframe_to_png(df_ag_export.fillna(""), titulo_ag)

    _guardar_png_recortado(fig_ag, aguja_buffer, dpi=220, margen_px=6, ancho_objetivo=1200)
    plt.close(fig_ag)

    ficha = str(numero_ficha).strip() or "ficha"
    item = str(int(numero_item))
    base_archivo = f"{ficha}{item}"

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        # FICHA + ÍTEM + TIPO: 1=Ligamento, 2=Levas, 3=Agujas
        zf.writestr(f"{base_archivo}1.png", lig_buffer.getvalue())
        zf.writestr(f"{base_archivo}2.png", leva_buffer.getvalue())
        zf.writestr(f"{base_archivo}3.png", aguja_buffer.getvalue())

    zip_buffer.seek(0)
    return zip_buffer.getvalue()



def _trapecio_excel_png(orientation="up"):
    """Crea un pequeño PNG transparente del trapecio para insertarlo dentro de una celda Excel."""
    fig, ax = plt.subplots(figsize=(1.0, 0.45), dpi=120)
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    if orientation == "down":
        # Plato/Dial: ancho arriba, angosto abajo.
        pts = [(0.18, 0.78), (0.82, 0.78), (0.68, 0.22), (0.32, 0.22)]
    else:
        # Monofontura/Cilindro: angosto arriba, ancho abajo.
        pts = [(0.32, 0.78), (0.68, 0.78), (0.82, 0.22), (0.18, 0.22)]

    ax.add_patch(Polygon(pts, closed=True, fill=False, edgecolor="black", linewidth=2.2))

    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=120, transparent=True, bbox_inches="tight", pad_inches=0.02)
    plt.close(fig)
    buf.seek(0)
    return buf

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

    fmt_c_wrap = wb.add_format({
        "align":"center",
        "valign":"vcenter",
        "text_wrap":True,
        "border":1
    })

    # Ligamento como imagen
    img_lig = BytesIO()
    fig = (
        draw_minijacquard_ligament(minijacquard_pattern_sequences, minijacquard_pattern_assignments, visible_slots)
        if tipo_seleccion_agujas == "Minijacquard"
        else draw_system_ligaments(
            system_sequences, visible_slots, labels=system_labels, yarns=system_yarns,
            tipo_fontura=tipo_fontura
        )
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
        "GENERADOR AUTOMÁTICO DE LIGAMENTOS Y LEVAS – V5.12",
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
        ws.write("A9","Repeticiones selección agujas",fmt_h)
        ws.write("B9",int(needle_repetitions),fmt_c)
        sec_start = 10
    else:
        ws.write("A6","N° agujas",fmt_h)
        ws.write("B6",int(n_agujas),fmt_c)
        ws.write("A7","Repeticiones visibles ligamento",fmt_h)
        ws.write("B7",int(visible_slots),fmt_c)
        ws.write("A8","Repeticiones selección agujas",fmt_h)
        ws.write("B8",int(needle_repetitions),fmt_c)
        sec_start = 9

    # Secuencias
    ws.write(sec_start-1,0,"Sistema",fmt_h)
    ws.write(sec_start-1,1,"Secuencia",fmt_h)
    ws.write(sec_start-1,2,"Tipo / descripción",fmt_h)
    ws.write(sec_start-1,3,"Hilo / material",fmt_h)

    for i,seq in enumerate(system_sequences):
        ws.write(sec_start+i,0,i+1,fmt_c)
        ws.write(sec_start+i,1,"-".join(map(str,seq)),fmt_c)
        ws.write(sec_start+i,2,system_labels[i],fmt_c_wrap)
        ws.write(sec_start+i,3,system_yarns[i],fmt_c)
        # Aumentar la altura cuando la descripción tiene varias líneas.
        lineas_desc = max(1, str(system_labels[i]).count("\n") + 1)
        if lineas_desc > 1:
            ws.set_row(sec_start+i, 15 * lineas_desc)

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
            ws.write(r,j,str(col).replace("Sistema ", "S"),fmt_h)

        # Orientación visual de esta fontura en el Excel.
        orient_excel = "down" if "PLATO" in str(title).upper() or "DIAL" in str(title).upper() else "up"

        # El trapecio no se escribe como texto: se inserta como imagen dentro de la celda.
        # Así se ve igual que en las imágenes PNG del generador.
        trap_img = _trapecio_excel_png(orient_excel)

        for i in range(len(df_export)):
            fila_xls = r+i+1
            ws.set_row(fila_xls, 24)
            ws.write(fila_xls,0,df_export.index[i],fmt_h)

            for j in range(len(df_export.columns)):
                col_xls = j+1
                sval = str(df_export.iloc[i,j]).strip()

                if sval == "TRAP":
                    # Celda vacía + trapecio gráfico centrado.
                    ws.write_blank(fila_xls, col_xls, None, fmt_c)
                    trap_img.seek(0)
                    ws.insert_image(
                        fila_xls,
                        col_xls,
                        "trapecio.png",
                        {
                            "image_data": trap_img,
                            "x_offset": 26,
                            "y_offset": 4,
                            "x_scale": 0.52,
                            "y_scale": 0.52,
                            "object_position": 1
                        }
                    )
                elif sval.lower() == "nan" or sval == "":
                    ws.write_blank(fila_xls, col_xls, None, fmt_c)
                else:
                    # Malla (▲/▼) y Anulado (—) permanecen centrados como símbolos.
                    ws.write(fila_xls, col_xls, sval, fmt_c)

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
    file_name=f"{str(numero_ficha).strip() or 'ficha'}{int(numero_item)}-ligamento.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


st.download_button(
    "🗂️ EXPORTAR 3 PNG EN ZIP",
    data=generar_zip_png(),
    file_name=f"{str(numero_ficha).strip() or 'ficha'}{int(numero_item)}_imagenes.zip",
    mime="application/zip"
)

st.info(
    "V5.19: Minijacquard permite definir de 1 a 12 dibujos/patrones de ligamento y asignar libremente los sistemas de cada dibujo. Agujas Minijacquard se exportan en una sola tabla continua, sin bloques. En selección Normal se mantienen los bloques verticales. "
    "Se mantienen Plato/Dial descendente y hacia abajo, Cilindro ascendente y hacia arriba, "
    "tablas separadas y exportación Excel con trapecios gráficos."
)
