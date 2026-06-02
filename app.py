import io
import os
import glob
import pdfplumber
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# ----------------------------------------------------------------------
# Revisor de Instructivos de Embalaje — Garces Fruit
# App web: el usuario sube el PDF del instructivo y obtiene la revision
# contra el kit base maestro. Se ignoran las columnas Observacion y Pallet.
# ----------------------------------------------------------------------

st.set_page_config(page_title="Revisor de Instructivos", page_icon="🍒", layout="centered")

# ---------------- helpers ----------------
def norm(s):
    return "" if s is None else str(s).strip().upper()

def code(s):
    n = norm(s)
    return n.split()[0] if n else ""

def numv(s):
    d = "".join(ch for ch in str("" if s is None else s) if ch.isdigit())
    return int(d) if d else None

# ---------------- kit base ----------------
@st.cache_data(show_spinner=False)
def load_base(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    base = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        rows = list(ws.iter_rows(values_only=True))
        hi = env = emb = eti = caj = -1
        for i, r in enumerate(rows[:15]):
            u = [norm(c).replace(" ", "") for c in r]

            def find(keys):
                for j, c in enumerate(u):
                    if any(k in c for k in keys):
                        return j
                return -1

            e, b, t = find(["ENVASE"]), find(["EMBALAJE"]), find(["ETIQUETA"])
            if e >= 0 and b >= 0 and t >= 0:
                hi, env, emb, eti = i, e, b, t
                caj = find(["CANTIDADDECAJAS", "CAJASPALLET", "CAJAS", "CANTIDAD"])
                break
        if hi < 0:
            continue
        for r in rows[hi + 1:]:
            ev = code(r[env]) if env < len(r) else ""
            if not ev:
                continue
            base.append({
                "envase": ev,
                "embalaje": code(r[emb]) if emb < len(r) else "",
                "etiqueta": norm(r[eti]) if eti < len(r) else "",
                "cajas": numv(r[caj]) if 0 <= caj < len(r) else None,
            })
    return base

# ---------------- instructivo PDF ----------------
def extract_kits(pdf_bytes):
    kits = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            for t in page.extract_tables():
                if len(t) < 2:
                    continue
                hdr = [norm(c).replace(" ", "") for c in t[1]]
                if "ENVASE" in hdr and "EMBALAJE" in hdr and "ETIQUETA" in hdr:
                    iE, iB, iT = hdr.index("ENVASE"), hdr.index("EMBALAJE"), hdr.index("ETIQUETA")
                    iCaj = hdr.index("CAJAS/PALLET") if "CAJAS/PALLET" in hdr else -1
                    iCal = hdr.index("CALIBRES") if "CALIBRES" in hdr else (
                        hdr.index("CALIBRE") if "CALIBRE" in hdr else -1)
                    for r in t[2:]:
                        ev = code(r[iE]) if iE < len(r) else ""
                        if not ev:
                            continue
                        kits.append({
                            "envase": ev,
                            "embalaje": code(r[iB]) if iB < len(r) else "",
                            "etiqueta": norm(r[iT]) if iT < len(r) else "",
                            "calibre": norm(r[iCal]) if 0 <= iCal < len(r) else "",
                            "cajas": numv(r[iCaj]) if 0 <= iCaj < len(r) else None,
                        })
    return kits

# ---------------- comparacion ----------------
def review(base, kits):
    out = []
    for k in kits:
        m = [b for b in base if b["envase"] == k["envase"]
             and b["embalaje"] == k["embalaje"] and b["etiqueta"] == k["etiqueta"]]
        if not m:
            partial = [b for b in base if b["envase"] == k["envase"] and b["embalaje"] == k["embalaje"]]
            if partial:
                sug = ", ".join(sorted(set(b["etiqueta"] for b in partial))[:3])
                st_, note = "ERROR", f'Etiqueta "{k["etiqueta"] or "(vacía)"}" no existe para este envase/embalaje. La base registra: {sug}.'
            else:
                st_, note = "ERROR", "La combinación Envase + Embalaje no existe en el kit base."
        else:
            cb = sorted(set(b["cajas"] for b in m if b["cajas"] is not None))
            if k["cajas"] is None or not cb or k["cajas"] in cb:
                st_, note = "OK", ""
            else:
                st_, note = "ADVERTENCIA", f'Cajas/pallet = {k["cajas"]}; la base registra {" / ".join(map(str, cb))} para este kit.'
        out.append({**k, "estado": st_, "detalle": note})
    return out

# ---------------- UI ----------------
st.markdown("""
<style>
.block-container{max-width:820px}
.res{border:1px solid #e3e2d8;border-radius:12px;padding:13px 16px;margin-bottom:9px}
.res.OK{border-left:5px solid #97c459}
.res.ERROR{border-left:5px solid #f09595;background:#fefafa}
.res.ADVERTENCIA{border-left:5px solid #ef9f27;background:#fffdf8}
.rtop{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px}
.rid{font-family:monospace;font-size:13px;font-weight:600}
.tag{font-size:11px;font-weight:700;padding:3px 11px;border-radius:20px}
.tag.OK{background:#eaf3de;color:#3b6d11}
.tag.ERROR{background:#fcebeb;color:#a32d2d}
.tag.ADVERTENCIA{background:#faeeda;color:#854f0b}
.rfields{font-family:monospace;font-size:12.5px;color:#444}
.note{font-size:12.5px;margin-top:7px;color:#a32d2d}
.note.ADVERTENCIA{color:#854f0b}
</style>
""", unsafe_allow_html=True)

st.title("🍒 Revisor de Instructivos")
st.caption("Compara los kits del instructivo (PDF) contra el kit base maestro. "
           "Se ignoran las columnas **Observación** y **Pallet**.")

# --- cargar kit base (bundled o subido) ---
base = None
bundled = sorted(glob.glob(os.path.join(os.path.dirname(__file__), "KIT*.xls*")))
with st.expander("⚙️ Kit base maestro", expanded=not bundled):
    if bundled:
        st.success(f"Usando kit base incluido: **{os.path.basename(bundled[0])}**")
        with open(bundled[0], "rb") as f:
            base = load_base(f.read())
    up_base = st.file_uploader("Reemplazar kit base (opcional)", type=["xlsx", "xlsm", "xls"], key="base")
    if up_base is not None:
        base = load_base(up_base.read())
        st.success("Kit base actualizado.")
if base is not None:
    st.caption(f"Kit base cargado: **{len(base)} kits** de referencia.")

# --- subir instructivo ---
st.subheader("Sube el instructivo (PDF)")
up_pdf = st.file_uploader("Instructivo de embalaje", type=["pdf"], key="pdf",
                          label_visibility="collapsed")

if up_pdf is not None and base:
    kits = extract_kits(up_pdf.read())
    if not kits:
        st.error("No pude leer la tabla de órdenes específicas en este PDF. "
                 "¿Es un instructivo en formato GF-IND-PL-003?")
    else:
        res = review(base, kits)
        n_ok = sum(r["estado"] == "OK" for r in res)
        n_er = sum(r["estado"] == "ERROR" for r in res)
        n_wa = sum(r["estado"] == "ADVERTENCIA" for r in res)

        c1, c2, c3 = st.columns(3)
        c1.metric("OK", n_ok)
        c2.metric("Errores", n_er)
        c3.metric("Advertencias", n_wa)

        st.divider()
        for r in res:
            fields = f'{r["envase"]} · {r["embalaje"]} · {r["etiqueta"] or "—"} · cajas/pallet: {r["cajas"] if r["cajas"] is not None else "—"}'
            note = f'<div class="note {r["estado"]}">{r["detalle"]}</div>' if r["detalle"] else ""
            st.markdown(
                f'<div class="res {r["estado"]}"><div class="rtop">'
                f'<span class="rid">{r["envase"]} · {r["embalaje"]} · {r["etiqueta"] or "—"}</span>'
                f'<span class="tag {r["estado"]}">{r["estado"]}</span></div>'
                f'<div class="rfields">{fields}</div>{note}</div>',
                unsafe_allow_html=True)

        # exportar a Excel
        df = pd.DataFrame([{
            "Envase": r["envase"], "Embalaje": r["embalaje"], "Etiqueta": r["etiqueta"],
            "Calibre": r["calibre"], "Cajas/pallet": r["cajas"],
            "Estado": r["estado"], "Detalle": r["detalle"],
        } for r in res])
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name="Revisión")
        st.download_button("⬇ Descargar resultado en Excel", buf.getvalue(),
                           file_name="revision_instructivo.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif up_pdf is not None and not base:
    st.warning("Primero carga el kit base maestro (arriba).")
