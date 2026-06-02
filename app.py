import io
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

# ---------------- kit base incrustado ----------------
# Datos del kit base maestro (envase|embalaje|etiqueta|cajas_por_pallet).
# Para actualizarlo: reemplaza este bloque o sube un Excel nuevo en la app.
KIT_BASE = """\
C0535102|CEHFTE|SAN FRANCISCO|112
C0535102|CEHFTT|SAN FRANCISCO|112
C0535102|CEHFTG|SAN FRANCISCO|112
C0535102|CEHFTH|SAN FRANCISCO|112
C0535102|CEHFTE|GARCES|112
C0535102|CEHFTG|GARCES|112
C0535102|CEHFTT|GARCES|112
C0535102|CESFTG|GARCES|176
C0535102|CESFTG|SAN FRANCISCO|176
C0535102|CEAFTG|GARCES|176
C0535102|CEAFTE|GARCES|176
C0535102|CEAFTK|GARCES|176
C0535102|CEAFTI|GARCES|176
C0535102|CEAFTT|GARCES|176
C0535102|CEAFTH|GARCES|176
C0535102|CEAFTG|SAN FRANCISCO|176
C0535102|CEAFTE|SAN FRANCISCO|176
C0535102|CEAFTK|SAN FRANCISCO|176
C0535102|CEAFTI|SAN FRANCISCO|176
C0535102|CEAFTT|SAN FRANCISCO|176
C0535102|CEAFTH|SAN FRANCISCO|176
C0535102|CEFPPG|GARCES|176
C0535103|CEAS10|SAN FRANCISCO|152
C0535103|CEAS10|SAN FRANCISCO|176
C0535103|CEPS10|SAN FRANCISCO|112
C0546070|CEAB16|SAN FRANCISCO|145
C0546070|CEPB16|SAN FRANCISCO|100
C0535103|CEP10E|SAN FRANCISCO|112
C0535103|CESS10|SAN FRANCISCO|152
C4434125|CEAMGR|YELLOW DREAMS|180
C4635096|CEAMGR|YELLOW DREAMS|184
C4434125|CEAMGR|GARCES RAINIER|180
C4434125|CEAMGE|GARCES RAINIER|180
C4434125|CEHHGR|GARCES RAINIER|110
C4434125|CEHHGR|YELLOW DREAMS|110
C4434125|CEHHGE|GARCES RAINIER|110
C4434125|CEHHGT|GARCES RAINIER|110
C4635096|CEHHGE|YELLOW DREAMS|120
C4635096|CEHHGR|YELLOW DREAMS|120
C4434125|CEAMGT|GARCES RAINIER|180
C4434125|CEAMGH|GARCES RAINIER|180
C7246123|CEAT8C|GARCES|80
C7246123|CEHT8C|GARCES|55
C7246123|CEMI8T|GARCES|55
C7246123|CEMI8J|GARCES|45
C0846131|CEP20C|GARCES|55
C0846131|CEA20C|GARCES|75
C0846131|CEP20C|S/E|55
C0846131|CEA20C|S/E|75
C2532098|CEASGR|GARCES|368
C2532098|CEAMGR|GARCES|368
C0535096|CEAMGK|SAN FRANCISCO PREMIUM|184
C0535096|CEAMGK|SAN FRANCISCO|184
C0535096|CEAMGR|SAN FRANCISCO|184
C0535096|CEAMGR|LUCKY|184
C0535096|CEAMGR|RED DREAMS|184
C0535096|CEAMGE|SAN FRANCISCO|184
C0535096|CEAMGL|SAN FRANCISCO|184
C0535096|CEHHGK|SAN FRANCISCO PREMIUM|120
C0535096|CEHHGR|SAN FRANCISCO|120
C0535096|CEHHGR|RED DREAMS|120
C0535096|CEHHGK|SAN FRANCISCO|120
C0535096|CEAMGH|SAN FRANCISCO|184
C0535096|CEPHGR|SAN FRANCISCO|120
C0535096|CEPHGR|RED DREAMS|120
C0535096|CEASGR|LUCKY|184
C0535096|CEASGR|SAN FRANCISCO|184
C0535096|CEPHGR|LUCKY|120
C0535096|CEAMGX|SAN FRANCISCO|184
C0535096|CEASGK|SAN FRANCISCO|184
C0535096|CEASGR|RED DREAMS|184
C0535096|CEASGK|SAN FRANCISCO PREMIUM|184
C0535102|CEHHGT|SAN FRANCISCO|112
C0535102|CEHHGR|SAN FRANCISCO|112
C0535102|CEHHGH|SAN FRANCISCO|112
C0535102|CEHHGE|GARCES|112
C0535102|CEHHGR|GARCES|112
C0535102|CEHHGT|GARCES|112
C0535102|CEAMGR|GARCES|176
C0535102|CEAMGE|GARCES|176
C0535102|CEAMGK|GARCES|176
C0535102|CEAMGI|GARCES|176
C0535102|CEAMGT|GARCES|176
C0535102|CEAMGH|GARCES|176
C0535102|CEASGR|GARCES|176
C0535102|CEAMGR|SAN FRANCISCO|176
C0535102|CEAMGE|SAN FRANCISCO|176
C0535102|CEAMGK|SAN FRANCISCO|176
C0535102|CEAMGI|SAN FRANCISCO|176
C0535102|CEAMGT|SAN FRANCISCO|176
C0535102|CEAMGH|SAN FRANCISCO|176
C0535102|CEHHGK|GARCES|112
C0535102|CEAMGX|GARCES|176
C0535102|CEHHGX|GARCES|112
C0535102|CEASGK|GARCES|176
C0635120|CEAM3K|GARCES|152
C0635120|CEHHG3|GARCES|96
C0635120|CEASG3|GARCES|152
C0635120|CEAS3K|GARCES|152
P0232087|CEAMGR|GARCES|416
P0232087|CEHHGR|GARCES|256
P2532103|CEAMGR|GARCES|352
P2532103|CEHHGR|GARCES|224
P2532103|CEHHGE|GARCES|224
P0232087|CEASGR|GARCES|416
P2532103|CEASGR|GARCES|352
P2532103|CEAMGX|GARCES|352
P2532103|CEHHGX|GARCES|224
C8535133|CEAMGK|GARCES|136
C8535133|CEHHGE|GARCES|80
P1035140|CEAMGR|S/E|128
P1035140|CEAMGL|S/E|112
P1035150|CEAMGR|S/E|120
P1035140|CEAMGR|GARCES|128
C8535133|CEAMGH|GARCES|136
C8535133|CEHHGH|GARCES|80
P1035140|CEAMGH|GARCES|128
P1035150|CEAMGL|GARCES|128
P1035150|CEAMGL|S/E|112
C8535133|CEAMGR|GARCES|136
C9546123|CEAMGR|GARCES|80
P1035140|CEAMGE|S/E|128
P1035140|NAGRR|S/E|88
NACP34|NAGR|S/E|50
P1035140|NAGR|S/E|88
NACP34|NAGR|CONGELADO|45
C0535100|NAGR|SAN FRANCISCO PREMIUM|152
C0534125|NAGR|SAN FRANCISCO RAINIER|150
C0534112|NAGR|S/E|170
C0846135|NAGR|GARCES|70
C0646135|NAGR|S/E|50
C1534275|CISBEU|GARCES|80
C1534275|CISBEM|GARCES|80
C1534275|CISBEUM|GARCES|80
C1546178|NAGR|S/E|40
C1246156|NAGR|S/E|45
NACP34|NAGR|S/E|50
NABPV|NAGR|S/E|1
"""

def load_embedded_base():
    base = []
    for line in KIT_BASE.strip().splitlines():
        parts = line.split("|")
        if len(parts) < 4:
            continue
        ev, emb, eti, caj = parts[0], parts[1], parts[2], parts[3]
        if not ev:
            continue
        base.append({
            "envase": ev.strip().upper(),
            "embalaje": emb.strip().upper(),
            "etiqueta": eti.strip().upper(),
            "cajas": int(caj) if caj.strip().isdigit() else None,
        })
    return base

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
#MainMenu, footer, header, [data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"]{visibility:hidden; height:0;}
.block-container{max-width:840px; padding-top:1.4rem;}
.gf-header{background:#2f5d2a; border-radius:14px; padding:18px 22px; margin-bottom:6px; display:flex; align-items:center; gap:14px;}
.gf-logo{width:44px; height:44px; border-radius:11px; background:#ffffff; color:#2f5d2a; display:flex; align-items:center; justify-content:center; font-weight:700; font-size:19px; flex-shrink:0; font-family:sans-serif;}
.gf-header h1{margin:0; color:#ffffff; font-size:22px; font-weight:600; line-height:1.15;}
.gf-header p{margin:2px 0 0; color:#cfe0c6; font-size:12.5px;}
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

st.markdown(
    '<div class="gf-header"><div class="gf-logo">GF</div>'
    '<div><h1>Revisor de Instructivos</h1>'
    '<p>Garces Fruit · revisión de kits contra el kit base maestro</p></div></div>',
    unsafe_allow_html=True)
st.caption("Sube el instructivo en PDF y obtén la revisión. "
           "Se ignoran las columnas **Observación** y **Pallet**.")

# --- kit base (incrustado, con opción de reemplazar) ---
base = load_embedded_base()
with st.expander("⚙️ Kit base maestro", expanded=False):
    st.success(f"Kit base incluido en la app: **{len(base)} kits** de referencia.")
    up_base = st.file_uploader("Reemplazar kit base con un Excel (opcional)", type=["xlsx", "xlsm", "xls"], key="base")
    if up_base is not None:
        base = load_base(up_base.read())
        st.success(f"Kit base actualizado desde Excel: {len(base)} kits.")
if base is not None:
    st.caption(f"Comparando contra **{len(base)} kits** del kit base.")

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
