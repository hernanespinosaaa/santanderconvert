import streamlit as st
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import io
from datetime import datetime

# ── Paleta y Estilos ────────────────────────────────────────────────────────
C_HEADER_BG = "CC0000"
C_HEADER_FG = "FFFFFF"
C_SUBHEADER = "F2F2F2"
C_DEBITO    = "FFF0F0"
C_CREDITO   = "F0FFF0"
C_IMPUESTO  = "FFF8E1"
C_TOTAL_BG  = "EEEEEE"
C_WARN      = "FFF3CD"

THIN = Side(style="thin", color="CCCCCC")
BRD  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def hdr(cell, bg=C_HEADER_BG, fg=C_HEADER_FG, bold=True, sz=10):
    cell.font      = Font(name="Arial", bold=bold, color=fg, size=sz)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = BRD

# ── Helpers ─────────────────────────────────────────────────────────────────
def categorizar(d):
    d = d.lower()
    if "saldo inicial" in d:                           return "Saldo Inicial"
    if "haberes" in d or "sueldo" in d:                return "Sueldos y haberes"
    if "honorario" in d or "/ hon" in d:               return "Honorarios"
    if "alquiler" in d or "/ alq" in d:                return "Alquileres"
    if "afip" in d or "imp.afp" in d:                  return "Impuestos AFIP"
    if "ley 25.413" in d or "debito 0,6" in d:         return "Imp. débitos/créditos"
    if "ii bb" in d or "iibb" in d:                    return "IIBB"
    if "iva" in d:                                     return "IVA"
    if "fondos comunes" in d or "rescate" in d:        return "FCI"
    if "suscripcion" in d or "suscripción" in d:       return "FCI"
    if "debin" in d:                                   return "DEBIN"
    if "seguro" in d or "/ seg" in d:                  return "Seguros"
    if "comision" in d:                                return "Comisiones bancarias"
    if "prev.social" in d or "/ cuo" in d:             return "Previsión social"
    if "transferencia" in d or "transf" in d:          return "Transferencias"
    if "acreditacion" in d or "acreditación" in d:     return "Acreditaciones"
    if "cheque" in d or "echeq" in d:                  return "Cheques"
    return "Otros"

# ── Lógica de Extracción ────────────────────────────────────────────────────
RE_MONTO_STR = re.compile(r"-?\(?(?:U\$S|US\$|\$)?\s*[\d.]+,\d{2}\)?")
RE_FECHA     = re.compile(r"^\d{2}/\d{2}/\d{2}$")
RE_COMP      = re.compile(r"^\d{5,9}$")

def parse_monto(s):
    if s is None: return None
    s = re.sub(r"[^\d.,\-()]", "", str(s)).strip()
    if not s: return None
    negativo = s.startswith("-") or (s.startswith("(") and s.endswith(")"))
    s = s.lstrip("-(").rstrip(")")
    if re.search(r",\d{2}$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        v = float(s)
        return -v if negativo else v
    except ValueError:
        return None

def _parsear_importes(montos_str, desc, saldo_anterior=None):
    vals = [(parse_monto(m), m) for m in montos_str]
    vals = [(v, raw) for v, raw in vals if v is not None]
    debito = credito = saldo = None
    desc_l = desc.lower()

    if len(vals) >= 2:
        importe, raw_importe = vals[-2]
        saldo, _             = vals[-1]
        importe_abs = abs(importe)

        math_success = False
        if saldo_anterior is not None and saldo is not None:
            dif = round(saldo - saldo_anterior, 2)
            if dif > 0 and abs(dif) == round(importe_abs, 2):
                credito = importe_abs
                math_success = True
            elif dif < 0 and abs(dif) == round(importe_abs, 2):
                debito = importe_abs
                math_success = True

        if not math_success:
            if "rescate" in desc_l: credito = importe_abs
            elif "suscripcion" in desc_l or "suscripción" in desc_l: debito = importe_abs
            else:
                raw_clean = raw_importe.strip()
                es_negativo = (importe < 0 or raw_clean.startswith("-") or (raw_clean.startswith("(") and raw_clean.endswith(")")))
                if es_negativo: debito = importe_abs
                else: credito = importe_abs
    elif len(vals) == 1:
        saldo = vals[0][0]
    return debito, credito, saldo

def procesar_lineas_movimientos(lineas):
    movimientos = []
    saldo_actual = None
    idx_start = 0
    fecha_corriente = ""

    # 1. Cacería intensiva del "Saldo Inicial" en las primeras líneas
    idx_saldo = -1
    for i, l in enumerate(lineas):
        if "saldo inicial" in l.lower() or "saldo en cuenta" in l.lower():
            idx_saldo = i
            break
            
    if idx_saldo != -1:
        fecha_inicial = ""
        # Buscar la fecha un renglón antes o después
        for l in lineas[max(0, idx_saldo-2) : idx_saldo+3]:
            m_fecha = RE_FECHA.search(l)
            if m_fecha:
                fecha_inicial = m_fecha.group(0)
                break
                
        # Buscar el monto de plata en el mismo renglón o los siguientes
        for i in range(idx_saldo, min(len(lineas), idx_saldo+4)):
            montos = RE_MONTO_STR.findall(lineas[i])
            if montos:
                saldo_actual = parse_monto(montos[-1])
                idx_start = i + 1  # Empezar a leer movimientos DESPUÉS de esto
                break
                
        # Si encontramos cuánta plata había, armamos el movimiento manual
        if saldo_actual is not None:
            movimientos.append({
                "fecha": fecha_inicial,
                "comprobante": "",
                "descripcion": "Saldo Inicial",
                "debito": None,
                "credito": None,
                "saldo": saldo_actual,
                "categoria": "Saldo Inicial",
                "sin_fecha": not bool(fecha_inicial)
            })
            fecha_corriente = fecha_inicial

    # 2. Procesar el resto de los movimientos de forma normal
    i = idx_start
    while i < len(lineas):
        l = lineas[i].strip()
        if not l:
            i += 1
            continue

        # Esquivar basura de la tabla y encabezados repetidos
        if (re.match(r"^\d+\s*-\s*\d+$", l) or
            ("Fecha" in l and "Comprobante" in l) or
            ("Cuenta Corriente" in l and "CBU" in l) or
            l.startswith("* Salvo") or
            l.lower().startswith("total") or
            "No tenés movimientos" in l or
            "saldo inicial" in l.lower() or
            "saldo total" in l.lower()):
            i += 1
            continue

        if RE_FECHA.match(l):
            fecha_corriente = l
            i += 1
            continue

        montos_encontrados = RE_MONTO_STR.findall(l)
        if not montos_encontrados:
            if movimientos and movimientos[-1]["categoria"] != "Saldo Inicial" and not RE_FECHA.match(l):
                movimientos[-1]["descripcion"] += " | " + l
                movimientos[-1]["categoria"] = categorizar(movimientos[-1]["descripcion"])
            i += 1
            continue

        sin_montos = RE_MONTO_STR.sub("", l).strip()
        tokens     = sin_montos.split()
        fecha = comp = ""
        desc_t = []

        for t in tokens:
            if not fecha and RE_FECHA.match(t): fecha = t
            elif not comp and RE_COMP.match(t): comp = t
            else: desc_t.append(t)

        if not fecha: fecha = fecha_corriente
        desc = " ".join(desc_t).strip()

        if i + 1 < len(lineas):
            sig = lineas[i + 1].strip()
            if (sig and not RE_MONTO_STR.findall(sig) and
                not RE_FECHA.match(sig) and
                not sig.lower().startswith("total") and
                not "saldo total" in sig.lower() and
                not re.match(r"^\d+\s*-\s*\d+$", sig)):
                desc = (desc + " | " + sig) if desc else sig
                i += 1

        debito, credito, saldo = _parsear_importes(montos_encontrados, desc, saldo_actual)
        if saldo is not None: saldo_actual = saldo

        movimientos.append({
            "fecha": fecha, "comprobante": comp, "descripcion": desc,
            "debito": debito, "credito": credito, "saldo": saldo,
            "categoria": categorizar(desc), "sin_fecha": fecha == "",
        })
        i += 1
    return movimientos

def extraer_pdf_multicuenta(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        todas = []
        for page in pdf.pages:
            todas.extend((page.extract_text() or "").splitlines())

    info_global = {"cuit": "", "desde": "", "hasta": "", "razon_social": ""}
    cuentas = {}
    cuenta_actual = None

    for l in todas:
        l_strip = l.strip()

        if not info_global["cuit"]:
            m = re.search(r"CUIT[:\s]+([\d\-]+)", l_strip)
            if m: info_global["cuit"] = m.group(1)
        if not info_global["desde"]:
            m = re.search(r"Desde:\s*(\d{2}/\d{2}/\d{2})", l_strip)
            if m: info_global["desde"] = m.group(1)
        if not info_global["hasta"]:
            m = re.search(r"Hasta:\s*(\d{2}/\d{2}/\d{2})", l_strip)
            if m: info_global["hasta"] = m.group(1)
        if not info_global["razon_social"] and info_global["cuit"]:
            if (re.match(r"^[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s\.\-&,]+$", l_strip)
                    and 4 < len(l_strip) < 60
                    and l_strip not in ("EXTRACTO DE CUENTA", "CUENTA CORRIENTE", "BANCO SANTANDER", "RESUMEN DE CUENTA", "Detalle impositivo")):
                info_global["razon_social"] = l_strip

        if "Detalle impositivo" in l_strip or "Legales" in l_strip:
            cuenta_actual = None
            continue

        m_cta = re.search(r"N[°ºoO]?\s*([\d\-/]+)\s+CBU:\s*(\d+)", l_strip)
        if m_cta:
            cta = m_cta.group(1)
            cbu = m_cta.group(2)
            cuenta_actual = cta
            if cta not in cuentas:
                cuentas[cta] = {"nro_cuenta": cta, "cbu": cbu, "lineas": []}
            continue

        if cuenta_actual:
            cuentas[cuenta_actual]["lineas"].append(l_strip)

    resultados = []
    for cta, datos in cuentas.items():
        movimientos = procesar_lineas_movimientos(datos["lineas"])
        
        info_completa = {
            **info_global,
            "nro_cuenta": datos["nro_cuenta"],
            "cbu": datos["cbu"]
        }
        
        # Ignorar cuentas que solo trajeron la info del Saldo Inicial pero ningún movimiento real
        movs_reales = [m for m in movimientos if "saldo inicial" not in m["descripcion"].lower()]
        if movs_reales:
            resultados.append({"info": info_completa, "movimientos": movimientos})

    return resultados

# ── Creación del Excel en Memoria ───────────────────────────────────────────
def crear_excel_buffer(info, movimientos):
    wb = Workbook()
    ws = wb.active
    ws.title = "Movimientos"
    ws.freeze_panes = "A5"

    ws.merge_cells("A1:H1")
    ws["A1"].value = f"Resumen de Cuenta — {info.get('razon_social', '')}  |  {info.get('desde', '')} al {info.get('hasta', '')}"
    hdr(ws["A1"], sz=12)
    ws.row_dimensions[1].height = 22

    ws.merge_cells("A2:H2")
    s = ws["A2"]
    s.value = f"Cta N° {info.get('nro_cuenta', '')}  |  CBU: {info.get('cbu', '')}  |  CUIT: {info.get('cuit', '')}"
    s.font, s.fill, s.alignment, s.border = Font(name="Arial", size=9, color="555555"), PatternFill("solid", fgColor=C_SUBHEADER), Alignment(horizontal="center"), BRD
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 6

    for c, h in enumerate(["Fecha", "Comprobante", "Descripción", "Categoría", "Débito", "Crédito", "Saldo", ""], 1):
        cell = ws.cell(row=4, column=c, value=h)
        hdr(cell, bg="8B0000" if c in (5, 6, 7) else C_HEADER_BG)
    for i, w in enumerate([10, 12, 48, 22, 14, 14, 14, 2], 1): ws.column_dimensions[get_column_letter(i)].width = w

    fila = 5
    for mov in movimientos:
        bg = "FFFFFF"
        if mov.get("sin_fecha"): bg = C_WARN
        elif mov["categoria"] == "Saldo Inicial": bg = C_TOTAL_BG
        elif mov["debito"] and not mov["credito"]: bg = C_DEBITO
        elif mov["credito"] and not mov["debito"]: bg = C_CREDITO
        
        if mov["categoria"] in ("Imp. débitos/créditos", "IVA", "IIBB"): bg = C_IMPUESTO

        for c, v in enumerate([mov["fecha"], mov["comprobante"], mov["descripcion"], mov["categoria"], mov["debito"], mov["credito"], mov["saldo"]], 1):
            cell = ws.cell(row=fila, column=c, value=v)
            cell.font, cell.fill, cell.border = Font(name="Arial", size=9), PatternFill("solid", fgColor=bg), BRD
            if c in (5, 6, 7) and v is not None:
                cell.number_format, cell.alignment = '#,##0.00;[Red](#,##0.00);-', Alignment(horizontal="right")
        fila += 1

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ── INTERFAZ WEB STREAMLIT ──────────────────────────────────────────────────
st.set_page_config(page_title="Santander a Excel", page_icon="🏦", layout="centered")

# Inicializar la "memoria" de la página
if "archivos_procesados" not in st.session_state:
    st.session_state.archivos_procesados = {}

st.title("🏦 Conversor Santander a Excel")
st.write("Subí tus extractos en PDF. Te generaremos un Excel por cada cuenta que tenga movimientos.")

uploaded_files = st.file_uploader("Arrastrá los PDF acá", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Procesar Archivos", type="primary"):
        st.session_state.archivos_procesados = {} # Limpiamos la memoria vieja
        
        for file in uploaded_files:
            with st.spinner(f"Escaneando {file.name}..."):
                try:
                    resultados = extraer_pdf_multicuenta(file)
                    if resultados:
                        st.success(f"✅ {file.name} escaneado: ¡Se detectaron {len(resultados)} cuenta(s) con movimientos!")
                        
                        buffers_descarga = []
                        for res in resultados:
                            info = res["info"]
                            movs = res["movimientos"]
                            excel_buffer = crear_excel_buffer(info, movs)
                            cuenta_limpia = info['nro_cuenta'].replace('/', '-')
                            
                            buffers_descarga.append({
                                "nombre_archivo": f"Cta_{cuenta_limpia}_{file.name.replace('.pdf', '')}.xlsx",
                                "buffer": excel_buffer,
                                "titulo_boton": f"📥 Descargar Cuenta {info['nro_cuenta']} ({len(movs)-1} movs)"
                            })
                        
                        # Guardamos los botones en la memoria de la página
                        st.session_state.archivos_procesados[file.name] = buffers_descarga
                    else:
                        st.warning(f"⚠️ No se detectaron movimientos en {file.name}.")
                except Exception as e:
                    st.error(f"❌ Ocurrió un error al procesar {file.name}: {str(e)}")

    # Mostramos los botones de descarga usando la información de la memoria
    # ¡Así no se borran cuando hacés clic en uno!
    for file in uploaded_files:
        if file.name in st.session_state.archivos_procesados:
            st.markdown(f"**Descargas listas para: {file.name}**")
            for i, data_boton in enumerate(st.session_state.archivos_procesados[file.name]):
                st.download_button(
                    label=data_boton["titulo_boton"],
                    data=data_boton["buffer"],
                    file_name=data_boton["nombre_archivo"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{file.name}_{i}"
                )
else:
    # Si borrás los archivos del cuadro de subida, limpiamos la memoria
    st.session_state.archivos_procesados = {}
