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
def parse_monto(s):
    if s is None: return None
    s = re.sub(r"[\$\s]", "", str(s)).strip()
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

def categorizar(d):
    d = d.lower()
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
RE_MONTO_STR = re.compile(r"-?\(?\$?\s*[\d.]+,\d{2}\)?")
RE_FECHA     = re.compile(r"^\d{2}/\d{2}/\d{2}$")
RE_COMP      = re.compile(r"^\d{5,9}$")

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

def procesar_lineas_movimientos(lineas, saldo_inicial):
    """Procesa un bloque de líneas que ya se sabe que son movimientos puros."""
    movimientos = []
    fecha_corriente = ""
    i = 0
    saldo_actual = saldo_inicial

    while i < len(lineas):
        l = lineas[i]
        if RE_FECHA.match(l):
            fecha_corriente = l
            i += 1
            continue

        montos_encontrados = RE_MONTO_STR.findall(l)
        if not montos_encontrados:
            if movimientos and l and not RE_FECHA.match(l):
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
            if sig and not RE_MONTO_STR.findall(sig) and not RE_FECHA.match(sig):
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
    last_nro_cuenta = ""
    last_cbu = ""
    last_saldo_inicial = None
    
    bloques = []
    bloque_actual = []
    in_movimientos = False
    
    for l in todas:
        l_strip = l.strip()
        
        # 1. Info General
        if not info_global["cuit"]:
            m = re.search(r"CUIT[:\s]+([\d\-]+)", l)
            if m: info_global["cuit"] = m.group(1)
        if not info_global["desde"]:
            m = re.search(r"Desde:\s*(\d{2}/\d{2}/\d{2})", l)
            if m: info_global["desde"] = m.group(1)
        if not info_global["hasta"]:
            m = re.search(r"Hasta:\s*(\d{2}/\d{2}/\d{2})", l)
            if m: info_global["hasta"] = m.group(1)
        if not info_global["razon_social"] and info_global["cuit"]:
            if (re.match(r"^[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s\.\-&,]+$", l_strip)
                    and 4 < len(l_strip) < 60
                    and l_strip not in ("EXTRACTO DE CUENTA", "CUENTA CORRIENTE", "BANCO SANTANDER", "RESUMEN DE CUENTA")):
                info_global["razon_social"] = l_strip

        # 2. Info de la cuenta (se actualiza para guardar la última leída antes de los movimientos)
        m_cta = re.search(r"N[°º]\s*([\d\-/]+)\s+CBU:\s*(\d+)", l)
        if m_cta:
            last_nro_cuenta = m_cta.group(1)
            last_cbu = m_cta.group(2)
            
        m_saldo = re.search(r"Saldo Inicial\s+\$\s*([\d.,]+)", l)
        if m_saldo:
            last_saldo_inicial = parse_monto(m_saldo.group(1))

        # 3. Lógica de atrapado de bloques (Motor principal)
        if "Movimientos en pesos" in l:
            if not in_movimientos:
                in_movimientos = True
                bloque_actual = []
            continue
            
        if in_movimientos:
            if "Saldo total" in l:
                in_movimientos = False
                bloques.append({
                    "nro_cuenta": last_nro_cuenta,
                    "cbu": last_cbu,
                    "saldo_inicial": last_saldo_inicial,
                    "lineas": bloque_actual
                })
                last_saldo_inicial = None # Reseteamos por si la próxima cuenta no lo tiene
            else:
                s = l.strip()
                # Filtramos basura de la paginación y encabezados repetidos
                if s and not re.match(r"^\d+\s*-\s*\d+$", s) and not ("Fecha" in s and "Comprobante" in s) and not ("Cuenta Corriente" in s and "CBU" in s) and not s.startswith("* Salvo"):
                    bloque_actual.append(s)

    # Por si el PDF termina abruptamente sin la frase "Saldo total"
    if in_movimientos and bloque_actual:
        bloques.append({
            "nro_cuenta": last_nro_cuenta,
            "cbu": last_cbu,
            "saldo_inicial": last_saldo_inicial,
            "lineas": bloque_actual
        })

    # Procesar cada bloque atrapado
    resultados = []
    for i, blk in enumerate(bloques):
        movimientos = procesar_lineas_movimientos(blk["lineas"], blk["saldo_inicial"])
        
        # Si no encontró N° de cuenta, le ponemos uno genérico para no perder la data
        cta_name = blk["nro_cuenta"] if blk["nro_cuenta"] else f"Desconocida_{i+1}"
        
        info_completa = {
            **info_global,
            "nro_cuenta": cta_name,
            "cbu": blk["cbu"],
            "saldo_inicial": blk["saldo_inicial"]
        }
        
        if movimientos:
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

    for c, h in enumerate(["Fecha", "Comprobante", "Descripción", "Categoría", "Débito ($)", "Crédito ($)", "Saldo ($)", ""], 1):
        cell = ws.cell(row=4, column=c, value=h)
        hdr(cell, bg="8B0000" if c in (5, 6, 7) else C_HEADER_BG)
    for i, w in enumerate([10, 12, 48, 22, 14, 14, 14, 2], 1): ws.column_dimensions[get_column_letter(i)].width = w

    fila = 5
    for mov in movimientos:
        bg = C_WARN if mov.get("sin_fecha") else C_DEBITO if (mov["debito"] and not mov["credito"]) else C_CREDITO if (mov["credito"] and not mov["debito"]) else "FFFFFF"
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

st.title("🏦 Conversor Santander a Excel")
st.write("Subí tus extractos en PDF. Si el archivo tiene varias cuentas adentro, te generaremos un Excel separado para cada una.")

uploaded_files = st.file_uploader("Arrastrá los PDF acá", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Procesar Archivos", type="primary"):
        for file in uploaded_files:
            with st.spinner(f"Escaneando {file.name}..."):
                try:
                    resultados = extraer_pdf_multicuenta(file)
                    
                    if resultados:
                        st.success(f"✅ {file.name} escaneado: ¡Se detectaron {len(resultados)} bloque(s) de movimientos!")
                        
                        for i, res in enumerate(resultados):
                            info = res["info"]
                            movs = res["movimientos"]
                            excel_buffer = crear_excel_buffer(info, movs)
                            
                            cuenta_limpia = info['nro_cuenta'].replace('/', '-')
                            nombre_excel = f"Extracto_{cuenta_limpia}_{i+1}.xlsx"
                            
                            st.download_button(
                                label=f"📥 Descargar Excel — Cuenta {info['nro_cuenta']} ({len(movs)} movs)",
                                data=excel_buffer,
                                file_name=nombre_excel,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"{file.name}_btn_{i}"
                            )
                    else:
                        st.warning(f"⚠️ No se detectaron movimientos en {file.name}.")
                except Exception as e:
                    st.error(f"❌ Ocurrió un error al procesar {file.name}: {str(e)}")
