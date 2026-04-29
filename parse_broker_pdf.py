"""
parse_broker_pdf.py
───────────────────
Extracts structured transaction data from broker account statement PDFs.
Designed for COFCO INTL format, extensible to other brokers via config.

Usage:
    df_all = parse_broker_pdf("COFCO.pdf")
    # Returns a DataFrame with all transactions parsed from the PDF

Returns DataFrame with columns:
    seccion, fecha, vto, tipo_doc, comprobante, comprobante_orig,
    concepto, moneda, contrato, importe, saldo
"""

import pdfplumber
import re
import pandas as pd
from pathlib import Path


# ─────────────────────────────────────────────────────────────
# NUMBER UTILITIES
# ─────────────────────────────────────────────────────────────

# Matches: 1.234.567,89  or  1.234.567,89-
_NUM_RE = re.compile(r'[\d]+(?:\.[\d]{3})*,[\d]+(-)?')

def _parse_num(s):
    """Parse Argentine/EU number format '1.234,56' or '1.234,56-' → float."""
    if not s or s.strip() in ('-', '', '—', '0,00'):
        return 0.0
    s = str(s).strip()
    neg = s.endswith('-')
    s = s.rstrip('-').strip()
    s = s.replace('.', '').replace(',', '.')
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return 0.0

def _extract_nums(text):
    """Extract all numeric values from a text string, preserving sign."""
    return [_parse_num(m.group()) for m in _NUM_RE.finditer(text)]

def _strip_nums(text):
    """Remove all trailing numeric tokens from a line (leaving the description)."""
    # Remove everything from the first number cluster onward
    m = _NUM_RE.search(text)
    if m:
        return text[:m.start()].rstrip()
    return text.rstrip()


# ─────────────────────────────────────────────────────────────
# DOCUMENT TYPE DETECTION
# ─────────────────────────────────────────────────────────────

def _detect_doc_type(middle):
    """
    Classify the middle portion of a transaction line into a document type
    and extract (tipo_doc, comprobante, comprobante_orig, concepto, moneda, contrato).
    """
    m = middle.strip()

    # --- Patterns ordered from most specific to most general ---

    # Nota de Crédito / Nota de Débito  (FX differences)
    nc_nd = re.match(
        r'(Nota de (?:Crédito|Debito|Débito|Credito))\s+(\S+)\s+(\S+)\s+(.*?)(?:\s+(USD|ARS))?$',
        m, re.IGNORECASE)
    if nc_nd:
        tipo = 'Nota de Crédito' if re.search(r'cr[eé]d', nc_nd.group(1), re.IGNORECASE) else 'Nota de Débito'
        comp = nc_nd.group(2)
        comp_orig = nc_nd.group(3)
        rest = nc_nd.group(4).strip()
        moneda = nc_nd.group(5) or _detect_currency(m)
        contrato = _extract_contrato(rest + ' ' + comp_orig)
        concepto = rest
        return tipo, comp, comp_orig, concepto, moneda, contrato

    # Formulario 1116B
    f116 = re.match(
        r'Formulario\s+1116B\s+(\S+)\s+(.*?)\s+(ARS|USD)\s+(\S+)',
        m, re.IGNORECASE)
    if f116:
        comp = f116.group(1)         # e.g. 3301-30685269
        concepto = f116.group(2).strip()  # e.g. "Mercadería Parcial"
        moneda = f116.group(3)
        contrato = f116.group(4)
        return 'Formulario 1116B', comp, '', concepto, moneda, contrato

    # Factura normal
    fac = re.match(r'Factura\s+(\S+)\s+(.*?)\s+(ARS|USD)\s+(\S+)', m, re.IGNORECASE)
    if fac:
        comp = fac.group(1)
        concepto = fac.group(2).strip()
        moneda = fac.group(3)
        contrato = fac.group(4)
        return 'Factura', comp, '', concepto, moneda, contrato

    # Factura sin contrato (e.g. anticipo NPC)
    fac2 = re.match(r'Factura\s+(\S+)\s+(.*?)\s+(ARS|USD)\s*$', m, re.IGNORECASE)
    if fac2:
        comp = fac2.group(1)
        concepto = fac2.group(2).strip()
        moneda = fac2.group(3)
        contrato = _extract_contrato(m)
        return 'Factura', comp, '', concepto, moneda, contrato

    # Retención IIBB / IVA / Otro
    ret = re.match(
        r'([\dA-Z\-]+)\s+(Ret\s+\S+.*?)\s+C\.N[oº°]?\s+(\S+)\s+(ARS|USD)',
        m, re.IGNORECASE)
    if ret:
        comp = ret.group(1)          # 3301-NNNNNN
        concepto = ret.group(2).strip()
        moneda = ret.group(4)
        cert = ret.group(3)          # certificate number
        return _ret_tipo(concepto), comp, cert, concepto, moneda, ''

    # Transferencia Bancaria
    if re.search(r'transferencia\s+bancaria', m, re.IGNORECASE):
        moneda = _detect_currency(m)
        return 'Transferencia Bancaria', 'TRANSF', '', 'Datanet Pago', moneda, ''

    # Recibo
    rec = re.match(r'Recibo\s+(\S+)\s+(.*?)\s+(ARS|USD)', m, re.IGNORECASE)
    if rec:
        return 'Recibo', rec.group(1), '', rec.group(2).strip(), rec.group(3), ''

    # Fallback: grab any doc-like token
    moneda = _detect_currency(m)
    tokens = m.split()
    comp = tokens[0] if tokens else ''
    concepto = ' '.join(tokens[1:]) if len(tokens) > 1 else m
    contrato = _extract_contrato(m)
    return 'Otro', comp, '', concepto, moneda, contrato


def _ret_tipo(concepto):
    c = concepto.upper()
    if 'IIBB' in c:
        return 'Ret IIBB'
    if 'IVA' in c:
        return 'Ret IVA'
    return 'Retención'


def _detect_currency(text):
    if re.search(r'\bUSD\b', text):
        return 'USD'
    if re.search(r'\bARS\b', text):
        return 'ARS'
    return ''


def _extract_contrato(text):
    """Extract contract codes like GBU0001452908, GRO0001484474, NPC, etc."""
    m = re.search(r'\b([A-Z]{2,4}\d{5,})\b', text)
    if m:
        return m.group(1)
    # Short codes: NPC, GYO (appear in DC section)
    m2 = re.search(r'\b(NPC|GYO|GAL|GBU|GRO)\b', text)
    if m2:
        return m2.group(1)
    return ''


# ─────────────────────────────────────────────────────────────
# SECTION DETECTION
# ─────────────────────────────────────────────────────────────

def _detect_section(line):
    """
    Return new section name if this line is a section header, else None.
    Sections: USD_vencido, USD_avencer, ARS_vencido, ARS_avencer, DIF_CAMBIO
    """
    l = line.strip()
    if re.search(r'CUENTA EN DOLARES.*Saldo\s*v\s*enc', l, re.IGNORECASE):
        return 'USD_vencido'
    if re.search(r'CUENTA EN DOLARES.*Saldo\s*a\s*v', l, re.IGNORECASE):
        return 'USD_avencer'
    if re.search(r'CUENTA EN PESOS.*Saldo\s*vencido', l, re.IGNORECASE):
        return 'ARS_vencido'
    if re.search(r'CUENTA EN PESOS.*Saldo\s*a\s*vencer', l, re.IGNORECASE):
        return 'ARS_avencer'
    if re.search(r'DIFERENCIA DE CAMBIO', l, re.IGNORECASE):
        return 'DIF_CAMBIO'
    return None


# ─────────────────────────────────────────────────────────────
# LINE PARSER
# ─────────────────────────────────────────────────────────────

_DATE_RE = re.compile(r'^\s*(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2})\.(\d{2})\.(\d{4})\s+(.*)')

def _parse_line(line, seccion):
    """
    Parse a single transaction line into a dict.
    Returns None if not a transaction line.
    """
    m = _DATE_RE.match(line)
    if not m:
        return None

    fecha = f"{m.group(3)}-{m.group(2)}-{m.group(1)}"   # YYYY-MM-DD
    vto   = f"{m.group(6)}-{m.group(5)}-{m.group(4)}"
    rest  = m.group(7)

    # Extract trailing numbers (importe + saldo)
    nums = _extract_nums(rest)
    desc = _strip_nums(rest)

    # Amounts: last number = saldo, second-to-last = importe
    # Some rows only have 1 number visible when saldo=0 may be blank
    if len(nums) >= 2:
        importe = nums[-2]
        saldo   = nums[-1]
    elif len(nums) == 1:
        importe = nums[0]
        saldo   = 0.0
    else:
        importe = 0.0
        saldo   = 0.0

    # In the DIF_CAMBIO section, only 1 amount column is shown (no saldo column)
    if seccion == 'DIF_CAMBIO':
        if len(nums) == 1:
            importe = nums[0]
            saldo = 0.0
        elif len(nums) >= 2:
            # first number is importe, second could be saldo (usually 0 or same)
            importe = nums[0]
            saldo = nums[1] if len(nums) > 1 else 0.0

    tipo_doc, comp, comp_orig, concepto, moneda, contrato = _detect_doc_type(desc)

    # Infer moneda from section if not found in description
    if not moneda:
        if seccion.startswith('USD'):
            moneda = 'USD'
        elif seccion.startswith('ARS') or seccion == 'DIF_CAMBIO':
            moneda = 'ARS'

    return {
        'seccion':         seccion,
        'fecha':           fecha,
        'vto':             vto,
        'tipo_doc':        tipo_doc,
        'comprobante':     comp,
        'comprobante_orig': comp_orig,
        'concepto':        concepto,
        'moneda':          moneda,
        'contrato':        contrato,
        'importe':         importe,
        'saldo':           saldo,
    }


# ─────────────────────────────────────────────────────────────
# TOTALS / BALANCE EXTRACTOR
# ─────────────────────────────────────────────────────────────

def _extract_totals(pages_text):
    """Extract section totals and overall balances from PDF text."""
    totals = {}
    for line in pages_text.split('\n'):
        l = line.strip()
        nums = _extract_nums(l)
        if not nums:
            continue
        if re.search(r'TOTAL CUENTA EN DOLARES', l, re.IGNORECASE):
            totals['total_USD'] = nums[-1]
        elif re.search(r'TOTAL CUENTA EN PESOS', l, re.IGNORECASE):
            totals['total_ARS'] = nums[-1]
        elif re.search(r'SALDO.*CUENTA EN DOLARES', l, re.IGNORECASE) and 'saldo_USD' not in totals:
            totals['saldo_USD'] = nums[-1]
        elif re.search(r'SALDO.*CUENTA EN PESOS', l, re.IGNORECASE) and 'saldo_ARS' not in totals:
            totals['saldo_ARS'] = nums[-1]
    return totals


# ─────────────────────────────────────────────────────────────
# MAIN PARSER
# ─────────────────────────────────────────────────────────────

def parse_broker_pdf(pdf_path, x_density=7, y_density=13):
    """
    Parse a broker PDF account statement and return structured DataFrames.

    Parameters
    ----------
    pdf_path : str or Path
        Path to the PDF file.
    x_density : float
        Horizontal spacing for layout-aware text extraction (default=7).
    y_density : float
        Vertical spacing for layout-aware text extraction (default=13).

    Returns
    -------
    dict with keys:
        'transactions' : pd.DataFrame   – all parsed transaction rows
        'totals'       : dict           – extracted totals/balances from PDF
        'raw_text'     : str            – full extracted text (for debugging)
    """
    pdf_path = Path(pdf_path)
    rows = []
    all_text = []
    current_section = None

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(layout=True, x_density=x_density, y_density=y_density)
            if not text:
                continue
            all_text.append(text)

            for line in text.splitlines():
                # Check for section header
                new_sec = _detect_section(line)
                if new_sec:
                    current_section = new_sec
                    continue

                # Skip header/footer/total lines
                stripped = line.strip()
                if (not stripped
                        or re.match(r'^SALDO\s+CUENTA', stripped)
                        or re.search(r'TOTAL CUENTA', stripped)
                        or re.search(r'Estado de Cuenta', stripped, re.IGNORECASE)
                        or re.search(r'Página\s+\d', stripped, re.IGNORECASE)):
                    continue

                # Parse transaction line if we're in a section
                if current_section:
                    row = _parse_line(line, current_section)
                    if row:
                        rows.append(row)

    full_text = '\n'.join(all_text)
    totals = _extract_totals(full_text)

    df = pd.DataFrame(rows)
    if not df.empty:
        df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
        df['vto']   = pd.to_datetime(df['vto'],   errors='coerce')
        df['importe'] = pd.to_numeric(df['importe'], errors='coerce').fillna(0)
        df['saldo']   = pd.to_numeric(df['saldo'],   errors='coerce').fillna(0)

    return {
        'transactions': df,
        'totals': totals,
        'raw_text': full_text,
    }


# ─────────────────────────────────────────────────────────────
# CONVENIENCE FILTERS
# ─────────────────────────────────────────────────────────────

def get_formularios(df):
    """Return only Formulario 1116B rows."""
    return df[df['tipo_doc'] == 'Formulario 1116B'].copy()

def get_facturas(df, moneda=None):
    """Return Factura rows, optionally filtered by moneda (USD/ARS)."""
    mask = df['tipo_doc'] == 'Factura'
    if moneda:
        mask &= df['moneda'] == moneda
    return df[mask].copy()

def get_pagos(df):
    """Return Transferencia Bancaria (payment) rows."""
    return df[df['tipo_doc'] == 'Transferencia Bancaria'].copy()

def get_retenciones(df):
    """Return Retención rows (IIBB + IVA)."""
    return df[df['tipo_doc'].str.startswith('Ret', na=False)].copy()

def get_dif_cambio(df):
    """Return Diferencia de Cambio rows (NC/ND)."""
    return df[df['seccion'] == 'DIF_CAMBIO'].copy()

def get_pending(df):
    """Return rows with saldo ≠ 0 (open/outstanding items)."""
    return df[df['saldo'].abs() > 0.01].copy()


# ─────────────────────────────────────────────────────────────
# QUICK TEST
# ─────────────────────────────────────────────────────────────

if __name__ == '__main__':
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else 'mnt/CONCILIACION COFCO/COFCO.pdf'

    result = parse_broker_pdf(path)
    df = result['transactions']
    totals = result['totals']

    print(f"\n{'='*60}")
    print(f"  BROKER PDF PARSER — {Path(path).name}")
    print(f"{'='*60}")
    print(f"  Total rows parsed: {len(df)}")
    print(f"  Totals from PDF:   {totals}")
    print()

    for sec in ['USD_vencido', 'USD_avencer', 'ARS_vencido', 'ARS_avencer', 'DIF_CAMBIO']:
        sub = df[df['seccion'] == sec]
        if sub.empty:
            continue
        print(f"  [{sec}] — {len(sub)} rows")
        by_type = sub.groupby('tipo_doc')['importe'].sum()
        for t, s in by_type.items():
            print(f"    {t:<30}  {s:>18,.2f}")
        print(f"    {'TOTAL':<30}  {sub['importe'].sum():>18,.2f}")
        print()

    print("  Pending items (saldo ≠ 0):")
    pending = get_pending(df)
    for _, r in pending.iterrows():
        print(f"    {r['comprobante']:<20}  saldo={r['saldo']:>16,.2f} {r['moneda']}")
