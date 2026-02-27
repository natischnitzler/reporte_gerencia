"""
Reporte Automático de Alertas - Temponovo
Envío: Lunes y jueves
Adjuntos: Excel descuentos + PDF cobranza
"""

import xmlrpc.client
import smtplib
import ssl
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, date, timedelta
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ══════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════
ODOO_URL      = os.environ.get("ODOO_URL", "")
ODOO_DB       = os.environ.get("ODOO_DB", "temponovo")
ODOO_USER     = os.environ.get("ODOO_USER", "")
ODOO_PASSWORD = os.environ.get("ODOO_PASSWORD", "")

SMTP_HOST     = os.environ.get("SMTP_HOST", "srv10.akkuarios.com")
SMTP_PORT     = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER     = os.environ.get("SMTP_USER", "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")

DESTINATARIOS = [
    "daniel@temponovo.cl",
    "natalia@temponovo.cl",
]

# ── Umbrales ──
DESC_AMARILLO   = 30.0
DESC_ROJO       = 50.0
DESC_DIAS       = 3
COBR_DIAS       = 30
DIAS_COTIZACION = 3
DIAS_SIN_PICK   = 3
DIAS_SIN_OUT    = 3


# ══════════════════════════════════════════════
# CONEXIÓN ODOO
# ══════════════════════════════════════════════
def conectar_odoo():
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(ODOO_DB, ODOO_USER, ODOO_PASSWORD, {})
    if not uid:
        raise Exception("Autenticación fallida.")
    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
    return uid, models

def buscar(models, uid, modelo, dominio, campos, limite=1000):
    return models.execute_kw(
        ODOO_DB, uid, ODOO_PASSWORD,
        modelo, 'search_read',
        [dominio],
        {'fields': campos, 'limit': limite}
    )


# ══════════════════════════════════════════════
# 1. DESCUENTOS > 30% — últimos 3 días
# ══════════════════════════════════════════════
def get_descuentos(models, uid):
    hoy        = date.today()
    fecha_from = (hoy - timedelta(days=DESC_DIAS)).strftime('%Y-%m-%d')

    lineas_pedido = buscar(models, uid,
        'sale.order.line',
        [
            ['discount', '>', DESC_AMARILLO],
            ['order_id.state', 'in', ['sale', 'done']],
            ['order_id.date_order', '>=', fecha_from],
        ],
        ['order_id', 'product_id', 'discount', 'price_unit', 'product_uom_qty', 'price_subtotal']
    )

    lineas_factura = buscar(models, uid,
        'account.move.line',
        [
            ['discount', '>', DESC_AMARILLO],
            ['move_id.move_type', '=', 'out_invoice'],
            ['move_id.state', '=', 'posted'],
            ['display_type', '=', 'product'],
            ['move_id.invoice_date', '>=', fecha_from],
        ],
        ['move_id', 'partner_id', 'product_id', 'discount', 'price_unit', 'quantity', 'price_subtotal']
    )

    # Obtener info de pedidos (cliente, fecha)
    pedido_ids = list({l['order_id'][0] for l in lineas_pedido if l['order_id']})
    pedidos_info = {}
    if pedido_ids:
        raw = buscar(models, uid, 'sale.order',
            [['id', 'in', pedido_ids]], ['id', 'partner_id', 'date_order', 'name'])
        pedidos_info = {p['id']: p for p in raw}

    resumen = []
    detalle = []

    for l in lineas_pedido:
        oid    = l['order_id'][0] if l['order_id'] else None
        pinfo  = pedidos_info.get(oid, {})
        cliente= pinfo.get('partner_id', [None,''])[1] if pinfo else ''
        pedido = pinfo.get('name', '')
        fecha  = pinfo.get('date_order', '')[:10] if pinfo else ''
        prod   = l['product_id'][1] if l['product_id'] else ''
        codigo = prod.split(']')[0].replace('[','').strip() if ']' in prod else ''
        nombre = prod.split('] ')[-1] if ']' in prod else prod

        existe = next((r for r in resumen if r['N° Pedido'] == pedido), None)
        if not existe:
            resumen.append({'Cliente': cliente, 'N° Pedido': pedido,
                            'Fecha': fecha, 'Descuento': l['discount']})
        elif l['discount'] > existe['Descuento']:
            existe['Descuento'] = l['discount']

        detalle.append({
            'Tipo': 'Pedido', 'Cliente': cliente, 'N° Pedido': pedido, 'Fecha': fecha,
            'Código': codigo, 'Producto': nombre,
            'Precio Unit': l['price_unit'], 'Descuento %': l['discount'],
            'Cantidad': l['product_uom_qty'], 'Subtotal': l['price_subtotal'],
        })

    for l in lineas_factura:
        cliente = l['partner_id'][1] if l['partner_id'] else ''
        factura = l['move_id'][1] if l['move_id'] else ''
        prod    = l['product_id'][1] if l['product_id'] else ''
        codigo  = prod.split(']')[0].replace('[','').strip() if ']' in prod else ''
        nombre  = prod.split('] ')[-1] if ']' in prod else prod

        existe = next((r for r in resumen if r['N° Pedido'] == factura), None)
        if not existe:
            resumen.append({'Cliente': cliente, 'N° Pedido': factura,
                            'Fecha': '', 'Descuento': l['discount']})
        elif l['discount'] > existe['Descuento']:
            existe['Descuento'] = l['discount']

        detalle.append({
            'Tipo': 'Factura', 'Cliente': cliente, 'N° Pedido': factura, 'Fecha': '',
            'Código': codigo, 'Producto': nombre,
            'Precio Unit': l['price_unit'], 'Descuento %': l['discount'],
            'Cantidad': l['quantity'], 'Subtotal': l['price_subtotal'],
        })

    resumen.sort(key=lambda x: x['Descuento'], reverse=True)
    detalle.sort(key=lambda x: x['Descuento %'], reverse=True)
    return resumen, detalle


# ══════════════════════════════════════════════
# 2. COBRANZA VENCIDA
# ══════════════════════════════════════════════
def get_cobranza(models, uid):
    hoy = date.today()

    facturas = buscar(models, uid,
        'account.move',
        [
            ['move_type', '=', 'out_invoice'],
            ['payment_state', 'in', ['not_paid', 'partial']],
            ['state', '=', 'posted'],
        ],
        ['name', 'partner_id', 'invoice_date_due', 'amount_residual', 'invoice_user_id']
    )

    clientes = {}
    for f in facturas:
        pid = f['partner_id'][0] if f['partner_id'] else 0
        if pid not in clientes:
            clientes[pid] = {
                'Cliente': f['partner_id'][1] if f['partner_id'] else '',
                'Vendedor': f['invoice_user_id'][1] if f['invoice_user_id'] else 'Sin vendedor',
                'Ciudad': '',
                'A la fecha': 0.0, '1-30': 0.0, 'Vencido >30': 0.0, 'Total': 0.0,
                'facturas': [],
            }

        venc_str = f.get('invoice_date_due') or ''
        monto    = f['amount_residual'] or 0.0

        if venc_str:
            venc = datetime.strptime(venc_str, '%Y-%m-%d').date()
            dias = (hoy - venc).days
            if dias <= 0:
                clientes[pid]['A la fecha'] += monto
            elif dias <= 30:
                clientes[pid]['1-30'] += monto
            else:
                clientes[pid]['Vencido >30'] += monto
        else:
            clientes[pid]['A la fecha'] += monto

        clientes[pid]['Total'] += monto
        dias_v = (hoy - datetime.strptime(venc_str, '%Y-%m-%d').date()).days if venc_str else 0
        clientes[pid]['facturas'].append({
            'Factura': f['name'], 'Fecha Venc.': venc_str,
            'Días Vencido': dias_v, 'Monto Pendiente': monto,
        })

    if clientes:
        partners = buscar(models, uid, 'res.partner',
            [['id', 'in', list(clientes.keys())]], ['id', 'city'])
        for p in partners:
            if p['id'] in clientes:
                clientes[p['id']]['Ciudad'] = p.get('city') or ''

    resumen = [v for v in clientes.values() if v['Vencido >30'] > 0]
    resumen.sort(key=lambda x: x['Vencido >30'], reverse=True)
    return resumen


# ══════════════════════════════════════════════
# 3. PEDIDOS ATRASADOS
# ══════════════════════════════════════════════
def get_pedidos_atrasados(models, uid):
    hoy = date.today()
    cotizaciones, no_pickeados, no_en_bulto = [], [], []

    cots = buscar(models, uid, 'sale.order',
        [
            ['state', '=', 'draft'],
            ['date_order', '<', (hoy - timedelta(days=DIAS_COTIZACION)).strftime('%Y-%m-%d %H:%M:%S')],
        ],
        ['name', 'partner_id', 'date_order', 'amount_total', 'user_id']
    )
    for p in cots:
        dias = (hoy - datetime.strptime(p['date_order'][:10], '%Y-%m-%d').date()).days
        cotizaciones.append({
            'N° Pedido': p['name'],
            'Cliente':   p['partner_id'][1] if p['partner_id'] else '',
            'Vendedor':  p['user_id'][1] if p['user_id'] else '',
            'Estado':    'Sin confirmar',
            'Días':      dias,
        })

    confirmados = buscar(models, uid, 'sale.order',
        [
            ['state', '=', 'sale'],
            ['date_order', '<', (hoy - timedelta(days=DIAS_SIN_PICK)).strftime('%Y-%m-%d %H:%M:%S')],
        ],
        ['name', 'partner_id', 'date_order', 'amount_total', 'user_id', 'picking_ids']
    )

    for p in confirmados:
        picking_ids = p.get('picking_ids', [])
        dias = (hoy - datetime.strptime(p['date_order'][:10], '%Y-%m-%d').date()).days
        base = {
            'N° Pedido': p['name'],
            'Cliente':   p['partner_id'][1] if p['partner_id'] else '',
            'Vendedor':  p['user_id'][1] if p['user_id'] else '',
            'Días':      dias,
        }

        if not picking_ids:
            no_pickeados.append({**base, 'Estado': 'No pickeado'})
            continue

        pickings = buscar(models, uid, 'stock.picking',
            [['id', 'in', picking_ids]], ['name', 'state'])
        picks = [pk for pk in pickings if 'PICK' in (pk.get('name') or '')]
        outs  = [pk for pk in pickings if 'OUT'  in (pk.get('name') or '')]

        pick_ok = any(pk['state'] == 'done' for pk in picks)
        out_ok  = any(pk['state'] == 'done' for pk in outs)

        if not pick_ok:
            no_pickeados.append({**base, 'Estado': 'No pickeado'})
        elif not out_ok and dias >= DIAS_SIN_OUT:
            no_en_bulto.append({**base, 'Estado': 'No en bulto'})

    cotizaciones.sort(key=lambda x: x['Días'], reverse=True)
    no_pickeados.sort(key=lambda x: x['Días'], reverse=True)
    no_en_bulto.sort(key=lambda x: x['Días'], reverse=True)

    return cotizaciones + no_pickeados + no_en_bulto


# ══════════════════════════════════════════════
# EXCEL — DESCUENTOS
# ══════════════════════════════════════════════
AZUL    = '1B3A6B'
BLANCO  = 'FFFFFF'
ROJO_BG = 'FFEBEE'
AMA_BG  = 'FFFDE7'
GRIS_BG = 'F5F5F5'

def _h(cell):
    cell.font      = Font(bold=True, color=BLANCO, name='Arial', size=10)
    cell.fill      = PatternFill('solid', start_color=AZUL)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border    = Border(bottom=Side(style='thin', color='CCCCCC'),
                            right =Side(style='thin', color='CCCCCC'))

def _d(cell, bg=BLANCO, bold=False):
    cell.font      = Font(name='Arial', size=9, bold=bold)
    cell.fill      = PatternFill('solid', start_color=bg)
    cell.alignment = Alignment(vertical='center')
    cell.border    = Border(bottom=Side(style='thin', color='E0E0E0'),
                            right =Side(style='thin', color='E0E0E0'))

def excel_descuentos(detalle):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Descuentos"
    ws.sheet_properties.tabColor = "D32F2F"

    headers = ['Tipo','Cliente','N° Pedido','Fecha','Código','Producto',
               'Precio Unit','Descuento %','Cantidad','Subtotal']
    n = len(headers)

    # Título
    ws.merge_cells(f'A1:{get_column_letter(n)}1')
    c = ws['A1']
    c.value     = f'DESCUENTOS > {int(DESC_AMARILLO)}% — Últimos {DESC_DIAS} días'
    c.font      = Font(bold=True, size=13, color=BLANCO, name='Arial')
    c.fill      = PatternFill('solid', start_color=AZUL)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 32

    for i, h in enumerate(headers, 1):
        _h(ws.cell(row=2, column=i, value=h))
    ws.row_dimensions[2].height = 24

    for ri, row in enumerate(detalle, 3):
        bg = ROJO_BG if row.get('Descuento %', 0) >= DESC_ROJO else AMA_BG
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=ri, column=ci, value=row.get(h, ''))
            _d(c, bg)
            if h in ('Precio Unit', 'Subtotal'):
                c.number_format = '$#,##0;($#,##0);"-"'
            elif h == 'Descuento %':
                c.number_format = '0.0"%"'

    if detalle:
        tr = len(detalle) + 3
        ws.cell(row=tr, column=1, value='TOTAL').font = Font(bold=True, name='Arial', size=9)
        c = ws.cell(row=tr, column=10, value=f'=SUM(J3:J{tr-1})')
        _d(c, GRIS_BG, bold=True)
        c.number_format = '$#,##0;($#,##0);"-"'

    for i, h in enumerate(headers, 1):
        w = max(len(h)+4, max((min(len(str(r.get(h,'')))+2,45) for r in detalle), default=10))
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A3'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════
# PDF — COBRANZA
# ══════════════════════════════════════════════
def fmt_clp(v):
    try: return f"$ {int(v):,}".replace(',','.')
    except: return str(v)

def pdf_cobranza(resumen):
    buf    = io.BytesIO()
    doc    = SimpleDocTemplate(buf, pagesize=landscape(A4),
                               leftMargin=1.5*cm, rightMargin=1.5*cm,
                               topMargin=1.5*cm, bottomMargin=1.5*cm)
    AZUL_RL = colors.HexColor('#1B3A6B')
    ROJO_RL = colors.HexColor('#C62828')
    styles  = getSampleStyleSheet()
    hoy_str = date.today().strftime('%d/%m/%Y')
    elementos = []

    # Header
    header = Table([[f'COBRANZA PENDIENTE COMPLETA — TEMPONOVO    {hoy_str}']],
                   colWidths=[26*cm])
    header.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), AZUL_RL),
        ('TEXTCOLOR',     (0,0),(-1,-1), colors.white),
        ('FONTNAME',      (0,0),(-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE',      (0,0),(-1,-1), 14),
        ('ALIGN',         (0,0),(-1,-1), 'CENTER'),
        ('TOPPADDING',    (0,0),(-1,-1), 14),
        ('BOTTOMPADDING', (0,0),(-1,-1), 14),
    ]))
    elementos.append(header)
    elementos.append(Spacer(1, 0.5*cm))

    # Tabla
    col_widths = [7*cm, 3.5*cm, 3*cm, 3.5*cm, 3.5*cm, 5.5*cm]
    data = [['Cliente · Vendedor · Ciudad', 'A la fecha', '1-30 días',
             'Vencido >30', 'Total', 'Facturas vencidas']]

    for r in resumen:
        facs_vencidas = ', '.join(
            f['Factura'] for f in r['facturas'] if f['Días Vencido'] > 30
        )
        data.append([
            f"{r['Cliente']}\n{r['Vendedor']} · {r['Ciudad']}",
            fmt_clp(r['A la fecha']) if r['A la fecha'] else '—',
            fmt_clp(r['1-30'])       if r['1-30']       else '—',
            fmt_clp(r['Vencido >30']),
            fmt_clp(r['Total']),
            facs_vencidas,
        ])

    # Fila total
    data.append([
        'TOTAL',
        fmt_clp(sum(r['A la fecha']  for r in resumen)),
        fmt_clp(sum(r['1-30']        for r in resumen)),
        fmt_clp(sum(r['Vencido >30'] for r in resumen)),
        fmt_clp(sum(r['Total']       for r in resumen)),
        '',
    ])

    n = len(data)
    tabla = Table(data, colWidths=col_widths, repeatRows=1)
    tabla.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,0),   AZUL_RL),
        ('TEXTCOLOR',     (0,0), (-1,0),   colors.white),
        ('FONTNAME',      (0,0), (-1,0),   'Helvetica-Bold'),
        ('FONTSIZE',      (0,0), (-1,-1),  8),
        ('ALIGN',         (1,0), (-1,-1),  'RIGHT'),
        ('ALIGN',         (0,0), (0,-1),   'LEFT'),
        ('VALIGN',        (0,0), (-1,-1),  'MIDDLE'),
        ('TOPPADDING',    (0,0), (-1,-1),  5),
        ('BOTTOMPADDING', (0,0), (-1,-1),  5),
        ('TEXTCOLOR',     (3,1), (3,n-2),  ROJO_RL),
        ('FONTNAME',      (3,1), (3,n-2),  'Helvetica-Bold'),
        *[('BACKGROUND',  (0,i), (-1,i),   colors.HexColor('#F5F5F5'))
          for i in range(2, n-1, 2)],
        ('BACKGROUND',    (0,n-1),(-1,n-1),colors.HexColor('#E3E8F0')),
        ('FONTNAME',      (0,n-1),(-1,n-1),'Helvetica-Bold'),
        ('GRID',          (0,0), (-1,-1),  0.3, colors.HexColor('#CCCCCC')),
    ]))
    elementos.append(tabla)
    doc.build(elementos)
    buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════
# EMAIL HTML
# ══════════════════════════════════════════════
def generar_html(desc_res, cobr_res, pedidos):
    hoy = date.today().strftime('%d/%m/%Y')

    def fmt(v):
        try: return f"$ {int(v):,}".replace(',','.')
        except: return str(v)

    def tabla_desc(rows):
        if not rows:
            return '<p style="color:#4caf50;font-style:italic;">✅ Sin descuentos altos en los últimos 3 días</p>'
        t = '''<table style="width:100%;border-collapse:collapse;font-size:12px;">
        <tr>
          <th style="background:#1B3A6B;color:#fff;padding:8px 12px;text-align:left;">Cliente</th>
          <th style="background:#1B3A6B;color:#fff;padding:8px 12px;text-align:left;">N° Pedido</th>
          <th style="background:#1B3A6B;color:#fff;padding:8px 12px;text-align:center;">Fecha</th>
          <th style="background:#1B3A6B;color:#fff;padding:8px 12px;text-align:center;">Descuento</th>
        </tr>'''
        for r in rows:
            bg = '#FFEBEE' if r['Descuento'] >= DESC_ROJO else '#FFFDE7'
            t += f'''<tr style="background:{bg};">
              <td style="padding:7px 12px;border-bottom:1px solid #eee;">{r["Cliente"]}</td>
              <td style="padding:7px 12px;border-bottom:1px solid #eee;">{r["N° Pedido"]}</td>
              <td style="padding:7px 12px;border-bottom:1px solid #eee;text-align:center;">{r["Fecha"]}</td>
              <td style="padding:7px 12px;border-bottom:1px solid #eee;text-align:center;font-weight:bold;">{r["Descuento"]:.1f}%</td>
            </tr>'''
        t += '</table>'
        t += '<p style="font-size:11px;color:#888;margin-top:6px;">📎 Ver detalle por producto en el Excel adjunto</p>'
        return t

    def tabla_cobr(rows):
        if not rows:
            return '<p style="color:#4caf50;font-style:italic;">✅ Sin facturas vencidas</p>'
        t = '''<table style="width:100%;border-collapse:collapse;font-size:12px;">
        <tr>
          <th style="background:#1B3A6B;color:#fff;padding:8px 12px;text-align:left;">Cliente · Vendedor · Ciudad</th>
          <th style="background:#1B3A6B;color:#fff;padding:8px 12px;text-align:right;">A la fecha</th>
          <th style="background:#1B3A6B;color:#fff;padding:8px 12px;text-align:right;">1-30 días</th>
          <th style="background:#1B3A6B;color:#fff;padding:8px 12px;text-align:right;">Vencido &gt;30</th>
          <th style="background:#1B3A6B;color:#fff;padding:8px 12px;text-align:right;">Total</th>
        </tr>'''
        for r in rows:
            info = f"<strong>{r['Cliente']}</strong><br><small style='color:#888;'>{r['Vendedor']} · {r['Ciudad']}</small>"
            t += f'''<tr>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;">{info}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">{fmt(r["A la fecha"]) if r["A la fecha"] else "—"}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">{fmt(r["1-30"]) if r["1-30"] else "—"}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;font-weight:bold;color:#c62828;">{fmt(r["Vencido >30"])}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;font-weight:bold;">{fmt(r["Total"])}</td>
            </tr>'''
        t += f'''<tr style="background:#f0f4f8;">
          <td style="padding:8px 12px;font-weight:bold;border-top:2px solid #ddd;">TOTAL</td>
          <td style="padding:8px 12px;text-align:right;font-weight:bold;border-top:2px solid #ddd;">{fmt(sum(r["A la fecha"] for r in rows))}</td>
          <td style="padding:8px 12px;text-align:right;font-weight:bold;border-top:2px solid #ddd;">{fmt(sum(r["1-30"] for r in rows))}</td>
          <td style="padding:8px 12px;text-align:right;font-weight:bold;color:#c62828;border-top:2px solid #ddd;">{fmt(sum(r["Vencido >30"] for r in rows))}</td>
          <td style="padding:8px 12px;text-align:right;font-weight:bold;border-top:2px solid #ddd;">{fmt(sum(r["Total"] for r in rows))}</td>
        </tr>'''
        t += '</table>'
        t += '<p style="font-size:11px;color:#888;margin-top:6px;">📎 Ver detalle completo en el PDF adjunto</p>'
        return t

    ESTADO_BG    = {'Sin confirmar': '#F5F5F5', 'No pickeado': '#F5F5F5', 'No en bulto': '#F5F5F5'}
    ESTADO_COLOR = {'Sin confirmar': '#1B3A6B', 'No pickeado': '#1B3A6B', 'No en bulto': '#1B3A6B'}
    ESTADO_ICON  = {'Sin confirmar': '📋', 'No pickeado': '📦', 'No en bulto': '🚚'}

    def tabla_ped(rows):
        if not rows:
            return '<p style="color:#4caf50;font-style:italic;">✅ Sin pedidos atrasados</p>'
        t = '''<table style="width:100%;border-collapse:collapse;font-size:12px;">
        <tr>
          <th style="background:#1B3A6B;color:#fff;padding:8px 12px;text-align:left;">N° Pedido</th>
          <th style="background:#1B3A6B;color:#fff;padding:8px 12px;text-align:left;">Cliente</th>
          <th style="background:#1B3A6B;color:#fff;padding:8px 12px;text-align:left;">Vendedor</th>
          <th style="background:#1B3A6B;color:#fff;padding:8px 12px;text-align:center;">Estado</th>
          <th style="background:#1B3A6B;color:#fff;padding:8px 12px;text-align:center;">Días</th>
        </tr>'''
        estado_actual = None
        for r in rows:
            if r['Estado'] != estado_actual:
                estado_actual = r['Estado']
                color = ESTADO_COLOR.get(estado_actual, '#555')
                icon  = ESTADO_ICON.get(estado_actual, '')
                t += f'<tr><td colspan="5" style="background:{color};color:#fff;padding:6px 12px;font-weight:bold;font-size:11px;">{icon} {estado_actual.upper()}</td></tr>'
            bg = ESTADO_BG.get(r['Estado'], '#fff')
            t += f'''<tr style="background:{bg};">
              <td style="padding:7px 12px;border-bottom:1px solid #eee;font-weight:bold;">{r["N° Pedido"]}</td>
              <td style="padding:7px 12px;border-bottom:1px solid #eee;">{r["Cliente"]}</td>
              <td style="padding:7px 12px;border-bottom:1px solid #eee;color:#666;">{r["Vendedor"]}</td>
              <td style="padding:7px 12px;border-bottom:1px solid #eee;text-align:center;">{r["Estado"]}</td>
              <td style="padding:7px 12px;border-bottom:1px solid #eee;text-align:center;font-weight:bold;">{r["Días"]}d</td>
            </tr>'''
        return t + '</table>'

    def seccion(emoji, titulo, n, color, contenido):
        return f'''<div style="margin-bottom:32px;">
          <h3 style="margin:0 0 12px;color:#1B3A6B;font-size:15px;border-left:4px solid {color};padding-left:12px;">
            {emoji} {titulo}
            <span style="margin-left:8px;background:{color};color:#fff;padding:2px 9px;border-radius:12px;font-size:12px;">{n}</span>
          </h3>
          {contenido}
        </div>'''

    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;background:#f0f2f5;margin:0;padding:24px;">
<div style="max-width:820px;margin:auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.08);">

  <div style="background:#1B3A6B;padding:28px 32px;">
    <h1 style="color:#fff;margin:0;font-size:22px;">Reporte Temponovo</h1>
    <p style="color:#8eadd4;margin:6px 0 0;font-size:13px;">{hoy}</p>
  </div>

  <div style="padding:28px 32px;">
    {seccion('🏷️', f'Descuentos superiores al {int(DESC_AMARILLO)}%', len(desc_res), '#D32F2F', tabla_desc(desc_res))}
    {seccion('💸', 'Cobranza vencida', len(cobr_res), '#E65100', tabla_cobr(cobr_res))}
    {seccion('📦', 'Pedidos atrasados', len(pedidos), '#1565C0', tabla_ped(pedidos))}
  </div>

  <div style="background:#f8f9fb;padding:14px 32px;border-top:1px solid #e8eaed;text-align:center;">
    <p style="margin:0;font-size:11px;color:#aaa;">Reporte automático · Temponovo · Odoo ERP</p>
  </div>
</div>
</body></html>"""


# ══════════════════════════════════════════════
# ENVÍO EMAIL
# ══════════════════════════════════════════════
def enviar_email(html, excel_desc_bytes, pdf_cobr_bytes):
    fecha_str = date.today().strftime('%Y%m%d')

    msg = MIMEMultipart('mixed')
    msg['From']    = SMTP_USER
    msg['To']      = ', '.join(DESTINATARIOS)
    msg['Subject'] = f"Reporte Temponovo — {date.today().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText(html, 'html', 'utf-8'))

    part_xlsx = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    part_xlsx.set_payload(excel_desc_bytes)
    encoders.encode_base64(part_xlsx)
    part_xlsx.add_header('Content-Disposition', f'attachment; filename="descuentos_{fecha_str}.xlsx"')
    msg.attach(part_xlsx)

    part_pdf = MIMEBase('application', 'pdf')
    part_pdf.set_payload(pdf_cobr_bytes)
    encoders.encode_base64(part_pdf)
    part_pdf.add_header('Content-Disposition', f'attachment; filename="cobranza_{fecha_str}.pdf"')
    msg.attach(part_pdf)

    ctx = ssl.create_default_context()
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as srv:
        srv.ehlo()
        srv.starttls(context=ctx)
        srv.login(SMTP_USER, SMTP_PASSWORD)
        srv.sendmail(SMTP_USER, DESTINATARIOS, msg.as_bytes())

    print(f"✅ Email enviado")


# ══════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════
def main():
    print(f"\n🚀 Reporte Temponovo — {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    uid, models = conectar_odoo()
    print("✅ Conectado a Odoo")

    print("  🏷️  Descuentos...")
    desc_res, desc_det = get_descuentos(models, uid)
    print(f"     {len(desc_res)} pedidos/facturas con descuento alto")

    print("  💸 Cobranza...")
    cobr_res, cobr_todos = get_cobranza(models, uid)
    print(f"     {len(cobr_res)} clientes con deuda vencida >30d")

    print("  📦 Pedidos atrasados...")
    pedidos = get_pedidos_atrasados(models, uid)
    print(f"     {len(pedidos)} pedidos atrasados")

    print("  📄 Generando adjuntos...")
    excel_desc = excel_descuentos(desc_det)
    pdf_cobr   = pdf_cobranza(cobr_todos)

    print("  📧 Enviando email...")
    html = generar_html(desc_res, cobr_res, pedidos)
    enviar_email(html, excel_desc, pdf_cobr)

    print("✅ Proceso completado\n")


if __name__ == "__main__":
    main()
