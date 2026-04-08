import sys, json, base64 as b64mod, os, io
from docx import Document
from docx.shared import Cm

def set_cell_text(cell, text):
    for para in cell.paragraphs:
        for run in para.runs:
            run.text = ''
    if cell.paragraphs:
        cell.paragraphs[0].add_run(str(text or ''))

def add_image_to_cell(cell, img_b64):
    if not img_b64:
        return
    if ',' in img_b64:
        img_b64 = img_b64.split(',')[1]
    img_bytes = b64mod.b64decode(img_b64)
    try:
        from PIL import Image as PILImage
        pil_img = PILImage.open(io.BytesIO(img_bytes))
        pil_img.thumbnail((800, 600), PILImage.LANCZOS)
        buf = io.BytesIO()
        pil_img.save(buf, format='JPEG', quality=75)
        img_bytes = buf.getvalue()
    except:
        pass
    tmp = 'C:\\t\\g.jpg'
    os.makedirs('C:\\t', exist_ok=True)
    with open(tmp, 'wb') as f:
        f.write(img_bytes)
    try:
        para = cell.paragraphs[0]
        run = para.add_run()
        run.add_picture(tmp, width=Cm(5), height=Cm(4.5))
    except:
        pass
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)

def fill_template_gis(data_file, output_path, template_path):
    with open(data_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    doc = Document(template_path)

    # TABLE 0: Realizado por / Aprobado por
    t0 = doc.tables[0]
    set_cell_text(t0.rows[0].cells[1], data.get('realizado',''))

    # TABLE 1: Nombre del punto
    set_cell_text(doc.tables[1].rows[0].cells[0], data.get('nombre_punto',''))

    # TABLE 2: Nombre del cliente
    set_cell_text(doc.tables[2].rows[0].cells[0], data.get('nombre_cliente',''))

    # TABLE 3: Dirección del punto
    set_cell_text(doc.tables[3].rows[0].cells[0], data.get('direccion',''))

    # TABLE 4: LOGIN
    set_cell_text(doc.tables[4].rows[0].cells[0], data.get('login',''))

    # TABLE 5: Números de contacto (2 cells) — uses cell[0]
    set_cell_text(doc.tables[5].rows[0].cells[0], data.get('contacto',''))

    # TABLE 6: Empresa | (second col empty)
    set_cell_text(doc.tables[6].rows[0].cells[0], data.get('empresa',''))

    # TABLE 7: Técnico Responsable | Zona Asignada
    set_cell_text(doc.tables[7].rows[0].cells[0], data.get('tecnico',''))
    if len(doc.tables[7].rows[0].cells) > 1:
        set_cell_text(doc.tables[7].rows[0].cells[1], data.get('zona',''))

    # TABLE 8: Descripción del trabajo (cell[1])
    if len(doc.tables[8].rows[0].cells) > 1:
        set_cell_text(doc.tables[8].rows[0].cells[1], data.get('descripcion',''))

    # TABLE 9: Nodo inicio/final/ruta/caja/buffer/hilos/codigo
    t9 = doc.tables[9]
    if len(t9.rows) > 1:
        row = t9.rows[1]
        vals = ['nodo_inicio','nodo_final','ruta','caja','buffer','hilos','codigo_caja']
        for i, key in enumerate(vals):
            if i < len(row.cells):
                set_cell_text(row.cells[i], data.get(key,''))

    # TABLES 10-20: DESCRIPCION FOTOGRAFICA
    # Table 10: single cell - foto 1
    # Table 11: 2 cells - labels (Fig1, Fig2)
    # Table 12: 2 cells - foto2, foto3
    # Table 13: 2 cells - labels
    # Table 14: 2 cells - foto4, foto5
    # Table 15: single - label
    # Table 16: 2 cells - foto6, foto7
    # Table 17: 2 cells - labels
    # Table 18: 2 cells - foto8, foto9
    # Table 19: 2 cells - labels
    # Table 20: single - foto10
    fotos = data.get('fotos', [])

    # Save ALL table references BEFORE any deletions
    all_tables = list(doc.tables)

    # Photo tables in template: 10=single, 11=labels, 12=pair, 13=labels, 14=pair,
    # 15=label, 16=pair, 17=labels, 18=pair, 19=labels, 20=single
    # Tables with actual photo cells: 10, 12, 14, 16, 18, 20
    # Tables with only labels (11,13,15,17,19) stay as-is
    # Strategy: fill slots that have photos, DELETE entire table groups that are empty

    fotos_con_img = [f for f in fotos if f.get('img') and f['img'] != '']
    n_fotos = len(fotos_con_img)

    # Map: (photo_table_idx, col) for each slot
    foto_slots = [
        (10, 0),    # slot 0: table 10, col 0
        (12, 0),    # slot 1: table 12, col 0
        (12, 1),    # slot 2: table 12, col 1
        (14, 0),    # slot 3
        (14, 1),    # slot 4
        (16, 0),    # slot 5
        (16, 1),    # slot 6
        (18, 0),    # slot 7
        (18, 1),    # slot 8
        (20, 0),    # slot 9
    ]

    # Fill slots that have photos
    for idx, (t_i, col) in enumerate(foto_slots):
        if t_i < len(all_tables):
            cell = all_tables[t_i].rows[0].cells[col]
            if idx < n_fotos:
                add_image_to_cell(cell, fotos_con_img[idx]['img'])

    # Delete empty photo table groups (photo table + its label table)
    # Groups: (label_tbl, photo_tbl, min_foto_idx_needed)
    groups = [
        (11, 12, 1),   # group 1: needs foto 1 or 2
        (13, 14, 3),   # group 2: needs foto 3 or 4
        (15, 16, 5),   # group 3: needs foto 5 or 6
        (17, 18, 7),   # group 4: needs foto 7 or 8
        (19, 20, 9),   # group 5: needs foto 9
    ]

    # Also handle table 10 (single, no label table before it)
    if n_fotos == 0 and 10 < len(all_tables):
        tbl = all_tables[10]._tbl
        tbl.getparent().remove(tbl)

    # Remove groups in reverse order to preserve indices
    body = doc.element.body
    for (lbl_t, img_t, min_idx) in reversed(groups):
        if n_fotos <= min_idx:
            # Remove both label table and image table
            for t_i in [img_t, lbl_t]:
                if t_i < len(all_tables):
                    tbl = all_tables[t_i]._tbl
                    parent = tbl.getparent()
                    if parent is not None:
                        parent.remove(tbl)

    # For photos beyond 10, add them after the last photo table
    if n_fotos > 10:
        from docx.oxml import OxmlElement as _OE
        from docx.oxml.ns import qn as _qn
        # Find last photo table and add rows
        remaining = fotos_con_img[10:]
        # Add a new table after the document body for extras
        i = 0
        while i < len(remaining):
            tr = _OE('w:tr')
            for col in range(2):
                tc = _OE('w:tc')
                tcp = _OE('w:tcPr')
                tcw = _OE('w:tcW'); tcw.set(_qn('w:w'),'4627'); tcw.set(_qn('w:type'),'dxa')
                tcp.append(tcw); tc.append(tcp)
                p = _OE('w:p'); tc.append(p)
                tr.append(tc)
            # Append to last available photo table
            last_photo_tbl = None
            for t in doc.tables:
                last_photo_tbl = t
            if last_photo_tbl:
                last_photo_tbl._tbl.append(tr)
                tcs = tr.findall(_qn('w:tc'))
                for col_idx in range(min(2, len(remaining)-i)):
                    from docx.table import _Cell
                    cell = _Cell(tcs[col_idx], last_photo_tbl)
                    if remaining[i+col_idx].get('img'):
                        add_image_to_cell(cell, remaining[i+col_idx]['img'])
            i += 2

    # TABLE 21 and 22 use original indices from all_tables
    # TABLE 21: Lista de materiales (rows 1-21, cols: 0=material(fixed), 1=cantidad, 2=medida, 3=metros, 4=total)
    t21 = all_tables[21]
    materiales = data.get('materiales', [])
    material_keys = ['fibra','canaleta','manguera_anillada','manguera_plastica','tubo_conduit',
                     'grapas_plasticas','grapas_metalicas','tacos','alambre','amarras',
                     'caja_bmx','caja_multimedia','pachtcord_fibra','patchcord_utp',
                     'duplex','simplex','tubos_fusion','sliders','abrazaderas','broca','minimanga']
    for i, key in enumerate(material_keys):
        row_idx = i + 1
        if row_idx < len(t21.rows):
            mat = next((m for m in materiales if m.get('key') == key), {})
            set_cell_text(t21.rows[row_idx].cells[1], mat.get('cantidad',''))
            set_cell_text(t21.rows[row_idx].cells[2], mat.get('medida',''))
            set_cell_text(t21.rows[row_idx].cells[3], mat.get('metros',''))
            set_cell_text(t21.rows[row_idx].cells[4], mat.get('total',''))

    # Materiales Adicionales — find paragraph and add text after it
    from docx.oxml.ns import qn as _qn
    body = doc.element.body
    children = list(body)
    for idx, child in enumerate(children):
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            from docx.text.paragraph import Paragraph
            p = Paragraph(child, doc)
            if 'Materiales Adicionales' in p.text:
                # Find next table after this paragraph
                for j in range(idx+1, min(idx+5, len(children))):
                    if children[j].tag.split('}')[-1] == 'tbl':
                        from docx.table import Table
                        t_adicional = Table(children[j], doc)
                        if t_adicional.rows:
                            set_cell_text(t_adicional.rows[0].cells[0], data.get('mat_adicionales',''))
                        break
                break

    # TABLE 22: Conclusiones / Tiempo estimado
    t22 = all_tables[22]
    if len(t22.rows) > 1:
        set_cell_text(t22.rows[1].cells[0], data.get('conclusiones',''))
        if len(t22.rows[1].cells) > 1:
            set_cell_text(t22.rows[1].cells[1], data.get('tiempo',''))

    doc.save(output_path)
    print('SAVED')

if __name__ == '__main__':
    fill_template_gis(sys.argv[1], sys.argv[2], sys.argv[3])
