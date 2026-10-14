"""A4 HTML preview built from the exact exported workbook, without remote assets."""
from html import escape
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries, get_column_letter


def workbook_html(data, name='Анкета клиента'):
    wb = load_workbook(BytesIO(data), data_only=True)
    pages = []
    for ws in wb:
        area = str(ws.print_area).split('!')[-1].replace("'", '').split(',')[0].replace('$', '')
        try:
            c1, r1, c2, r2 = range_boundaries(area)
        except ValueError:
            c1, r1, c2, r2 = 1, 1, ws.max_column, ws.max_row
        merges = {(m.min_row, m.min_col): m for m in ws.merged_cells.ranges}
        skip = {(r, c) for m in ws.merged_cells.ranges for r in range(m.min_row, m.max_row+1) for c in range(m.min_col, m.max_col+1) if (r,c)!=(m.min_row,m.min_col)}
        total_height = sum(ws.row_dimensions[r].height or 15 for r in range(r1,r2+1))
        # Fit exactly the worksheet's A4 print region on a page.
        scale = min(1, 650/max(total_height,1))
        widths=[ws.column_dimensions[get_column_letter(c)].width or 10 for c in range(c1,c2+1)]
        cols=''.join(f'<col style="width:{w/sum(widths)*100}%">' for w in widths)
        rows = []
        for r in range(r1,r2+1):
            cells = []
            for c in range(c1,c2+1):
                if (r,c) in skip: continue
                x=ws.cell(r,c);m=merges.get((r,c));spans=f' colspan="{m.max_col-m.min_col+1}" rowspan="{m.max_row-m.min_row+1}"' if m else ''
                value=x.value
                if isinstance(value,(int,float)) and x.number_format!='@': value=f'{value:,.0f}'.replace(',', ' ')
                color = '#'+x.fill.fgColor.rgb[-6:] if x.fill.fgColor.type=='rgb' and x.fill.fgColor.rgb and x.fill.fgColor.rgb[:2]!='00' else 'white'
                edges=[]
                for side in ['top','bottom','left','right']:
                    if getattr(x.border,side) is not None and getattr(x.border,side).style: edges.append(f'border-{side}:1px solid #999')
                style=f'font-size:{(x.font.sz or 9)*scale}pt;font-weight:{700 if x.font.b else 400};text-align:{x.alignment.horizontal or "left"};background:{color};'+ ';'.join(edges)
                cells.append(f'<td{spans} style="{style}">{escape(str(value if value is not None else ""))}</td>')
            rows.append(f'<tr style="height:{(ws.row_dimensions[r].height or 15)*scale}pt">'+''.join(cells)+'</tr>')
        pages.append('<section class="page"><div class="brand">BEREKELY EREKE · BRE Calc — '+escape(ws.title)+'</div><table><colgroup>'+cols+'</colgroup>'+''.join(rows)+'</table></section>')
    return '<!doctype html><html lang="ru"><meta charset="utf-8"><title>'+escape(name or "Анкета клиента")+'</title><style>@page{size:A4 portrait;margin:8mm}*{box-sizing:border-box}body{margin:0;background:#ddd;font-family:Arial,sans-serif}.page{width:194mm;min-height:281mm;background:white;margin:12px auto;padding:0;break-after:page}.brand{font-size:8pt;color:#7623b6;height:14pt}table{border-collapse:collapse;width:100%;table-layout:fixed}td{padding:1px 3px;vertical-align:middle;overflow-wrap:anywhere;white-space:pre-wrap} @media print{body{background:white}.page{margin:0;min-height:0}.page:last-child{break-after:auto}}</style>'+''.join(pages)+'</html>'
