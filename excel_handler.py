from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from parser_engine import ContractData

HEADERS = ["Собственник товара","тип","код ЛС льготный","№ договора","№ аукциона",
    "Международное непатентованное название (МНН)","Торговое наименование",
    "Лек. Форма (форма выпуска), дозировка","Ед. изм.","Общее кол-во, ед. изм.",
    "Производитель","Цена ед. продукции, руб. в месте поставки","Общая цена, руб.",
    "поставщик","способ размещения"]
COLUMN_WIDTHS = [15,18,15,30,25,30,25,50,10,15,45,20,20,25,25]
SHEET_NAME = "Контракты"
COL_DOSAGE=8; COL_QTY=10; COL_UNIT_PRICE=12; COL_TOTAL=13
RED_FILL=PatternFill(start_color="FF9999",end_color="FF9999",fill_type="solid")
YELLOW_FILL=PatternFill(start_color="FFFF99",end_color="FFFF99",fill_type="solid")
ORANGE_FILL=PatternFill(start_color="FFB347",end_color="FFB347",fill_type="solid")

def _owner_abbreviation(n):
    u=n.upper()
    for k,a in {"МИНЗДРАВ":"МЗ","МИНИСТЕРСТВО ЗДРАВООХРАНЕНИЯ":"МЗ"}.items():
        if k in u: return a
    return n

def _qty_value(p):
    return int(p) if p==int(p) else p

def contract_to_row(data):
    cs=data.contract_number+(f" от {data.contract_date}" if data.contract_date else "")
    qv=data.quantity_all_values if (data.quantity_mismatch and data.quantity_all_values) else _qty_value(data.quantity_packages)
    return [_owner_abbreviation(data.customer_short_name),"Основная заявка","",cs,"",data.mnn,
        data.trade_name,data.dosage_form,data.unit,qv,data.manufacturer,data.unit_price,
        data.total_price,data.supplier_short_name,data.procurement_method]

def _hdr(ws,r):
    f=PatternFill(start_color="4472C4",end_color="4472C4",fill_type="solid")
    fo=Font(bold=True,size=10,color="FFFFFF")
    b=Border(*[Side(style="thin")]*4)
    a=Alignment(horizontal="center",vertical="center",wrap_text=True)
    for c in range(1,len(HEADERS)+1):
        cell=ws.cell(row=r,column=c); cell.font=fo; cell.fill=f; cell.border=b; cell.alignment=a

def _data_style(ws,r):
    b=Border(*[Side(style="thin")]*4); a=Alignment(vertical="center",wrap_text=True)
    for c in range(1,len(HEADERS)+1):
        cell=ws.cell(row=r,column=c); cell.border=b; cell.alignment=a
        if c in (COL_UNIT_PRICE,COL_TOTAL): cell.number_format="#,##0.00"

def create_new_excel(fp):
    wb=Workbook(); ws=wb.active; ws.title=SHEET_NAME
    for c,h in enumerate(HEADERS,1): ws.cell(row=1,column=c,value=h)
    _hdr(ws,1)
    for c,w in enumerate(COLUMN_WIDTHS,1): ws.column_dimensions[get_column_letter(c)].width=w
    if not fp.endswith(".xlsx"): fp+=".xlsx"
    wb.save(fp); return fp

def write_contracts_to_excel(fp,contracts,sheet_name=None):
    path=Path(fp); ts=sheet_name or SHEET_NAME
    try:
        if path.exists() and path.suffix==".xlsx":
            wb=load_workbook(str(path))
            if ts in wb.sheetnames: ws=wb[ts]
            else:
                ws=wb.create_sheet(ts)
                for c,h in enumerate(HEADERS,1): ws.cell(row=1,column=c,value=h)
                _hdr(ws,1)
                for c,w in enumerate(COLUMN_WIDTHS,1): ws.column_dimensions[get_column_letter(c)].width=w
        else:
            wb=Workbook(); ws=wb.active; ws.title=ts
            for c,h in enumerate(HEADERS,1): ws.cell(row=1,column=c,value=h)
            _hdr(ws,1)
            for c,w in enumerate(COLUMN_WIDTHS,1): ws.column_dimensions[get_column_letter(c)].width=w
        lr=ws.max_row
        if lr==1 and ws.cell(row=1,column=1).value is None: lr=0
        w=0
        for ct in contracts:
            rn=lr+1+w
            for c,v in enumerate(contract_to_row(ct),1): ws.cell(row=rn,column=c,value=v)
            _data_style(ws,rn)
            if ct.quantity_mismatch: ws.cell(row=rn,column=COL_QTY).fill=RED_FILL
            if ct.price_mismatch:
                ws.cell(row=rn,column=COL_UNIT_PRICE).fill=RED_FILL
                ws.cell(row=rn,column=COL_TOTAL).fill=RED_FILL
            if ct.dosage_form_empty: ws.cell(row=rn,column=COL_DOSAGE).fill=ORANGE_FILL
            elif ct.dosage_form_mnn_only or ct.dosage_form_uncertain: ws.cell(row=rn,column=COL_DOSAGE).fill=YELLOW_FILL
            w+=1
        sp=str(path)
        if not sp.endswith(".xlsx"): sp=sp.rsplit(".",1)[0]+".xlsx"
        wb.save(sp); return True,f"Записано {w} строк(и) в {sp}"
    except PermissionError:
        return False,"Файл занят другой программой. Закройте Excel и повторите."
    except Exception as e:
        return False,f"Ошибка записи: {e}"

def get_sheet_names(fp):
    try:
        p=Path(fp)
        if p.suffix==".xlsx": return load_workbook(str(p),read_only=True).sheetnames
    except Exception: pass
    return []
