import json, subprocess, time
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent
CONFIG = BASE / 'config.json'


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}")


def load_config():
    return json.loads(CONFIG.read_text(encoding='utf-8'))


def extract_data(excel_path, sheet_name):
    excel_path = Path(excel_path)
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f'No existe la hoja: {sheet_name}')
    ws = wb[sheet_name]

    header_row = None
    headers = None
    for row in ws.iter_rows(values_only=True):
        vals = [str(v).strip() if v is not None else '' for v in row]
        if 'BASIS' in vals and 'SAP' in vals and 'DESCRIPCIÓN' in vals and 'Familia' in vals and 'Status' in vals:
            header_row = vals
            headers = {h: i for i, h in enumerate(vals) if h}
            break
    if headers is None:
        raise ValueError('No se encontró la fila de encabezados de Ingreso Inv. Sap')

    def col(*names):
        for name in names:
            if name in headers:
                return headers[name]
        raise ValueError('No se encontró ninguna de las columnas: ' + ', '.join(names))

    i_basis = col('BASIS')
    i_sap = col('SAP')
    i_desc = col('DESCRIPCIÓN')
    i_stock = col('STOCK PT', 'STOCK')
    i_familia = col('Familia')
    i_status = col('Status')

    rows = []
    families = set()
    seen = set()
    data_started = False
    for row in ws.iter_rows(values_only=True):
        vals = list(row)
        if not data_started:
            normalized = [str(v).strip() if v is not None else '' for v in vals]
            if normalized == header_row:
                data_started = True
            continue
        basis = vals[i_basis] if i_basis < len(vals) else None
        sap = vals[i_sap] if i_sap < len(vals) else None
        desc = vals[i_desc] if i_desc < len(vals) else None
        stock = vals[i_stock] if i_stock < len(vals) else None
        familia = vals[i_familia] if i_familia < len(vals) else None
        status = vals[i_status] if i_status < len(vals) else None
        if desc in (None, '', 0, '#DIV/0!'):
            continue
        if sap in (None, '', 0, '#DIV/0!') or basis in (None, '', 0, '#DIV/0!'):
            continue
        if isinstance(stock, str) and stock.startswith('#'):
            stock = None
        item = {
            'BASIS': basis,
            'SAP': sap,
            'DESCRIPCIÓN': desc,
            'STOCK PT': stock,
            'Familia': str(familia).strip() if familia is not None else '',
            'Status': str(status).strip() if status is not None else ''
        }
        key = (str(basis), str(sap), str(desc), str(item['Familia']), str(item['Status']))
        if key in seen:
            continue
        seen.add(key)
        rows.append(item)
        if item['Familia']:
            families.add(item['Familia'])
    wb.close()

    rows.sort(key=lambda r: (str(r['Familia']).lower(), str(r['DESCRIPCIÓN']).lower()))
    return {
        'updated_at': datetime.now().isoformat(timespec='seconds'),
        'source': excel_path.name,
        'rows': rows,
        'options': {'Familia': sorted(families, key=str.lower)},
        'filters': {'Status': 'En circulación', 'Familia': 'Todos'}
    }


def git_push(commit_message):
    subprocess.run(['git', 'add', 'data.json'], cwd=BASE, check=True)
    result = subprocess.run(['git', 'diff', '--cached', '--quiet'], cwd=BASE)
    if result.returncode == 0:
        return False
    subprocess.run(['git', 'commit', '-m', commit_message], cwd=BASE, check=True)
    subprocess.run(['git', 'push'], cwd=BASE, check=True)
    return True


def update_once(cfg):
    excel = BASE / cfg['excel_filename']
    if not excel.exists():
        raise FileNotFoundError(f'No se encontró el Excel: {excel}')
    payload = extract_data(excel, cfg.get('sheet_name', 'Ingreso Inv. Sap'))
    out = BASE / 'data.json'
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    log(f"Datos actualizados: {len(payload['rows'])} registros desde {cfg['excel_filename']}")
    if cfg.get('auto_push', True):
        if git_push(cfg.get('commit_message', 'Actualizar stock')):
            log('Cambios publicados en GitHub.')
        else:
            log('Sin cambios nuevos para publicar.')


def main():
    cfg = load_config()
    print('STOCK CDA CAMANA - ACTUALIZADOR AUTOMÁTICO')
    print(f"Excel: {BASE / cfg['excel_filename']}")
    print(f"Hoja fuente: {cfg.get('sheet_name', 'Ingreso Inv. Sap')}")
    print(f"GitHub: https://github.com/{cfg['github_username']}/{cfg['github_repository']}.git")
    print(f"Revisión cada {cfg.get('poll_seconds', 5)} segundos.")
    last_mtime = None
    while True:
        try:
            excel = BASE / cfg['excel_filename']
            mtime = excel.stat().st_mtime
            if last_mtime is None or mtime != last_mtime:
                update_once(cfg)
                last_mtime = mtime
        except Exception as e:
            log(f'ERROR: {e}')
        time.sleep(cfg.get('poll_seconds', 5))

if __name__ == '__main__':
    main()
