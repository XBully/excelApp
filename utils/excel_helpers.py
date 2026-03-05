import pandas as pd
import openpyxl
import xlrd
from io import BytesIO
import base64
from copy import copy as obj_copy

def to_xlsx_stream(file):
    try:
        file.seek(0)
        df = pd.read_excel(file, header=None, engine='xlrd')
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=False)
        output.seek(0)
        return output
    except: return None

def clean_columns(df):
    if df is None: return None
    new_cols, seen = [], {}
    
    for i, col in enumerate(df.columns):
        parts = []
        if isinstance(col, tuple):
            for p in col:
                p_str = str(p).strip()
                if p_str and "nan" not in p_str.lower() and "unnamed" not in p_str.lower():
                    parts.append(p_str)
        else:
            p_str = str(col).strip()
            if p_str and "nan" not in p_str.lower() and "unnamed" not in p_str.lower():
                parts = [p_str]
        
        name = " - ".join(parts) if parts else f"未命名_{i}"
        
        if name in seen:
            seen[name] += 1
            final_name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
            final_name = name
            
        new_cols.append(final_name)
    
    df.columns = new_cols
    df.dropna(how='all', inplace=True)
    return df

def load_excel(file, start_row, row_count):
    try:
        file.seek(0)
        if file.name.lower().endswith('.xls'):
            df = pd.read_excel(file, header=list(range(start_row, start_row + row_count)), engine='xlrd')
        else:
            df = pd.read_excel(file, header=list(range(start_row, start_row + row_count)), engine='openpyxl')
        return clean_columns(df)
    except Exception as e:
        return None

def get_headers_only(file, start_row, row_count):
    try:
        file.seek(0)
        if file.name.lower().endswith('.xls'):
            df = pd.read_excel(file, header=list(range(start_row, start_row + row_count)), nrows=0, engine='xlrd')
        else:
            df = pd.read_excel(file, header=list(range(start_row, start_row + row_count)), nrows=0, engine='openpyxl')
        return clean_columns(df).columns.tolist()
    except: return []

def find_col_index(target, header_list):
    if not target: return -1
    try:
        return header_list.index(target) + 1
    except:
        clean_t = target.rsplit('_', 1)[0] if '_' in target else target
        for i, h in enumerate(header_list):
            if h and clean_t in str(h): return i + 1
    return -1

def queue_download_js(results):
    import json
    files_list = []
    for name, data in results:
        b64 = base64.b64encode(data).decode()
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        if name.lower().endswith('.xls'):
            mime = "application/vnd.ms-excel"
        files_list.append({"name": name, "base64": b64, "mime": mime})
    
    files_json = json.dumps(files_list)
    
    js_code = f"""
        <script>
        (function() {{
            const files = {files_json};
            async function download() {{
                for (const file of files) {{
                    const blob = await (await fetch(`data:${{file.mime}};base64,${{file.base64}}`)).blob();
                    const url = window.URL.createObjectURL(blob);
                    const link = document.createElement('a');
                    link.style.display = 'none';
                    link.href = url;
                    link.download = file.name;
                    document.body.appendChild(link);
                    link.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(link);
                    await new Promise(r => setTimeout(r, 800));
                }}
            }}
            download();
        }})();
        </script>
    """
    return js_code
