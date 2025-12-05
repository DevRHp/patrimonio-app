import os
import zipfile
import io
import time
from flask import Flask, render_template, request, send_file, jsonify, session
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from copy import copy

app = Flask(__name__)
app.secret_key = 'super_secret_key_sesi_sorocaba' # Change this in production!

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
REPORTS_FOLDER = os.path.join(UPLOAD_FOLDER, 'reports')
os.makedirs(REPORTS_FOLDER, exist_ok=True)

# --- Routes ---

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/login', methods=['POST'])
def login():
    data = request.json
    email = data.get('email')
    password = data.get('password')

    if email == 'admin@123' and password == 'admin123':
        session['is_admin'] = True
        return jsonify({'message': 'Login realizado com sucesso', 'success': True})
    else:
        return jsonify({'error': 'Credenciais inválidas', 'success': False}), 401

@app.route('/logout', methods=['POST'])
def logout():
    session.pop('is_admin', None)
    return jsonify({'message': 'Logout realizado com sucesso'})

@app.route('/check_auth', methods=['GET'])
def check_auth():
    return jsonify({'is_admin': session.get('is_admin', False)})

# --- File Management ---

@app.route('/upload_master', methods=['POST'])
def upload_master():
    if not session.get('is_admin'):
        return jsonify({'error': 'Acesso negado. Requer privilégios de administrador.'}), 403

    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file and file.filename.endswith('.xlsx'):
        filename = file.filename
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        return jsonify({'message': f'Planilha "{filename}" carregada com sucesso!'})
    
    return jsonify({'error': 'Formato de arquivo inválido. Apenas .xlsx'}), 400

@app.route('/list_masters', methods=['GET'])
def list_masters():
    masters = []
    if os.path.exists(app.config['UPLOAD_FOLDER']):
        for f in os.listdir(app.config['UPLOAD_FOLDER']):
            if f.endswith('.xlsx') and not f.startswith('~$') and f != 'master_spreadsheet.xlsx':
                 # Filtering out old default name if desired, or keep it.
                 # Also avoid temp excel files (~$)
                 masters.append(f)
    return jsonify({'masters': masters})

@app.route('/delete_master', methods=['POST'])
def delete_master():
    if not session.get('is_admin'):
        return jsonify({'error': 'Acesso negado. Requer privilégios de administrador.'}), 403

    data = request.json
    filename = data.get('filename')
    
    if not filename:
         return jsonify({'error': 'Nome do arquivo não fornecido'}), 400

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    if os.path.exists(file_path):
        try:
            os.remove(file_path)
            return jsonify({'message': f'Planilha "{filename}" removida com sucesso'})
        except Exception as e:
            return jsonify({'error': f'Erro ao remover: {str(e)}'}), 500
    else:
        return jsonify({'error': 'Planilha não encontrada'}), 404

# --- Data Fetching ---

@app.route('/get_rooms', methods=['POST'])
def get_rooms():
    # Accepts JSON: { "filenames": ["file1.xlsx", "file2.xlsx"] }
    data = request.json
    selected_files = data.get('filenames', [])
    
    if not selected_files:
        return jsonify({'rooms': []})

    all_rooms = []
    
    for filename in selected_files:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if not os.path.exists(file_path):
            continue
            
        try:
            wb = load_workbook(file_path, read_only=True, data_only=True)
            
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                room_display_name = sheet_name 
                
                # Search for "Denominação"
                found_header = False
                for row in ws.iter_rows(min_row=1, max_row=20, max_col=20):
                    for cell in row:
                        if cell.value and str(cell.value).strip() == "Denominação":
                            target_row = cell.row + 1
                            target_col = cell.column
                            try:
                                val = ws.cell(row=target_row, column=target_col).value
                                if val:
                                    room_display_name = str(val).strip()
                                    found_header = True
                            except:
                                pass
                            break
                    if found_header:
                        break
                
                all_rooms.append({
                    'id': sheet_name, 
                    'name': room_display_name,
                    'source': filename
                })
            wb.close()
        except Exception as e:
            print(f"Error reading {filename}: {e}")

    return jsonify({'rooms': all_rooms})

# --- Reports ---

@app.route('/list_reports', methods=['GET'])
def list_reports():
    try:
        reports = []
        if os.path.exists(REPORTS_FOLDER):
            for filename in os.listdir(REPORTS_FOLDER):
                if filename.endswith('.zip'):
                    filepath = os.path.join(REPORTS_FOLDER, filename)
                    stats = os.stat(filepath)
                    reports.append({
                        'filename': filename,
                        'created_at': stats.st_mtime,
                        'size': stats.st_size
                    })
        reports.sort(key=lambda x: x['created_at'], reverse=True)
        return jsonify({'reports': reports})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/delete_report', methods=['POST'])
def delete_report():
    if not session.get('is_admin'):
        return jsonify({'error': 'Acesso negado.'}), 403
        
    data = request.json
    filename = data.get('filename')
    if not filename: return jsonify({'error': 'Filename required'}), 400
    
    filepath = os.path.join(REPORTS_FOLDER, filename)
    if os.path.exists(filepath):
        os.remove(filepath)
        return jsonify({'message': 'Relatório removido'})
    return jsonify({'error': 'Relatório não encontrado'}), 404

@app.route('/get_report/<path:filename>', methods=['GET'])
def get_report(filename):
    try:
        return send_file(os.path.join(REPORTS_FOLDER, filename), as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 404

# --- Core Logic: Verify ---

@app.route('/verify', methods=['POST'])
def verify():
    data = request.json
    analyst_name = data.get('analyst_name', 'Analista')
    selected_room = data.get('room_name')
    source_file = data.get('source_file')
    selected_files = data.get('selected_files', [])
    scanned_codes_raw = data.get('scanned_codes', '')
    
    if not source_file:
         return jsonify({'error': 'Arquivo fonte da sala não identificado'}), 400

    scanned_codes = set(code.strip() for code in scanned_codes_raw.splitlines() if code.strip())

    if not selected_room:
        return jsonify({'error': 'Nenhuma sala selecionada'}), 400
        
    source_path = os.path.join(app.config['UPLOAD_FOLDER'], source_file)
    if not os.path.exists(source_path):
        return jsonify({'error': f'Arquivo fonte "{source_file}" não encontrado'}), 404

    try:
        # Load the source workbook (Target Room)
        wb = load_workbook(source_path)
        
        if selected_room not in wb.sheetnames:
             return jsonify({'error': f'Sala "{selected_room}" não encontrada no arquivo "{source_file}"'}), 400

        source_ws = wb[selected_room]
        green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
        
        # Identify found items in the target room
        found_in_room_codes = set()
        
        # We need to scan the sheet first to find matches
        for row in source_ws.iter_rows(min_row=2, values_only=True):
            row_values = [str(cell).strip() for cell in row if cell is not None]
            for val in row_values:
                if val in scanned_codes:
                    found_in_room_codes.add(val)
                    break

        # --- 1. Verified Items (Sheet 1: Verificados) ---
        verified_ws = wb.copy_worksheet(source_ws)
        verified_ws.title = "Verificados"
        
        for row in verified_ws.iter_rows(min_row=2):
            match_found = False
            for cell in row:
                if cell.value is not None and str(cell.value).strip() in scanned_codes:
                    match_found = True
                    break
            
            if match_found:
                for cell in row:
                    cell.fill = green_fill

        # --- 3. Missing Items (Sheet 2: Nao Encontrados) ---
        missing_ws = wb.create_sheet("Nao Encontrados")
        # Copy column dimensions
        for col_name, col_dim in source_ws.column_dimensions.items():
            missing_ws.column_dimensions[col_name].width = col_dim.width

        # Copy Header
        for row in source_ws.iter_rows(min_row=1, max_row=1):
            missing_ws.append([cell.value for cell in row])
            for i, cell in enumerate(row):
                new_cell = missing_ws.cell(row=1, column=i+1)
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = copy(cell.number_format)
                    new_cell.protection = copy(cell.protection)
                    new_cell.alignment = copy(cell.alignment)

        # Copy Missing Rows
        current_row_idx = 2
        for row in source_ws.iter_rows(min_row=2):
            is_found = False
            for cell in row:
                if cell.value is not None and str(cell.value).strip() in found_in_room_codes:
                    is_found = True
                    break
            
            if not is_found:
                for i, cell in enumerate(row):
                    new_cell = missing_ws.cell(row=current_row_idx, column=i+1, value=cell.value)
                    if cell.has_style:
                        new_cell.font = copy(cell.font)
                        new_cell.border = copy(cell.border)
                        new_cell.fill = copy(cell.fill)
                        new_cell.number_format = copy(cell.number_format)
                        new_cell.protection = copy(cell.protection)
                        new_cell.alignment = copy(cell.alignment)
                current_row_idx += 1


        # --- 2. Wrong Location Items (Sheet 3: Local Incorreto) ---
        wrong_location_ws = wb.create_sheet("Local Incorreto")
        wrong_location_ws.append(["Codigo", "Encontrado na Sala", "Deveria estar em", "Descricao do Item"]) 
        # Style header
        if source_ws.cell(1,1).has_style:
             for cell in wrong_location_ws[1]:
                cell.font = copy(source_ws.cell(1,1).font)

        scanned_but_not_in_room = scanned_codes - found_in_room_codes
        
        # Cross-referencing Logic
        files_to_search = selected_files if selected_files else [source_file]
        
        for code in scanned_but_not_in_room:
            found_location = "Nao encontrado nas planilhas selecionadas"
            item_description = "Desconhecido"
            found_in_cross_ref = False
            
            for fname in files_to_search:
                fpath = os.path.join(app.config['UPLOAD_FOLDER'], fname)
                if not os.path.exists(fpath): continue
                
                try:
                    # Optimization: open read-only
                    wb_search = load_workbook(fpath, read_only=True, data_only=True)
                    for sheet_name in wb_search.sheetnames:
                        if fname == source_file and sheet_name == selected_room: continue
                        
                        sheet = wb_search[sheet_name]
                        for row in sheet.iter_rows(values_only=True):
                            row_str_values = [str(v).strip() for v in row if v is not None]
                            if code in row_str_values:
                                found_location = f"{sheet_name} ({fname})"
                                # Heuristic for description: 2nd column
                                if len(row) > 1:
                                    item_description = str(row[1]) 
                                found_in_cross_ref = True
                                break
                        if found_in_cross_ref: break
                    wb_search.close()
                except:
                    continue
                if found_in_cross_ref: break
            
            wrong_location_ws.append([code, selected_room, found_location, item_description])

        # --- Save and Zip ---
        memory_file = io.BytesIO()
        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
            
            temp_buffer = io.BytesIO()
            wb.save(temp_buffer)
            temp_buffer.seek(0)
            
            # File 1: Verified
            wb_v = load_workbook(temp_buffer)
            for s in wb_v.sheetnames:
                if s != "Verificados": del wb_v[s]
            with io.BytesIO() as f_out:
                wb_v.save(f_out)
                zf.writestr(f"{analyst_name}_Verificados.xlsx", f_out.getvalue())
            
            # File 2: Missing
            temp_buffer.seek(0)
            wb_m = load_workbook(temp_buffer)
            for s in wb_m.sheetnames:
                if s != "Nao Encontrados": del wb_m[s]
            with io.BytesIO() as f_out:
                wb_m.save(f_out)
                zf.writestr(f"{analyst_name}_Nao_Encontrados.xlsx", f_out.getvalue())

            # File 3: Wrong Location
            temp_buffer.seek(0)
            wb_w = load_workbook(temp_buffer)
            for s in wb_w.sheetnames:
                if s != "Local Incorreto": del wb_w[s]
            with io.BytesIO() as f_out:
                wb_w.save(f_out)
                zf.writestr(f"{analyst_name}_Local_Incorreto.xlsx", f_out.getvalue())

        memory_file.seek(0)
        
        # Save unique report by overwriting based on Analyst + Room
        safe_room = "".join([c for c in selected_room if c.isalnum() or c in (' ','-','_')]).strip()
        report_filename = f"{analyst_name}_{safe_room}_Analise.zip"
        report_path = os.path.join(REPORTS_FOLDER, report_filename)
        
        with open(report_path, 'wb') as f:
            f.write(memory_file.getvalue())

        memory_file.seek(0)
        
        return send_file(
            memory_file,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'{analyst_name}_{safe_room}_Resultados.zip'
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
