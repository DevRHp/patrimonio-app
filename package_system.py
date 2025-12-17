import os
import shutil
import sys

SOURCE_DIR = os.path.dirname(os.path.abspath(__file__))
DEST_DIR = os.path.join(SOURCE_DIR, 'sistema_patrimonio_entregavel')

FILES_TO_COPY = [
    ('backend/app.py', 'backend/app.py'),
    ('backend/drive_manager.py', 'backend/drive_manager.py'),
    ('backend/database.db', 'backend/database.db'), # Include DB with data
    ('requirements.txt', 'requirements.txt'),
    ('Procfile', 'Procfile'),
    ('.gitignore', '.gitignore'),
    ('GUIA_DE_USO.txt', 'GUIA_DE_USO.txt'),
]

FOLDERS_TO_COPY = [
    ('backend/templates', 'backend/templates'),
    ('backend/static', 'backend/static'),
]

EMPTY_FOLDERS = [
    'uploads',
    'uploads/scanned_data',
    'Relatorios_Gerados',
]

def copy_files():
    if not os.path.exists(DEST_DIR):
        os.makedirs(DEST_DIR)
        print(f"Created {DEST_DIR}")
    else:
        print(f"Updating {DEST_DIR}...")

    # Copy individual files
    for src, dst in FILES_TO_COPY:
        src_path = os.path.join(SOURCE_DIR, src)
        dst_path = os.path.join(DEST_DIR, dst)
        
        if os.path.exists(src_path):
            os.makedirs(os.path.dirname(dst_path), exist_ok=True)
            shutil.copy2(src_path, dst_path)
            print(f"Copied {src} -> {dst}")
        else:
            print(f"WARNING: Source file {src} not found. Skipped.")

    # Copy folders
    for src, dst in FOLDERS_TO_COPY:
        src_path = os.path.join(SOURCE_DIR, src)
        dst_path = os.path.join(DEST_DIR, dst)
        
        if os.path.exists(src_path):
            shutil.copytree(src_path, dst_path, dirs_exist_ok=True)
            print(f"Copied folder {src} -> {dst}")
        else:
            print(f"WARNING: Source folder {src} not found. Skipped.")

    # Create empty folders
    for folder in EMPTY_FOLDERS:
        path = os.path.join(DEST_DIR, folder)
        os.makedirs(path, exist_ok=True)
        print(f"Created empty folder: {folder}")

    # Create README
    readme_content = """Sistema de Auditoria Patrimonial
================================

Instruções de Instalação:

1. Instale o Python (versão 3.9 ou superior).
2. Abra o terminal nesta pasta.
3. Instale as dependências:
   pip install -r requirements.txt

4. Para rodar o sistema:
   python backend/app.py

5. Abra o navegador em: http://127.0.0.1:5000

Notas:
- Os arquivos de upload ficarão na pasta 'uploads'.
- Os relatórios gerados ficarão na pasta 'Relatorios_Gerados'.
- O banco de dados está em 'backend/database.db'.
- Para ativar a integração com Google Drive, coloque o arquivo 'credentials.json' na pasta 'backend/'.
"""
    with open(os.path.join(DEST_DIR, 'README.txt'), 'w', encoding='utf-8') as f:
        f.write(readme_content)
    
    # Create Run Script (Windows)
    bat_content = """@echo off
echo Iniciando Sistema de Patrimonio...
pip install -r requirements.txt
python backend/app.py
pause
"""
    with open(os.path.join(DEST_DIR, 'iniciar_sistema.bat'), 'w', encoding='utf-8') as f:
        f.write(bat_content)

    print("\nPackaging Complete!")

if __name__ == '__main__':
    copy_files()
