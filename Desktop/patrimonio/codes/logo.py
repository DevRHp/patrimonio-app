import base64
import os

# Caminho da sua pasta
pasta = r"C:\Users\Administrador\Desktop\patrimonio\codes"
# Nome EXATO do seu novo arquivo PNG transparente
nome_arquivo_png = "image.png"

caminho_completo = os.path.join(pasta, nome_arquivo_png)

print(f"Procurando: {caminho_completo}")

if os.path.exists(caminho_completo):
    with open(caminho_completo, "rb") as img_file:
        # Lê o PNG
        base64_string = base64.b64encode(img_file.read()).decode('utf-8')
    
    with open("codigo_logo_png.txt", "w") as txt_file:
        txt_file.write(base64_string)
        
    print("\nSUCESSO! O código do PNG foi salvo em 'codigo_logo_png.txt'.")
    print("ABRA ESSE ARQUIVO TXT E COPIE TUDO.")
else:
    print(f"ERRO: Não achei o arquivo '{nome_arquivo_png}'. Verifique o nome e a pasta.")