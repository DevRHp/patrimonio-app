import base64
import os

# Paths
input_path = r"c:\Users\Administrador\Desktop\patrimonio\codes\new_logo_b64.txt"
output_path = r"c:\Users\Administrador\Desktop\patrimonio\backend\static\logo.png"

try:
    with open(input_path, "r") as f:
        b64_data = f.read().strip()
        
    # Remove header if present (e.g., "data:image/png;base64,")
    if "," in b64_data:
        b64_data = b64_data.split(",")[1]

    img_data = base64.b64decode(b64_data)

    with open(output_path, "wb") as f:
        f.write(img_data)
    
    print(f"Successfully saved logo to {output_path}")

except Exception as e:
    print(f"Error processing logo: {e}")
