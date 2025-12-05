from PIL import Image, ImageEnhance, ImageFilter
import base64
import io

def process_logo():
    try:
        # Load the original image
        # Load the original image
        img = Image.open("images (1).jpg")
        
        # Upscale the image to improve visual quality on UI
        base_width = 500
        w_percent = (base_width / float(img.size[0]))
        h_size = int((float(img.size[1]) * float(w_percent)))
        img = img.resize((base_width, h_size), Image.Resampling.LANCZOS)
        
        img = img.convert("RGBA")
        
        # Process data to make white transparent
        datas = img.getdata()
        new_data = []
        for item in datas:
            # Change all white (also shades of whites) pixels to transparent
            if item[0] > 200 and item[1] > 200 and item[2] > 200:
                new_data.append((255, 255, 255, 0))
            else:
                new_data.append(item)
        
        img.putdata(new_data)
        
        # Enhance sharpness
        enhancer = ImageEnhance.Sharpness(img)
        img = enhancer.enhance(2.0) # Increase sharpness
        
        # Enhance contrast
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(1.2)
        
        # Save to buffer as PNG to keep transparency
        buffered = io.BytesIO()
        img.save(buffered, format="PNG")
        
        # Encode to base64
        img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
        
        # Save to file
        with open("new_logo_b64.txt", "w") as f:
            f.write(img_str)
            
        print("Logo processed successfully.")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    process_logo()
