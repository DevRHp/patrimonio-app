import re

def update_logo_in_file():
    try:
        # Read the new base64 string
        with open("new_logo_b64.txt", "r") as f:
            new_logo_b64 = f.read().strip()

        # Read the main_fe.py file
        with open("main_fe.py", "r", encoding="utf-8") as f:
            content = f.read()

        # Define the pattern to find the existing logo_base64 assignment
        # We look for self.logo_base64 = """ ... """
        # Using dotall to match newlines inside the string
        pattern = r'(self\.logo_base64\s*=\s*""")[\s\S]*?("""\s*)'
        
        # Check if pattern exists
        if not re.search(pattern, content):
            print("Error: Could not find logo_base64 pattern in main_fe.py")
            return

        # Replace with the new content
        # \1 is the start (self.logo_base64 = """), \2 is the end (""")
        new_content = re.sub(pattern, f'\\1{new_logo_b64}\\2', content)

        # Write back to file
        with open("main_fe.py", "w", encoding="utf-8") as f:
            f.write(new_content)

        print("Successfully updated main_fe.py with the new logo.")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    update_logo_in_file()
