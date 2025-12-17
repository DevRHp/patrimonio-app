import os

file_path = 'd:/patrimonio/backend/templates/index.html'

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Find the second occurrence of <!DOCTYPE html>
# The file likely starts with it, so we look for another one.
marker = '<!DOCTYPE html>'
first_index = content.find(marker)
second_index = content.find(marker, first_index + 1)

if second_index != -1:
    print(f"Found second DOCTYPE at index {second_index}. Truncating file...")
    new_content = content[second_index:]
    # Optional: dedent if the whole block is indented?
    # Let's simple write it back. Browsers handle indented HTML fine.
    
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    print("File fixed.")
else:
    print("Second DOCTYPE not found. File might be already fixed or structure is different.")
