import os

files_to_check = [
    'mps_tab_debug.txt',
    'mps_tab_structure.txt',
    'mps_structure.txt',
    'mps_out.txt',
    'mps_debug_output.txt'
]

for filename in files_to_check:
    path = os.path.join(r'C:\Users\i0215099\Desktop\MPS_UPDATE', filename)
    if os.path.exists(path):
        print(f"--- File: {filename} ---")
        try:
            # Try different encodings
            for enc in ['utf-8', 'utf-16', 'utf-16-le', 'cp949']:
                try:
                    with open(path, 'r', encoding=enc) as f:
                        content = f.read()
                        if content.strip():
                            print(f"Content (Encoding: {enc}):")
                            print(content[:2000])  # type: ignore # First 2000 chars
                            break
                except:
                    continue
        except Exception as e:
            print(f"Error reading {filename}: {e}")
    else:
        print(f"File not found: {filename}")
