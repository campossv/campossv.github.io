import html
import re
import os

try:
    base_dir = r'c:\Users\vcampos\OneDrive - Devel Security, S.A\Documentos\GitHub\campossv.github.io'
    script_path = os.path.join(base_dir, 'MisScripts', 'HealthCheck_utf8.ps1')
    html_path = os.path.join(base_dir, 'HTML', 'HealthCheck.html')

    with open(script_path, 'r', encoding='utf-8') as f:
        script_content = f.read()
    
    escaped_script_content = html.escape(script_content)

    with open(html_path, 'r', encoding='utf-8') as f:
        html_content = f.read()

    start_tag = '<pre><code>'
    end_tag = '</code></pre>'
    
    pattern = re.compile(f"({re.escape(start_tag)})(.*?)({re.escape(end_tag)})", re.DOTALL)

    def replace_func(match):
        return match.group(1) + escaped_script_content + match.group(3)

    new_html_content, num_replacements = pattern.subn(replace_func, html_content, count=1)

    if num_replacements > 0:
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(new_html_content)
        print(f"Successfully updated '{os.path.basename(html_path)}'.")
    else:
        print("Error: Could not find the <pre><code> block to replace in the HTML file.")

except Exception as e:
    print(f"An error occurred: {e}")
