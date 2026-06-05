import re

with open('gui_app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Fix QMessageBox icons
content = re.sub(r'QMessageBox\.(Information|Warning|Critical)', r'QMessageBox.Icon.\1', content)

# Fix genai import
content = content.replace('import google.generativeai as genai', 'from google import genai')

# Fix genai calls
old_genai_call1 = '''genai.configure(api_key=ai_settings["api_key"])
            model = genai.GenerativeModel('gemini-1.5-flash')
            prompt = "I am a software developer. Please generate a short, typical 1-sentence description (under 50 chars) of what I might have worked on in the past hour. Do not include quotes."
            response = model.generate_content(prompt)'''

new_genai_call1 = '''client = genai.Client(api_key=ai_settings["api_key"])
            prompt = "I am a software developer. Please generate a short, typical 1-sentence description (under 50 chars) of what I might have worked on in the past hour. Do not include quotes."
            response = client.models.generate_content(
                model='gemini-1.5-flash',
                contents=prompt,
            )'''

content = content.replace(old_genai_call1, new_genai_call1)

old_genai_call2 = '''genai.configure(api_key=ai_settings["api_key"])
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    prompt = "I am a software developer. Please generate a short, typical 1-sentence description (under 50 chars) of what I might have worked on in the past hour. Do not include quotes."
                    response = model.generate_content(prompt)'''

new_genai_call2 = '''client = genai.Client(api_key=ai_settings["api_key"])
                    prompt = "I am a software developer. Please generate a short, typical 1-sentence description (under 50 chars) of what I might have worked on in the past hour. Do not include quotes."
                    response = client.models.generate_content(
                        model='gemini-1.5-flash',
                        contents=prompt,
                    )'''

content = content.replace(old_genai_call2, new_genai_call2)

with open('gui_app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Done")
