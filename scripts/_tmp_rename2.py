doc = r"C:\Users\smd32\Downloads\PM Project\README.md"
content = open(doc).read()
content = content.replace('Template (No Macros).xlsx', 'PM_Requests_Master_Template_(No Macros).xlsx')
open(doc, "w").write(content)
print("Updated README.md")
