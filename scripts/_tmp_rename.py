bas = r"C:\Users\smd32\Downloads\PM Project\scripts\build-and-package.ps1"
content = open(bas).read()
content = content.replace('Template (No Macros).xlsx', 'PM_Requests_Master_Template_(No Macros).xlsx')
open(bas, "w").write(content)
print("Updated build-and-package.ps1")

# Update BUILD-INSTRUCTIONS.txt
doc = r"C:\Users\smd32\Downloads\PM Project\BUILD-INSTRUCTIONS.txt"
content2 = open(doc).read()
content2 = content2.replace('Template (No Macros).xlsx', 'PM_Requests_Master_Template_(No Macros).xlsx')
open(doc, "w").write(content2)
print("Updated BUILD-INSTRUCTIONS.txt")
