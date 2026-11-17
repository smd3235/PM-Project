bas = r"C:\Users\smd32\Downloads\PM Project\scripts\sync_template.py"
content = open(bas).read()

# Add format_email_template_sheet calls before save
old = '''    wb.Save()
    wb.Close(True)
    print(f"Synced {len(bas_files)} module(s) and button '{button_name}' into {template}")'''

new = '''    # Format email template sheets with distinct colors
    for sheet_name, accent, accent_light in [("Manager Email Template", 0x8B4513, 0xE8F0FE), ("Director Email Template", 0x006400, 0xE8F5E9)]:
        try:
            ws = wb.Worksheets(sheet_name)
            ws.Unprotect()
            ws.Cells.Interior.Color = 0xFFFFFF
            ws.Cells.Borders.LineStyle = -4142
            ws.Range("B2:K2").Interior.Color = accent
            ws.Range("B2").Font.Color = 0xFFFFFF
            ws.Range("B3:K3").Interior.Color = accent_light
            ws.Range("B4:K4").Interior.Color = 0xFFF2CC
            for lr in [5, 7, 9, 11]:
                ws.Range("C" + str(lr)).Interior.Color = accent_light
                ws.Range("C" + str(lr)).Font.Color = accent
                ws.Range("C" + str(lr)).Font.Bold = True
            for vr in [5, 7, 9]:
                rng = ws.Range("D" + str(vr) + ":K" + str(vr))
                rng.Interior.Color = 0xFFFFFF
                rng.Borders.LineStyle = 1
                rng.Borders.Color = accent
                rng.Borders.Weight = 2
            rng = ws.Range("D11:K31")
            rng.Interior.Color = 0xFFFFFF
            rng.Borders.LineStyle = 1
            rng.Borders.Color = accent
            rng.Borders.Weight = 2
            ws.Protect()
            ws.EnableSelection = 2
            print(f"Formatted {sheet_name} with accent {hex(accent)}")
        except Exception as e:
            print(f"Format {sheet_name} failed: {e}")

    wb.Save()
    wb.Close(True)
    print(f"Synced {len(bas_files)} module(s) and button '{button_name}' into {template}")'''

assert old in content, "old block not found"
content = content.replace(old, new, 1)

open(bas, "w").write(content)
print("Updated")
