import sys
p = r"C:\Users\mcloudoc\projects\invoice-generator\static\index.html"
data = open(p, encoding="utf-8").read().lstrip("﻿")
for cp in ("cp949", "cp1252", "mbcs"):
    try:
        fixed = data.encode(cp).decode("utf-8")
    except Exception as e:
        print(cp, "FAIL", e)
        continue
    ok = ("양식" in fixed) and ("비용청구" in fixed) and ("연결" in fixed)
    print(cp, "ok?" , ok, "| sample:", repr(fixed[fixed.find("AI"):fixed.find("AI")+40]) if "AI" in fixed else "n/a")
    if ok and "--write" in sys.argv:
        open(p, "w", encoding="utf-8").write(fixed)
        print("WROTE with", cp)
        break
