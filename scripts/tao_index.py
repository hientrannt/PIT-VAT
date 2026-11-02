"""Tạo web/index.html (trang độc lập để đưa lên Vercel / Cloudflare Pages) từ web/thue_hkd.html.

Chạy lại mỗi khi sửa web/thue_hkd.html:  python scripts/tao_index.py
"""
from pathlib import Path

goc = Path(__file__).resolve().parent.parent / "web"
noi_dung = (goc / "thue_hkd.html").read_text(encoding="utf-8")
khung = (
    "<!doctype html>\n<html lang=\"vi\">\n<head>\n<meta charset=\"utf-8\">\n"
    "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1,viewport-fit=cover\">\n"
    "<meta name=\"robots\" content=\"noindex\">\n"
    "<style>:root{color-scheme:light}body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>\n"
    "</head>\n<body>\n"
)
(goc / "index.html").write_text(khung + noi_dung + "\n</body>\n</html>\n", encoding="utf-8")
print("Đã tạo", goc / "index.html")
