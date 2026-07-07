from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

root = Path(r"C:\Users\LENOVO\Documents\Codex\2026-07-01\https-eldcm9-oss-github-io-brd")
source = root / "outputs" / "laravel-excel-backend-scaffold"
target = root / "outputs" / "laravel-excel-backend-scaffold.zip"

with ZipFile(target, "w", ZIP_DEFLATED) as zf:
    for path in source.rglob("*"):
        if path.is_file():
            if path.name == "Investment_Backend_Template.xlsx":
                continue
            zf.write(path, path.relative_to(source.parent))

print(target)
