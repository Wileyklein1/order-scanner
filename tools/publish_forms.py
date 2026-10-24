"""Publish order forms to the site, encrypted with a passcode (AES-256-GCM, PBKDF2-SHA256 key).

Each source workbook is copied through Excel with the Accounts sheet removed (and any cells that looked
it up blanked), then encrypted into forms/<slug>.bin. forms/manifest.json lists them. The app's
"Download latest forms" button decrypts them on the device.

usage (Windows, Excel installed):
  python tools/publish_forms.py "<source.xlsx>=<Published Name>" [...]
  python tools/publish_forms.py --new-passcode "<source.xlsx>=<Published Name>" [...]

The passcode lives only in .forms-secret (git-ignored). A new passcode means every iPad re-enters it once.
"""
import base64, datetime, hashlib, json, os, re, secrets, shutil, sys
from pathlib import Path

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

ROOT = Path(__file__).resolve().parent.parent
SECRET = ROOT / ".forms-secret"
STAGE = ROOT / "_staging"
OUT = ROOT / "forms"
ITERATIONS = 600_000
CHECK_TEXT = b"order-scanner-forms"
WORDS = ("cedar fern otter heron maple canyon river summit meadow juniper aspen willow falcon badger "
         "lantern compass harbor beacon ember glacier prairie spruce sequoia redwood coastal tide driftwood "
         "puffin osprey bison marmot pinecone trail ridge valley boulder campfire paddle kayak lighthouse "
         "sandstone granite basalt mossy acorn hollow brook creek lagoon cove pebble quail raven sparrow "
         "thistle clover sage yarrow larkspur poppy lupine trillium hemlock alder birch cypress tamarack").split()


def b64(b):
    return base64.b64encode(b).decode()


def new_passcode():
    w = [secrets.choice(WORDS).capitalize() for _ in range(4)]
    return "-".join(w) + "-" + str(secrets.randbelow(900) + 100)


def sanitize(src, dst):
    """Copy through Excel: drop the Accounts sheet, blank cells that referenced it, drop broken names/links."""
    import win32com.client as win32
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(Path(src).resolve()), UpdateLinks=0, ReadOnly=True)
        wb.SaveAs(str(dst.resolve()), 51)
        cleared = 0
        for ws in wb.Worksheets:
            if ws.Name.lower() == "accounts":
                continue
            ur = ws.UsedRange
            for c in ur:
                f = c.Formula
                if isinstance(f, str) and "accounts!" in f.lower():
                    c.ClearContents()
                    cleared += 1
        for ws in list(wb.Worksheets):
            if ws.Name.lower() == "accounts":
                ws.Delete()
        for n in list(wb.Names):
            if "#REF!" in str(n.RefersTo) or "[" in str(n.RefersTo):
                n.Delete()
        for link in (wb.LinkSources(1) or []):            # xlExcelLinks
            try:
                wb.BreakLink(link, 1)
            except Exception:
                pass
        xl.CalculateFull()
        wb.Save()
        sheets = [ws.Name for ws in wb.Worksheets]
        wb.Close(False)
        return cleared, sheets
    finally:
        xl.Quit()


def main(argv):
    fresh = "--new-passcode" in argv
    pairs = [a for a in argv if a != "--new-passcode"]
    if not pairs:
        sys.exit(__doc__)
    if fresh or not SECRET.exists():
        SECRET.write_text(new_passcode() + "\n", encoding="utf-8")
        print("new passcode written to", SECRET)
    pw = SECRET.read_text(encoding="utf-8").strip().encode()

    old = json.loads((OUT / "manifest.json").read_text()) if (OUT / "manifest.json").exists() else None
    salt = base64.b64decode(old["kdf"]["salt"]) if old and not fresh else secrets.token_bytes(16)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITERATIONS).derive(pw)
    aes = AESGCM(key)

    shutil.rmtree(STAGE, ignore_errors=True); STAGE.mkdir()
    shutil.rmtree(OUT, ignore_errors=True); OUT.mkdir()
    civ = secrets.token_bytes(12)
    manifest = {"v": 1, "updated": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
                "kdf": {"name": "PBKDF2", "hash": "SHA-256", "iterations": ITERATIONS, "salt": b64(salt)},
                "check": {"iv": b64(civ), "ct": b64(aes.encrypt(civ, CHECK_TEXT, None))}, "forms": []}
    for pair in pairs:
        src, name = pair.split("=", 1)
        slug = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")
        staged = STAGE / f"{slug}.xlsx"
        cleared, sheets = sanitize(src, staged)
        data = staged.read_bytes()
        iv = secrets.token_bytes(12)
        (OUT / f"{slug}.bin").write_bytes(aes.encrypt(iv, data, None))
        manifest["forms"].append({"name": name, "file": f"forms/{slug}.bin", "iv": b64(iv), "bytes": len(data),
                                  "sha256": hashlib.sha256(data).hexdigest()})
        print(f"{name}: {len(data):,} bytes, sheets {sheets}, cleared {cleared} Accounts lookups")
    (OUT / "manifest.json").write_text(json.dumps(manifest, indent=1), encoding="utf-8")
    print("wrote", OUT / "manifest.json")


if __name__ == "__main__":
    main(sys.argv[1:])
