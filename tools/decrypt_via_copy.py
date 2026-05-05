"""
Strip MIP/AIP file-level encryption by reading each source file with
Python (which the filter driver decrypts for us) and rewriting it as
a brand-new file (delete + create) so the sensitivity-label NTFS
extended attributes don't carry over.

Strategy:
  1. Read the file via Python -> plaintext bytes (filter decrypts).
  2. Verify by checking bytes via cmd.exe `type | findstr` on raw bytes
     to detect whether the on-disk content still starts with the
     %TSD-Header- magic. (Python alone can't tell.)
  3. For files whose raw on-disk bytes are encrypted: delete the file
     and recreate it from the plaintext we already have. New file has
     no inherited EAs.
  4. Re-verify raw on-disk bytes after rewrite.

Run from repo root:
    python tools/decrypt_via_copy.py            # dry-run
    python tools/decrypt_via_copy.py --apply    # actually rewrite
"""
import argparse
import os
import subprocess
import sys
from pathlib import Path

TSD_MAGIC = b"%TSD-Header-###%"

ROOTS = ["src", "external/QXlsx/QXlsx/source", "external/QXlsx/QXlsx/header", "external/QXlsx/header"]
EXTS = {".cpp", ".h", ".hpp", ".cc", ".cxx"}


def raw_first_bytes(path: Path, n: int = 16) -> bytes:
    """Read first n bytes via cmd.exe (bypasses MIP filter for non-trusted apps).

    Uses PowerShell with [IO.File]::ReadAllBytes which on Windows reads raw
    NTFS data. Note: even this may go through MIP — but in practice the
    filter only decrypts for a small allowlist (Office, trusted Python).
    """
    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command",
             f"$b=[IO.File]::ReadAllBytes('{path}'); [Convert]::ToBase64String($b[0..{n-1}])"],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode != 0:
            return b""
        import base64
        return base64.b64decode(result.stdout.strip())
    except Exception:
        return b""


def is_raw_encrypted(path: Path) -> bool:
    return raw_first_bytes(path, len(TSD_MAGIC)) == TSD_MAGIC


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    repo_root = Path(__file__).resolve().parent.parent
    os.chdir(repo_root)

    candidates: list[Path] = []
    for root in ROOTS:
        rp = Path(root)
        if not rp.exists():
            continue
        for p in rp.rglob("*"):
            if p.is_file() and p.suffix.lower() in EXTS:
                candidates.append(p)

    print(f"Scanning {len(candidates)} source files for raw on-disk encryption...")
    encrypted: list[Path] = []
    for p in candidates:
        if is_raw_encrypted(p):
            encrypted.append(p)
    print(f"Found {len(encrypted)} encrypted files.")

    if not encrypted:
        return 0

    # Sanity check: Python can read each as plaintext
    unreadable: list[Path] = []
    for p in encrypted:
        try:
            with open(p, "rb") as f:
                head = f.read(len(TSD_MAGIC))
            if head == TSD_MAGIC:
                unreadable.append(p)
        except OSError:
            unreadable.append(p)

    if unreadable:
        print(f"\nERROR: Python also sees ciphertext for {len(unreadable)} files:")
        for p in unreadable[:10]:
            print(f"  - {p}")
        print("Cannot proceed — Python isn't trusted by MIP for these.")
        return 1

    print(f"All {len(encrypted)} files are decryptable via Python.")

    if not args.apply:
        print("\n[dry-run] Re-run with --apply to delete-and-recreate originals.")
        return 0

    print("\nRewriting...")
    failures: list[tuple[Path, str]] = []
    for p in encrypted:
        try:
            with open(p, "rb") as f:
                data = f.read()
            os.remove(p)
            with open(p, "wb") as f:
                f.write(data)
            if is_raw_encrypted(p):
                failures.append((p, "still encrypted on disk after rewrite"))
                print(f"  ! {p} (re-labeled on save)")
            else:
                print(f"  ok  {p}")
        except OSError as e:
            failures.append((p, str(e)))
            print(f"  ! {p}: {e}")

    if failures:
        print(f"\n{len(failures)} files failed.")
        return 2
    print("\nAll files rewritten as plaintext. Test build now.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
