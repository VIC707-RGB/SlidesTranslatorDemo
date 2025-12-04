import os
import shutil
from pathlib import Path

def delete_folder(path: Path):
    if path.exists() and path.is_dir():
        print(f"🗑️ Deleting: {path}")
        shutil.rmtree(path, ignore_errors=True)
        print("✅ Deleted\n")
    else:
        print(f"⚠️ Not found: {path}")

def find_and_delete_nllb():
    print("🔍 Searching for NLLB models...")

    # HuggingFace cache default paths
    possible_paths = [
        Path.home() / ".cache" / "huggingface" / "hub",
        Path.home() / ".cache" / "huggingface" / "transformers",
        Path.home() / "AppData" / "Local" / "huggingface" / "hub",
        Path.home() / "AppData" / "Local" / "huggingface" / "transformers",
    ]

    found_any = False

    for base in possible_paths:
        if not base.exists():
            continue

        for folder in base.iterdir():
            folder_name = folder.name.lower()
            if "nllb" in folder_name or "facebook--nllb" in folder_name:
                found_any = True
                delete_folder(folder)

    if not found_any:
        print("❌ No NLLB models found!")
    else:
        print("🎉 All NLLB models cleaned successfully!")

if __name__ == "__main__":
    find_and_delete_nllb()
