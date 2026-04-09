from __future__ import annotations

import json
import shutil
import sys
import time
from pathlib import Path

from generate_plan import generate_from_file


BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "drive_config.json"
SUPPORTED_EXTENSIONS = {".txt", ".docx", ".pdf"}


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"No se encontro {CONFIG_PATH.name}. Crea una copia de drive_config.example.json y renombral a drive_config.json."
        )
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def ensure_directories(config: dict) -> dict[str, Path]:
    directories = {
        "watch_folder": Path(config["watch_folder"]),
        "processed_folder": Path(config["processed_folder"]),
        "output_folder": Path(config["output_folder"]),
        "error_folder": Path(config["error_folder"]),
    }
    for path in directories.values():
        path.mkdir(parents=True, exist_ok=True)
    return directories


def stable_files(folder: Path) -> list[Path]:
    files = [path for path in folder.iterdir() if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS]
    ready = []
    for path in files:
        size_now = path.stat().st_size
        time.sleep(0.5)
        size_after = path.stat().st_size
        if size_now == size_after:
            ready.append(path)
    return sorted(ready, key=lambda item: item.stat().st_mtime)


def move_with_suffix(source: Path, target_dir: Path) -> Path:
    target = target_dir / source.name
    counter = 1
    while target.exists():
        target = target_dir / f"{source.stem}_{counter}{source.suffix}"
        counter += 1
    shutil.move(str(source), str(target))
    return target


def copy_output_files(outputs: list[Path], output_dir: Path) -> list[Path]:
    copied: list[Path] = []
    for file_path in outputs:
        target = output_dir / file_path.name
        counter = 1
        while target.exists():
            target = output_dir / f"{file_path.stem}_{counter}{file_path.suffix}"
            counter += 1
        shutil.copy2(file_path, target)
        summary = file_path.with_suffix(".txt")
        if summary.exists():
            summary_target = output_dir / summary.name
            counter = 1
            while summary_target.exists():
                summary_target = output_dir / f"{summary.stem}_{counter}{summary.suffix}"
                counter += 1
            shutil.copy2(summary, summary_target)
        copied.append(target)
    return copied


def process_one(file_path: Path, config: dict, directories: dict[str, Path]) -> None:
    selected_mode = None if config.get("default_mode", "auto") == "auto" else config["default_mode"]
    try:
        outputs, data, detected = generate_from_file(file_path, mode=selected_mode, prompt_on_combined=False)
        copied = copy_output_files(outputs, directories["output_folder"])
        archived = move_with_suffix(file_path, directories["processed_folder"])
        print(f"OK | {archived.name} | modo={detected}")
        for item in copied:
            print(f"  salida: {item}")
        if data.pending:
            print("  pendientes:")
            for pending in data.pending:
                print(f"  - {pending}")
    except Exception as exc:
        failed = move_with_suffix(file_path, directories["error_folder"])
        print(f"ERROR | {failed.name} | {exc}")


def main() -> int:
    config = load_config()
    directories = ensure_directories(config)
    poll_seconds = max(5, int(config.get("poll_seconds", 15)))

    print("Modo Drive iniciado.")
    print(f"Entrada: {directories['watch_folder']}")
    print(f"Salida: {directories['output_folder']}")
    print("Presiona Ctrl+C para detener.")

    try:
        while True:
            files = stable_files(directories["watch_folder"])
            for file_path in files:
                process_one(file_path, config, directories)
            time.sleep(poll_seconds)
    except KeyboardInterrupt:
        print("\nProceso detenido por el usuario.")
        return 0


if __name__ == "__main__":
    sys.exit(main())
