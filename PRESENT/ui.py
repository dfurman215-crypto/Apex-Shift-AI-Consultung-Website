import json
import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from builder import build_deck


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_SPEC = BASE_DIR / "specs" / "pearl.json"
DEFAULT_OUTPUT = BASE_DIR / "output" / "PEARL_PRESENT_MVE.pptx"


class PresentUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PRESENT MVE")
        self.geometry("760x330")
        self.minsize(700, 300)

        self.spec_var = tk.StringVar(value=str(DEFAULT_SPEC))
        self.output_var = tk.StringVar(value=str(DEFAULT_OUTPUT))
        self.status_var = tk.StringVar(value="Ready")

        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self, padding=18)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text="PRESENT", font=("Segoe UI", 20, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        ttk.Label(
            root,
            text="Minimal interface for building PowerPoint decks from PRESENT specifications.",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 18))

        ttk.Label(root, text="Deck specification").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(root, textvariable=self.spec_var).grid(
            row=2, column=1, sticky="ew", padx=(12, 8), pady=6
        )
        ttk.Button(root, text="Browse", command=self._browse_spec).grid(row=2, column=2, pady=6)

        ttk.Label(root, text="Output PowerPoint").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(root, textvariable=self.output_var).grid(
            row=3, column=1, sticky="ew", padx=(12, 8), pady=6
        )
        ttk.Button(root, text="Browse", command=self._browse_output).grid(row=3, column=2, pady=6)

        actions = ttk.Frame(root)
        actions.grid(row=4, column=0, columnspan=3, sticky="w", pady=(22, 14))

        ttk.Button(actions, text="Build PowerPoint", command=self._build_deck).pack(side="left")
        ttk.Button(actions, text="Open Output", command=self._open_output).pack(side="left", padx=8)
        ttk.Button(actions, text="Open Output Folder", command=self._open_output_folder).pack(side="left")

        ttk.Separator(root).grid(row=5, column=0, columnspan=3, sticky="ew", pady=(4, 12))
        ttk.Label(root, text="Status:").grid(row=6, column=0, sticky="w")
        ttk.Label(root, textvariable=self.status_var).grid(row=6, column=1, columnspan=2, sticky="w")

        root.columnconfigure(1, weight=1)

    def _browse_spec(self):
        path = filedialog.askopenfilename(
            title="Select PRESENT deck specification",
            initialdir=str(BASE_DIR / "specs"),
            filetypes=[("JSON specification", "*.json"), ("All files", "*.*")],
        )
        if path:
            self.spec_var.set(path)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save PowerPoint",
            initialdir=str(BASE_DIR / "output"),
            initialfile="PEARL_PRESENT_MVE.pptx",
            defaultextension=".pptx",
            filetypes=[("PowerPoint presentation", "*.pptx")],
        )
        if path:
            self.output_var.set(path)

    def _build_deck(self):
        spec_path = Path(self.spec_var.get()).expanduser()
        output_path = Path(self.output_var.get()).expanduser()

        if not spec_path.exists():
            messagebox.showerror("PRESENT", f"Specification not found:\n{spec_path}")
            return

        try:
            self.status_var.set("Building...")
            self.update_idletasks()

            with spec_path.open("r", encoding="utf-8") as handle:
                spec = json.load(handle)

            build_deck(spec, str(output_path))
            self.status_var.set(f"Built successfully: {output_path.name}")
            messagebox.showinfo("PRESENT", f"PowerPoint created successfully.\n\n{output_path}")
        except Exception as exc:
            self.status_var.set("Build failed")
            messagebox.showerror("PRESENT", f"Build failed:\n\n{exc}")

    def _open_output(self):
        path = Path(self.output_var.get()).expanduser()
        if not path.exists():
            messagebox.showwarning("PRESENT", "The output PowerPoint does not exist yet.")
            return
        self._open_path(path)

    def _open_output_folder(self):
        path = Path(self.output_var.get()).expanduser().parent
        path.mkdir(parents=True, exist_ok=True)
        self._open_path(path)

    @staticmethod
    def _open_path(path: Path):
        if sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)


if __name__ == "__main__":
    app = PresentUI()
    app.mainloop()
