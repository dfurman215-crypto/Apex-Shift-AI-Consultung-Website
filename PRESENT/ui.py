import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from amend import amend_deck


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT = BASE_DIR / "output" / "PEARL_PRESENT_MVE.pptx"


class PresentUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PRESENT MVE")
        self.geometry("850x400")
        self.minsize(790, 360)

        self.source_var = tk.StringVar(value="")
        self.output_var = tk.StringVar(value=str(DEFAULT_OUTPUT))
        self.use_agent_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Choose your existing PowerPoint to begin.")

        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self, padding=18)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text="PRESENT", font=("Segoe UI", 20, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        ttk.Label(
            root,
            text="Use an existing PowerPoint as the source deck and build from its embedded Markdown instructions.",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 18))

        ttk.Label(root, text="Source PowerPoint").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(root, textvariable=self.source_var).grid(
            row=2, column=1, sticky="ew", padx=(12, 8), pady=6
        )
        ttk.Button(root, text="Browse", command=self._browse_source).grid(row=2, column=2, pady=6)

        ttk.Label(root, text="Output PowerPoint").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(root, textvariable=self.output_var).grid(
            row=3, column=1, sticky="ew", padx=(12, 8), pady=6
        )
        ttk.Button(root, text="Browse", command=self._browse_output).grid(row=3, column=2, pady=6)

        ttk.Checkbutton(
            root,
            text="Use local Gemma planner through Ollama (falls back to deterministic parser if unavailable)",
            variable=self.use_agent_var,
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(12, 4))

        actions = ttk.Frame(root)
        actions.grid(row=5, column=0, columnspan=3, sticky="w", pady=(18, 14))

        ttk.Button(actions, text="Build From Existing Deck", command=self._build_from_source).pack(side="left")
        ttk.Button(actions, text="Open Output", command=self._open_output).pack(side="left", padx=8)
        ttk.Button(actions, text="Open Output Folder", command=self._open_output_folder).pack(side="left")

        ttk.Separator(root).grid(row=6, column=0, columnspan=3, sticky="ew", pady=(4, 12))
        ttk.Label(root, text="Status:").grid(row=7, column=0, sticky="nw")
        ttk.Label(root, textvariable=self.status_var, wraplength=640).grid(
            row=7, column=1, columnspan=2, sticky="w"
        )

        root.columnconfigure(1, weight=1)

    def _browse_source(self):
        path = filedialog.askopenfilename(
            title="Select existing PowerPoint",
            filetypes=[("PowerPoint presentation", "*.pptx"), ("All files", "*.*")],
        )
        if path:
            self.source_var.set(path)
            source = Path(path)
            self.output_var.set(str(BASE_DIR / "output" / f"{source.stem}_PRESENT.pptx"))
            self.status_var.set("Source deck selected. PRESENT will preserve it and append slides from the embedded Markdown brief.")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save amended PowerPoint",
            initialdir=str(BASE_DIR / "output"),
            initialfile="PEARL_PRESENT_MVE.pptx",
            defaultextension=".pptx",
            filetypes=[("PowerPoint presentation", "*.pptx")],
        )
        if path:
            self.output_var.set(path)

    def _build_from_source(self):
        source_path = Path(self.source_var.get()).expanduser()
        output_path = Path(self.output_var.get()).expanduser()

        if not source_path.exists():
            messagebox.showerror("PRESENT", "Choose an existing PowerPoint first.")
            return

        if source_path.resolve() == output_path.resolve():
            messagebox.showerror("PRESENT", "Output must be a different file so the source deck is not overwritten.")
            return

        try:
            if self.use_agent_var.get():
                self.status_var.set("Reading source deck and asking local Gemma to plan the new slides...")
            else:
                self.status_var.set("Reading source deck and parsing embedded Markdown instructions...")
            self.update_idletasks()

            result = amend_deck(
                str(source_path),
                str(output_path),
                use_agent=self.use_agent_var.get(),
            )

            planner_note = f"Planner: {result['planner']}."
            if result.get("agent_error"):
                planner_note += " Local Gemma was unavailable or returned an invalid plan, so PRESENT used its deterministic fallback."

            self.status_var.set(
                f"Built successfully. {planner_note} Markdown found on slide {result['markdown_slide']}; "
                f"preserved {result['original_slide_count']} existing slides and added "
                f"{result['added_slide_count']} new slides."
            )
            messagebox.showinfo(
                "PRESENT",
                "PRESENT completed the amendment pass.\n\n"
                f"Planner: {result['planner']}\n"
                f"Source slides preserved: {result['original_slide_count']}\n"
                f"Slides added: {result['added_slide_count']}\n"
                f"Final slides: {result['final_slide_count']}\n\n"
                f"Output:\n{output_path}"
            )
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
