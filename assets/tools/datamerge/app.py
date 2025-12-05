import customtkinter as ctk
from tkinter import filedialog, messagebox
import logging
from merge_logic import merge_files, load_excel_clean

logging.basicConfig(level=logging.INFO, format="%(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

class DataMergeApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("DataMerge - Fusion Excel")
        self.geometry("800x600")

        ctk.set_appearance_mode("system")  # dark/light auto
        ctk.set_default_color_theme("blue")

        self.file1_path = None
        self.file2_path = None

        self.build_ui()

    def build_ui(self):

        # ✅ Conteneur centré
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(expand=True)

        # ✅ Frame principale
        frame = ctk.CTkFrame(container, corner_radius=15)
        frame.pack(padx=20, pady=20)

        # ✅ Titre
        title = ctk.CTkLabel(frame, text="DataMerge - Fusion Excel",
                             font=("Segoe UI", 26, "bold"))
        title.grid(row=0, column=0, columnspan=4, pady=(10, 30))

        # ✅ Fichier 1
        ctk.CTkLabel(frame, text="Fichier source (avec données)").grid(row=1, column=0, sticky="w")
        ctk.CTkButton(frame, text="Choisir fichier 1", command=self.load_file1).grid(row=1, column=1, padx=10)
        ctk.CTkButton(frame, text="Aperçu", command=self.preview_file1).grid(row=1, column=2)
        self.file1_label = ctk.CTkLabel(frame, text="Aucun fichier", text_color="#888")
        self.file1_label.grid(row=1, column=3, padx=10)

        # ✅ Fichier 2
        ctk.CTkLabel(frame, text="Fichier cible (à compléter)").grid(row=2, column=0, sticky="w", pady=(15, 0))
        ctk.CTkButton(frame, text="Choisir fichier 2", command=self.load_file2).grid(row=2, column=1, padx=10)
        ctk.CTkButton(frame, text="Aperçu", command=self.preview_file2).grid(row=2, column=2)
        self.file2_label = ctk.CTkLabel(frame, text="Aucun fichier", text_color="#888")
        self.file2_label.grid(row=2, column=3, padx=10)

        # ✅ Sélecteurs de colonnes (alignement + espacement uniforme)
        row_start = 3

        ctk.CTkLabel(frame, text="Clé fichier 1").grid(row=row_start, column=0, sticky="w", pady=(30, 5))
        self.key1_combo = ctk.CTkComboBox(frame, values=[], variable=ctk.StringVar(value=""))
        self.key1_combo.grid(row=row_start, column=1, pady=(30, 5))

        ctk.CTkLabel(frame, text="Clé fichier 2").grid(row=row_start+1, column=0, sticky="w", pady=5)
        self.key2_combo = ctk.CTkComboBox(frame, values=[], variable=ctk.StringVar(value=""))
        self.key2_combo.grid(row=row_start+1, column=1, pady=5)

        ctk.CTkLabel(frame, text="Colonne source (fichier 1)").grid(row=row_start+2, column=0, sticky="w", pady=5)
        self.source_combo = ctk.CTkComboBox(frame, values=[], variable=ctk.StringVar(value=""))
        self.source_combo.grid(row=row_start+2, column=1, pady=5)

        ctk.CTkLabel(frame, text="Colonne destination (fichier 2)").grid(row=row_start+3, column=0, sticky="w", pady=5)
        self.target_combo = ctk.CTkComboBox(frame, values=[], variable=ctk.StringVar(value=""))
        self.target_combo.grid(row=row_start+3, column=1, pady=5)

        # ✅ Bouton fusion
        ctk.CTkButton(frame, text="Fusionner", height=45, command=self.run_merge).grid(
            row=row_start+4, column=0, columnspan=4, pady=40
        )

    # ---------------------------------------------------------
    # Chargement fichiers
    # ---------------------------------------------------------
    def load_file1(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return

        self.file1_path = path
        df = load_excel_clean(path)
        cols = df.columns.tolist()

        self.key1_combo.configure(values=cols)
        self.source_combo.configure(values=cols)

        self.file1_label.configure(text=path.split("/")[-1])

        messagebox.showinfo("Fichier 1 chargé", f"Colonnes détectées :\n{cols}")

    def load_file2(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return

        self.file2_path = path
        df = load_excel_clean(path)
        cols = df.columns.tolist()

        self.key2_combo.configure(values=cols)
        self.target_combo.configure(values=cols)

        self.file2_label.configure(text=path.split("/")[-1])

        messagebox.showinfo("Fichier 2 chargé", f"Colonnes détectées :\n{cols}")

    # ---------------------------------------------------------
    # Aperçu fichiers
    # ---------------------------------------------------------
    def preview_file(self, path, title):
        if not path:
            messagebox.showerror("Erreur", "Aucun fichier chargé.")
            return

        df = load_excel_clean(path)
        preview = df.head(20).to_string()

        win = ctk.CTkToplevel(self)
        win.title(title)

        text = ctk.CTkTextbox(win, width=650, height=450)
        text.pack(padx=20, pady=20)
        text.insert("0.0", preview)
        text.configure(state="disabled")

    def preview_file1(self):
        self.preview_file(self.file1_path, "Aperçu fichier 1")

    def preview_file2(self):
        self.preview_file(self.file2_path, "Aperçu fichier 2")

    # ---------------------------------------------------------
    # Fusion
    # ---------------------------------------------------------
    def run_merge(self):
        if not self.file1_path or not self.file2_path:
            messagebox.showerror("Erreur", "Veuillez sélectionner les deux fichiers.")
            return

        try:
            merge_files(
                self.file1_path,
                self.file2_path,
                self.key1_combo.get(),
                self.key2_combo.get(),
                self.source_combo.get(),
                self.target_combo.get(),
                "output/merged.xlsx"
            )

            messagebox.showinfo("Succès", "Fusion terminée ! Fichier généré : output/merged.xlsx")

        except Exception as e:
            logger.error("Erreur lors de la fusion.", exc_info=True)
            messagebox.showerror("Erreur", str(e))


if __name__ == "__main__":
    app = DataMergeApp()
    app.mainloop()
