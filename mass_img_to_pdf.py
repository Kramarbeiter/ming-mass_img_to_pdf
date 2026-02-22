import os
import json
import sys
import ctypes
import zipfile
import io
import tkinter as tk
from tkinter import filedialog, messagebox, Menu
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image
from fpdf import FPDF  # WICHTIG: 'fpdf2' muss installiert sein


def resource_path(relative_path):
    """Gets the absolute path to the resource, compatible with PyInstaller."""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)


class ToolTip:
    """Creates a small hover window (tooltip) for GUI elements."""

    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(500, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(
            tw,
            text=self.text,
            justify=tk.LEFT,
            background="#ffffe0",
            relief=tk.SOLID,
            borderwidth=1,
            font=("tahoma", 8, "normal"),
        )
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()


class PDFConverter:
    """Core logic for extracting ZIPs and converting images to PDFs."""

    def __init__(
        self, input_folder: str, output_folder: str, delete_source: bool = False
    ):
        self.input_folder = input_folder
        self.output_folder = output_folder
        self.delete_source = delete_source

        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)

    def _extract_zip_files(self):
        print(f"\n🔍 Suche nach ZIP-Dateien in: {self.input_folder}")
        zip_count = 0
        for root, _, files in os.walk(self.input_folder):
            for file in files:
                if file.lower().endswith(".zip"):
                    zip_path = os.path.join(root, file)
                    try:
                        with zipfile.ZipFile(zip_path, "r") as zip_ref:
                            zip_ref.extractall(root)

                        if self.delete_source:
                            os.remove(zip_path)
                            print(f"📦 Entpackt und gelöscht: {file}")
                        else:
                            print(f"📦 Entpackt (behalten): {file}")
                        zip_count += 1
                    except Exception as e:
                        print(f"⚠️ Fehler beim Entpacken von {zip_path}: {e}")

    def convert_images_to_pdf(self):
        self._extract_zip_files()
        folders_with_images = []
        for root, _, files in os.walk(self.input_folder):
            if any(
                f.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp"))
                for f in files
            ):
                folders_with_images.append(root)

        total_folders = len(folders_with_images)
        if total_folders == 0:
            return 0  # Keine Bilder verarbeitet

        for folder in folders_with_images:
            relative_path = os.path.relpath(folder, self.input_folder)

            # Falls der Hauptordner selbst Bilder enthält, benennen wir das PDF nach dem Ordnernamen
            if relative_path == ".":
                safe_pdf_name = os.path.basename(os.path.normpath(self.input_folder))
            else:
                safe_pdf_name = relative_path.replace(os.sep, "_")

            self._create_pdf_from_images(folder, safe_pdf_name)

        return total_folders

    def _create_pdf_from_images(self, subfolder_path: str, pdf_name: str):
        pdf = FPDF()
        image_extensions = (".png", ".jpg", ".jpeg", ".gif", ".bmp")
        image_files = sorted(
            [
                f
                for f in os.listdir(subfolder_path)
                if f.lower().endswith(image_extensions)
            ]
        )
        processed_images = []

        margin = 10

        for image_file in image_files:
            image_path = os.path.join(subfolder_path, image_file)
            try:
                image = Image.open(image_path)
                if image.mode != "RGB":
                    image = image.convert("RGB")

                img_w, img_h = image.size

                # Dynamische Seitenausrichtung
                if img_w > img_h:
                    orientation = "L"  # Landscape
                else:
                    orientation = "P"  # Portrait

                # Seite mit der exakt passenden Ausrichtung hinzufügen
                pdf.add_page(orientation=orientation)

                # Maximalen Platz berechnen (passt sich der Ausrichtung automatisch an)
                max_w = pdf.w - 2 * margin
                max_h = pdf.h - 2 * margin

                # Skalierung und Zentrierung berechnen
                ratio = min(max_w / img_w, max_h / img_h)
                new_w = img_w * ratio
                new_h = img_h * ratio

                x_pos = (pdf.w - new_w) / 2
                y_pos = (pdf.h - new_h) / 2

                img_byte_arr = io.BytesIO()
                image.save(img_byte_arr, format="JPEG")

                pdf.image(img_byte_arr, x=x_pos, y=y_pos, w=new_w, h=new_h)
                processed_images.append(image_path)
            except Exception as e:
                print(f"⚠️ Fehler bei {image_file}: {e}")

        # PDF speichern
        pdf_output_path = os.path.join(self.output_folder, f"{pdf_name}.pdf")
        try:
            pdf.output(pdf_output_path)
        except Exception as e:
            print(f"⚠️ Fehler beim Speichern der PDF '{pdf_name}': {e}")
            return

        # Sichere Lösch-Logik
        if self.delete_source:
            for img_path in processed_images:
                try:
                    os.remove(img_path)
                except Exception:
                    pass
            if subfolder_path != self.input_folder:
                try:
                    os.rmdir(subfolder_path)
                except OSError:
                    pass


class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Image to PDF Converter")
        self.root.geometry("500x450")
        self.root.resizable(False, False)

        try:
            myappid = "img2pdf.converter.1.0"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

        icon_path = resource_path("icon.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.folder_paths = []
        self.config_path = os.path.join(
            os.path.expanduser("~"), ".img2pdf_converter_cfg.json"
        )

        self.out_dir_var = tk.StringVar(value=os.path.join(os.getcwd(), "pdf_output"))
        self.delete_source_var = tk.BooleanVar(value=False)

        self.load_settings()

        # --- Top Frame ---
        top_frame = tk.Frame(root)
        top_frame.pack(fill=tk.X, padx=20, pady=(15, 5))

        self.add_btn = tk.Button(
            top_frame, text="Add Folder", command=self.add_folders, width=10
        )
        self.add_btn.pack(side=tk.LEFT)
        ToolTip(self.add_btn, "Select one or more input folders from your PC.")

        tk.Label(
            top_frame, text="Selected Input Folders:", font=("Arial", 9, "bold")
        ).pack(side=tk.LEFT, padx=20)

        # --- Listbox Frame ---
        list_frame = tk.Frame(root)
        list_frame.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            yscrollcommand=scrollbar.set,
            font=("Arial", 10),
        )
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)

        self.listbox.drop_target_register(DND_FILES)  # type: ignore
        self.listbox.dnd_bind("<<Drop>>", self.drop_folders)  # type: ignore

        self.listbox.bind("<Delete>", self.remove_selected)
        ToolTip(
            self.listbox,
            "Drag & Drop your Input Folders here!\nMultiple selection enabled.\nPress 'Del' or right-click to remove folders.",
        )

        self.context_menu = Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Delete", command=self.remove_selected)
        self.listbox.bind("<Button-3>", self.show_context_menu)

        # --- Settings Frame ---
        settings_frame = tk.Frame(root)
        settings_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(settings_frame, text="Output Folder:").grid(
            row=0, column=0, sticky="w", pady=5
        )
        folder_frame = tk.Frame(settings_frame)
        folder_frame.grid(row=0, column=1, sticky="w", padx=5)

        self.out_dir_entry = tk.Entry(
            folder_frame, textvariable=self.out_dir_var, width=30
        )
        self.out_dir_entry.pack(side=tk.LEFT)
        ToolTip(
            self.out_dir_entry,
            "The folder where the converted .pdf files will be saved.",
        )

        self.browse_btn = tk.Button(
            folder_frame, text="...", command=self.choose_folder, width=3
        )
        self.browse_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(self.browse_btn, "Browse for output folder.")

        self.chk_delete = tk.Checkbutton(
            settings_frame,
            text="Delete source files (ZIPs/Images) after conversion",
            variable=self.delete_source_var,
        )
        self.chk_delete.grid(row=1, column=0, columnspan=2, sticky="w", pady=2)
        ToolTip(
            self.chk_delete,
            "If checked, successfully processed images and ZIP archives\nwill be permanently deleted from the input folders.",
        )

        # --- Convert Button ---
        self.convert_btn = tk.Button(
            root,
            text="Convert to PDF",
            command=self.process_folders,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=5,
        )
        self.convert_btn.pack(pady=(5, 15))
        ToolTip(
            self.convert_btn, "Starts processing all folders currently in the list."
        )

    def load_settings(self):
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r") as f:
                    config = json.load(f)
                    if "out_dir" in config:
                        self.out_dir_var.set(config["out_dir"])
                    if "delete_source" in config:
                        self.delete_source_var.set(config["delete_source"])
        except Exception:
            pass

    def save_settings(self):
        try:
            config = {
                "out_dir": self.out_dir_var.get(),
                "delete_source": self.delete_source_var.get(),
            }
            with open(self.config_path, "w") as f:
                json.dump(config, f)
        except Exception:
            pass

    def on_closing(self):
        self.save_settings()
        self.root.destroy()

    def add_folders(self):
        folder = filedialog.askdirectory(title="Select Input Folder")
        if folder and folder not in self.folder_paths:
            self.folder_paths.append(folder)
            self.listbox.insert(tk.END, folder)

    def drop_folders(self, event):
        files = self.root.tk.splitlist(event.data)
        for f in files:
            # Check if dropped item is a directory
            if os.path.isdir(f) and f not in self.folder_paths:
                self.folder_paths.append(f)
                self.listbox.insert(tk.END, f)

    def remove_selected(self, event=None):
        selection = self.listbox.curselection()
        if not selection:
            return
        for i in reversed(selection):
            self.listbox.delete(i)
            del self.folder_paths[i]

    def show_context_menu(self, event):
        try:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(self.listbox.nearest(event.y))
            self.listbox.activate(self.listbox.nearest(event.y))
            self.context_menu.tk_popup(event.x_root, event.y_root)
        except Exception:
            pass
        finally:
            self.context_menu.grab_release()

    def choose_folder(self):
        folder = filedialog.askdirectory(
            title="Select Output Folder", initialdir=self.out_dir_var.get()
        )
        if folder:
            self.out_dir_var.set(folder)

    def process_folders(self):
        if not self.folder_paths:
            messagebox.showwarning(
                "No Folders", "Please add at least one input folder to the list."
            )
            return

        out_dir = self.out_dir_var.get()
        if not os.path.exists(out_dir):
            try:
                os.makedirs(out_dir)
            except Exception as e:
                messagebox.showerror(
                    "Error", f"Could not create output directory:\n{e}"
                )
                return

        self.save_settings()
        delete_src = self.delete_source_var.get()

        total_processed_pdfs = 0
        total_folders = len(self.folder_paths)

        # Disable button during processing to prevent double clicks
        self.convert_btn.config(state=tk.DISABLED, text="Processing...")
        self.root.update()

        try:
            for folder_path in self.folder_paths:
                converter = PDFConverter(
                    input_folder=folder_path,
                    output_folder=out_dir,
                    delete_source=delete_src,
                )
                pdfs_created = converter.convert_images_to_pdf()
                total_processed_pdfs += pdfs_created

            messagebox.showinfo(
                "Success",
                f"Done!\n\nScanned {total_folders} main folder(s).\nCreated {total_processed_pdfs} .pdf files in:\n{out_dir}",
            )
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")
        finally:
            # Restore button
            self.convert_btn.config(state=tk.NORMAL, text="Convert to PDF")


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFConverterApp(root)
    root.mainloop()
