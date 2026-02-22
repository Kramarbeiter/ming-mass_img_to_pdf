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
from fpdf import FPDF


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
    """Core logic for extracting ZIPs and converting images to PDFs strictly in-memory."""

    def __init__(
        self, input_path: str, output_folder: str, delete_source: bool = False
    ):
        self.input_path = os.path.abspath(input_path)
        self.output_folder = os.path.abspath(output_folder)
        self.delete_source = delete_source

        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)

    def _get_unique_pdf_path(self, base_name: str) -> str:
        """Ensures that existing PDFs are not overwritten (appends (1), (2), etc.)."""
        pdf_name = f"{base_name}.pdf"
        pdf_path = os.path.abspath(os.path.join(self.output_folder, pdf_name))
        counter = 1

        while os.path.exists(pdf_path):
            pdf_name = f"{base_name} ({counter}).pdf"
            pdf_path = os.path.abspath(os.path.join(self.output_folder, pdf_name))
            counter += 1

        return pdf_path

    def convert(self):
        total_pdfs_created = 0

        # CASE 1: The input is directly a single ZIP file
        if os.path.isfile(self.input_path) and self.input_path.lower().endswith(".zip"):
            total_pdfs_created += self._create_pdfs_from_zip(self.input_path)
            return total_pdfs_created

        # CASE 2: The input is a folder (may contain ZIPs and/or images)
        if os.path.isdir(self.input_path):
            # 2.1 Search and process ZIP files in the folder
            for root, _, files in os.walk(self.input_path):
                for file in files:
                    if file.lower().endswith(".zip"):
                        zip_path = os.path.join(root, file)
                        total_pdfs_created += self._create_pdfs_from_zip(zip_path)

            # 2.2 Process loose images in regular folders
            folders_with_images = []
            for root, _, files in os.walk(self.input_path):
                if any(
                    f.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp"))
                    for f in files
                ):
                    folders_with_images.append(root)

            for folder in folders_with_images:
                relative_path = os.path.relpath(folder, self.input_path)

                if relative_path == ".":
                    safe_pdf_name = os.path.basename(os.path.normpath(self.input_path))
                else:
                    safe_pdf_name = relative_path.replace(os.sep, "_")

                if self._create_pdf_from_images(folder, safe_pdf_name):
                    total_pdfs_created += 1

            # 2.3 SAFE CLEANUP: Delete all empty folders (bottom-up), incl. root folder
            if self.delete_source:
                for root, dirs, files in os.walk(self.input_path, topdown=False):
                    try:
                        os.rmdir(
                            root
                        )  # Fails automatically and safely if the folder is not empty
                    except OSError:
                        pass

        return total_pdfs_created

    def _create_pdfs_from_zip(self, zip_path: str) -> int:
        """Reads images directly from the ZIP and builds a separate PDF for each virtual folder in the ZIP."""
        image_extensions = (".png", ".jpg", ".jpeg", ".gif", ".bmp")
        margin = 10
        pdfs_created_from_zip = 0

        try:
            with zipfile.ZipFile(zip_path, "r") as z:
                # Group images by folder structure WITHIN the ZIP
                images_by_folder = {}
                for f in z.namelist():
                    if f.lower().endswith(image_extensions):
                        dir_name = os.path.dirname(f)
                        if dir_name not in images_by_folder:
                            images_by_folder[dir_name] = []
                        images_by_folder[dir_name].append(f)

                if not images_by_folder:
                    return 0

                zip_base_name = os.path.splitext(os.path.basename(zip_path))[0]

                # Create a separate PDF for each subfolder in the ZIP
                for dir_name, image_files in images_by_folder.items():
                    pdf = FPDF()
                    image_files.sort()

                    if dir_name == "":
                        safe_pdf_name = zip_base_name
                    else:
                        safe_pdf_name = f"{zip_base_name}_{dir_name.replace('/', '_').replace(os.sep, '_')}"

                    processed_any = False
                    for img_name in image_files:
                        try:
                            with z.open(img_name) as img_file:
                                image = Image.open(img_file)
                                if image.mode != "RGB":
                                    image = image.convert("RGB")

                                img_w, img_h = image.size
                                orientation = "L" if img_w > img_h else "P"
                                pdf.add_page(orientation=orientation)

                                max_w = pdf.w - 2 * margin
                                max_h = pdf.h - 2 * margin

                                ratio = min(max_w / img_w, max_h / img_h)
                                new_w = img_w * ratio
                                new_h = img_h * ratio

                                x_pos = (pdf.w - new_w) / 2
                                y_pos = (pdf.h - new_h) / 2

                                img_byte_arr = io.BytesIO()
                                image.save(img_byte_arr, format="JPEG")

                                pdf.image(
                                    img_byte_arr, x=x_pos, y=y_pos, w=new_w, h=new_h
                                )
                                processed_any = True
                        except Exception as e:
                            print(
                                f"Error processing image {img_name} in ZIP {zip_path}: {e}"
                            )

                    if processed_any:
                        # Apply anti-overwrite logic here
                        pdf_output_path = self._get_unique_pdf_path(safe_pdf_name)

                        try:
                            pdf.output(pdf_output_path)
                            pdfs_created_from_zip += 1
                        except Exception as e:
                            print(f"Error saving PDF '{safe_pdf_name}': {e}")

        except Exception as e:
            print(f"Error reading ZIP {zip_path}: {e}")
            return 0

        # Delete ZIP file if requested
        if self.delete_source:
            try:
                os.remove(zip_path)
            except Exception:
                pass

        return pdfs_created_from_zip

    def _create_pdf_from_images(self, subfolder_path: str, pdf_name: str) -> bool:
        """Reads loose images from a regular folder and builds the PDF."""
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

        if not image_files:
            return False

        margin = 10

        for image_file in image_files:
            image_path = os.path.join(subfolder_path, image_file)
            try:
                image = Image.open(image_path)
                if image.mode != "RGB":
                    image = image.convert("RGB")

                img_w, img_h = image.size
                orientation = "L" if img_w > img_h else "P"
                pdf.add_page(orientation=orientation)

                max_w = pdf.w - 2 * margin
                max_h = pdf.h - 2 * margin

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
                print(f"Error processing {image_file}: {e}")

        # Apply anti-overwrite logic here
        pdf_output_path = self._get_unique_pdf_path(pdf_name)

        try:
            pdf.output(pdf_output_path)
        except Exception as e:
            print(f"Error saving PDF '{pdf_name}': {e}")
            return False

        # Delete processed images
        if self.delete_source:
            for img_path in processed_images:
                try:
                    os.remove(img_path)
                except Exception:
                    pass

        return True


class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ming - Mass Image to PDF Converter")
        self.root.geometry("520x450")
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

        self.input_paths = []
        self.config_path = os.path.join(
            os.path.expanduser("~"), ".img2pdf_converter_cfg.json"
        )

        self.out_dir_var = tk.StringVar(
            value=os.path.abspath(os.path.join(os.getcwd(), "pdf_output"))
        )
        self.delete_source_var = tk.BooleanVar(value=False)

        self.load_settings()

        # --- Top Frame ---
        top_frame = tk.Frame(root)
        top_frame.pack(fill=tk.X, padx=20, pady=(15, 5))

        self.add_folder_btn = tk.Button(
            top_frame, text="Add Folder", command=self.add_folders, width=10
        )
        self.add_folder_btn.pack(side=tk.LEFT, padx=(0, 5))
        ToolTip(self.add_folder_btn, "Select one or more input folders from your PC.")

        self.add_zip_btn = tk.Button(
            top_frame, text="Add ZIP", command=self.add_zips, width=10
        )
        self.add_zip_btn.pack(side=tk.LEFT)
        ToolTip(self.add_zip_btn, "Select one or more .zip files directly.")

        tk.Label(top_frame, text="Selected Inputs:", font=("Arial", 9, "bold")).pack(
            side=tk.LEFT, padx=15
        )

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
        self.listbox.dnd_bind("<<Drop>>", self.drop_items)  # type: ignore

        self.listbox.bind("<Delete>", self.remove_selected)
        ToolTip(
            self.listbox,
            "Drag & Drop Folders or .zip files here!\nMultiple selection enabled.\nPress 'Del' or right-click to remove items.",
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
            text="Delete source files after conversion",
            variable=self.delete_source_var,
        )
        self.chk_delete.grid(row=1, column=0, columnspan=2, sticky="w", pady=2)
        ToolTip(
            self.chk_delete,
            "If checked, successfully processed images, ZIPs, and their\ncontaining empty folders will be permanently deleted.",
        )

        # --- Convert Button ---
        self.convert_btn = tk.Button(
            root,
            text="Convert to PDF",
            command=self.process_items,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=5,
        )
        self.convert_btn.pack(pady=(5, 15))
        ToolTip(
            self.convert_btn,
            "Starts processing all folders and ZIPs currently in the list.",
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

    def _add_paths_if_valid(self, paths):
        """Checks for overlaps and adds paths to the list."""
        skipped_count = 0
        for new_path in paths:
            new_path_abs = os.path.abspath(new_path)
            overlap = False

            for existing in self.input_paths:
                existing_abs = os.path.abspath(existing)
                try:
                    # os.path.commonpath finds the common parent folder
                    common = os.path.commonpath([new_path_abs, existing_abs])
                    # If the common path matches exactly one of the two, one is inside the other (or they are the same)
                    if common == existing_abs or common == new_path_abs:
                        overlap = True
                        break
                except ValueError:
                    # Happens when paths are on different drives (e.g., C:\ and D:\)
                    pass

            if not overlap:
                self.input_paths.append(new_path)
                self.listbox.insert(tk.END, new_path)
            else:
                skipped_count += 1

        if skipped_count > 0:
            messagebox.showwarning(
                "Overlap Detected",
                f"{skipped_count} item(s) were skipped because they are already in the list or overlap with existing folders.",
            )

    def add_folders(self):
        folder = filedialog.askdirectory(title="Select Input Folder")
        if folder:
            self._add_paths_if_valid([folder])

    def add_zips(self):
        files = filedialog.askopenfilenames(
            title="Select ZIP files", filetypes=[("ZIP files", "*.zip")]
        )
        if files:
            self._add_paths_if_valid(files)

    def drop_items(self, event):
        items = self.root.tk.splitlist(event.data)
        valid_items = [
            f for f in items if os.path.isdir(f) or f.lower().endswith(".zip")
        ]
        if valid_items:
            self._add_paths_if_valid(valid_items)

    def remove_selected(self, event=None):
        selection = self.listbox.curselection()
        if not selection:
            return
        for i in reversed(selection):
            self.listbox.delete(i)
            del self.input_paths[i]

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
            self.out_dir_var.set(os.path.abspath(folder))

    def process_items(self):
        if not self.input_paths:
            messagebox.showwarning(
                "No Items",
                "Please add at least one input folder or ZIP file to the list.",
            )
            return

        out_dir = os.path.abspath(self.out_dir_var.get())
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
        total_items = len(self.input_paths)

        self.convert_btn.config(state=tk.DISABLED, text="Processing...")
        self.root.update()

        try:
            for item_path in self.input_paths:
                converter = PDFConverter(
                    input_path=item_path,
                    output_folder=out_dir,
                    delete_source=delete_src,
                )
                pdfs_created = converter.convert()
                total_processed_pdfs += pdfs_created

            # Dynamic completion message
            status_msg = (
                "Original files were deleted."
                if delete_src
                else "Original files were KEPT intact."
            )

            messagebox.showinfo(
                "Success",
                f"Done!\n\nProcessed {total_items} item(s).\nCreated {total_processed_pdfs} .pdf files in:\n{out_dir}\n\nNote: {status_msg}",
            )

            # Clear the list after success (optional, but user-friendly)
            self.listbox.delete(0, tk.END)
            self.input_paths.clear()

        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")
        finally:
            self.convert_btn.config(state=tk.NORMAL, text="Convert to PDF")


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFConverterApp(root)
    root.mainloop()
