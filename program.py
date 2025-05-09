import tkinter as tk
from tkinter import ttk, messagebox
from barcode import get_barcode_class
from barcode.writer import ImageWriter
from PIL import Image, ImageTk, ImageDraw, ImageFont
import win32print
import win32ui
from PIL import ImageWin
from io import BytesIO


class BarcodePrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Codabar Barcode Printer")

        self.input_var = tk.StringVar()
        self.printer_var = tk.StringVar()
        self.print_mode = tk.StringVar(value="single")  # New: default to single

        # === Input Frame ===
        input_frame = ttk.Frame(root, padding=15)
        input_frame.pack()

        ttk.Label(input_frame, text="Enter Number:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(input_frame, textvariable=self.input_var, width=30).grid(row=0, column=1, padx=5, pady=5)

        ttk.Button(input_frame, text="Generate Barcode", command=self.generate_barcode).grid(
            row=1, column=0, columnspan=2, pady=(10, 5)
        )

        # === Barcode Preview Canvas ===
        self.canvas = tk.Canvas(root, width=600, height=250, bg="white")
        self.canvas.pack(pady=(5, 10))

        # === Print Mode Selector ===
        mode_frame = ttk.LabelFrame(root, text="Print Mode", padding=10)
        mode_frame.pack(pady=5)

        ttk.Radiobutton(mode_frame, text="Single", variable=self.print_mode, value="single").pack(side="left", padx=10)
        ttk.Radiobutton(mode_frame, text="Triple", variable=self.print_mode, value="triple").pack(side="left", padx=10)

        # === Printer Selection ===
        try:
            server_name = r"\\printserver"
            printers = win32print.EnumPrinters(
                win32print.PRINTER_ENUM_NAME,
                server_name,
                2
            )

            printer_names = [
                printer['pPrinterName']
                for printer in printers
                if "card printer" in printer['pPrinterName'].lower()
            ]

        except Exception as e:
            messagebox.showerror("Printer Load Error", f"Could not load printers from printserver:\n{e}")
            printer_names = []





        ttk.Label(root, text="Select Printer:").pack()
        self.printer_dropdown = ttk.Combobox(root, textvariable=self.printer_var, values=printer_names,
                                             state="readonly")



        self.printer_dropdown.pack(pady=5)
        if printer_names:
            self.printer_var.set(printer_names[0])

        # === Unified Print Button ===
        ttk.Button(root, text="Print", command=self.print_barcode).pack(pady=(10, 20))

    def generate_barcode(self):
        number = self.input_var.get().strip()
        if not number.isdigit():
            messagebox.showerror("Input Error", "Please enter digits only.")
            return

        try:
            wrapped_number = f"A{number}A"
            Codabar = get_barcode_class('codabar')
            barcode = Codabar(wrapped_number, writer=ImageWriter())

            buffer = BytesIO()
            barcode.write(buffer, options={"write_text": False})
            buffer.seek(0)

            barcode_img = Image.open(buffer).convert("RGB")

            try:
                font = ImageFont.truetype("arial.ttf", 60)
            except:
                font = ImageFont.load_default()

            bbox = font.getbbox(number)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]

            barcode_width, barcode_height = barcode_img.size
            total_height = barcode_height + text_height + 50
            combined_img = Image.new("RGB", (barcode_width, total_height), "white")
            combined_img.paste(barcode_img, (0, 0))

            draw = ImageDraw.Draw(combined_img)
            text_x = (barcode_width - text_width) // 2
            text_y = barcode_height + 10
            draw.text((text_x, text_y), number, font=font, fill="black")

            self.image = combined_img

            preview_width = 600  # was 450
            preview_height = int(total_height * (preview_width / barcode_width))
            self.tk_image = ImageTk.PhotoImage(combined_img.resize((preview_width, preview_height)))
            self.canvas.config(width=preview_width, height=preview_height)
            self.canvas.delete("all")
            self.canvas.create_image(preview_width // 2, preview_height // 2, image=self.tk_image)


        except Exception as e:
            messagebox.showerror("Barcode Error", str(e))

    def print_barcode_single(self):
        if not hasattr(self, 'image'):
            messagebox.showerror("Print Error", "Generate the barcode first.")
            return

        printer_name = self.printer_var.get()
        if not printer_name:
            messagebox.showerror("Printer Error", "No printer selected.")
            return

        dpi = 300
        card_width_px = int(2.125 * dpi)  # 637 px
        card_height_px = int(3.375 * dpi)  # 1012 px

        barcode_target_width = 600
        barcode_target_height = 180

        image_resized = self.image.resize((barcode_target_width, barcode_target_height), Image.LANCZOS)

        left = (card_width_px - barcode_target_width) // 2
        top = 20  # fixed top margin
        right = left + barcode_target_width
        bottom = top + barcode_target_height

        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
        hdc.StartDoc("Codabar Print - Top-Aligned")
        hdc.StartPage()

        dib = ImageWin.Dib(image_resized)
        dib.draw(hdc.GetHandleOutput(), (left, top, right, bottom))

        hdc.EndPage()
        hdc.EndDoc()
        hdc.DeleteDC()

        messagebox.showinfo("Print Success", f"Barcode sent to {printer_name} at top of the card.")

    def print_barcode_triple(self):
        if not hasattr(self, 'image'):
            messagebox.showerror("Print Error", "Generate the barcode first.")
            return

        printer_name = self.printer_var.get()
        if not printer_name:
            messagebox.showerror("Printer Error", "No printer selected.")
            return

        dpi = 300
        card_width_px = int(2.125 * dpi)  # 637
        card_height_px = int(3.375 * dpi)  # 1012


        zone_height = card_height_px // 3

        barcode_width = 500
        barcode_height = 140
        header_spacing = 10

        left = (card_width_px - barcode_width) // 2
        right = left + barcode_width

        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
        hdc.StartDoc("Codabar Print - Triple Portrait")
        hdc.StartPage()

        image_resized = self.image.resize((barcode_width, barcode_height), Image.LANCZOS)
        dib = ImageWin.Dib(image_resized)

        font = win32ui.CreateFont({
            "name": "Arial",
            "height": 32,
            "weight": 700
        })
        hdc.SelectObject(font)

        header_text = "reginalibrary.ca | sasklibraries.ca"
        text_size = hdc.GetTextExtent(header_text)
        text_width = text_size[0]
        text_height = text_size[1]

        for i in range(3):
            zone_top = i * zone_height
            top = zone_top + (
            zone_height - barcode_height - text_height - header_spacing) // 2 + text_height + header_spacing
            bottom = top + barcode_height

            header_top = top - header_spacing - text_height
            header_left = (card_width_px - text_width) // 2
            hdc.TextOut(header_left, header_top, header_text)


            dib.draw(hdc.GetHandleOutput(), (left, top, right, bottom))

        hdc.EndPage()
        hdc.EndDoc()
        hdc.DeleteDC()

        messagebox.showinfo("Print Success", f"Triple barcodes printed in 3 horizontal zones on {printer_name}.")

    def print_barcode(self):
        if self.print_mode.get() == "single":
            self.print_barcode_single()
        elif self.print_mode.get() == "triple":
            self.print_barcode_triple()

# === Run the App ===
if __name__ == "__main__":
    root = tk.Tk()
    app = BarcodePrinterApp(root)
    root.mainloop()
