import os
import traceback
import pandas as pd
from io import BytesIO

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.filechooser import FileChooserListView
from kivy.utils import platform

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas


class MediMapProAutomator(App):
    def build(self):
        root = BoxLayout(orientation="vertical", padding=20, spacing=15)

        self.status = Label(
            text="MediMap Pro APK\nStable GitHub Build",
            halign="center",
            valign="middle"
        )
        self.status.bind(size=self._update_text_size)
        root.add_widget(self.status)

        initial_path = "/sdcard/Download" if platform == "android" else os.path.expanduser("~")
        self.pdf_chooser = FileChooserListView(path=initial_path, filters=["*.pdf"])
        root.add_widget(self.pdf_chooser)

        self.run_btn = Button(
            text="GENERATE AUTOMATED FORMS",
            size_hint_y=None,
            height=100
        )
        self.run_btn.bind(on_release=self.execute_logic)
        root.add_widget(self.run_btn)

        return root

    def _update_text_size(self, instance, value):
        instance.text_size = value

    def set_status(self, message):
        self.status.text = message

    def create_overlay_page(self, page_width, page_height, row):
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=(page_width, page_height))

        can.setFont("Helvetica", 10)
        can.setFillColorRGB(0, 0, 1)

        # Example fixed drawing logic
        # Replace these with your real coordinates later
        name = str(row.get("NAME", "Unknown"))
        ref = str(row.get("REF", ""))

        # Sample marks/text
        can.drawString(80, page_height - 100, f"NAME: {name}")
        if ref:
            can.drawString(80, page_height - 120, f"REF: {ref}")

        # Sample blue filled box
        can.rect(100, 500, 20, 20, fill=1, stroke=0)

        can.save()
        packet.seek(0)
        return PdfReader(packet).pages[0]

    def process_pdf(self, pdf_path, excel_path):
        df = pd.read_excel(excel_path)

        generated_files = []
        for _, row in df.iterrows():
            name = str(row.get("NAME", "Unknown")).strip() or "Unknown"

            reader = PdfReader(pdf_path)
            writer = PdfWriter()

            for page in reader.pages:
                page_width = float(page.mediabox.width)
                page_height = float(page.mediabox.height)

                overlay = self.create_overlay_page(page_width, page_height, row)
                page.merge_page(overlay)
                writer.add_page(page)

            safe_name = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in name).strip()
            if not safe_name:
                safe_name = "Unknown"

            output_path = os.path.join(
                os.path.dirname(pdf_path),
                f"Filled_{safe_name}.pdf"
            )

            with open(output_path, "wb") as f:
                writer.write(f)

            generated_files.append(output_path)

        return generated_files

    def execute_logic(self, instance):
        try:
            if not self.pdf_chooser.selection:
                self.set_status("Error: Please select a PDF form")
                return

            pdf_path = self.pdf_chooser.selection[0]
            base_dir = os.path.dirname(pdf_path)
            excel_path = os.path.join(base_dir, "data.xlsx")

            if not os.path.exists(excel_path):
                self.set_status("Error: 'data.xlsx' missing in folder")
                return

            self.set_status("Processing... Please wait.")
            generated = self.process_pdf(pdf_path, excel_path)

            self.set_status(
                "SUCCESS!\n\nGenerated files:\n" + "\n".join(os.path.basename(x) for x in generated[:10])
            )

        except Exception as e:
            traceback.print_exc()
            self.set_status(f"Process Error:\n{str(e)}")


if __name__ == "__main__":
    MediMapProAutomator().run()
