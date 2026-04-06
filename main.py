import webview
import pandas as pd
import fitz  # PyMuPDF для створення картинки прев'ю
import io
import base64
import os
import re
import sys
import tempfile
import subprocess
import threading
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

FONT_MAP = {
    'times-bold': {
        'paths': [
            'C:\\Windows\\Fonts\\timesbd.ttf',
            '/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf',
            '/Library/Fonts/Times New Roman Bold.ttf',
            '/Library/Fonts/TimesNewRomanPS-BoldMT.ttf',
        ],
        'fallback': 'Times-Bold',
        'reg_name': 'RegTimesBold',
    },
    'times': {
        'paths': [
            'C:\\Windows\\Fonts\\times.ttf',
            '/System/Library/Fonts/Supplemental/Times New Roman.ttf',
            '/Library/Fonts/Times New Roman.ttf',
        ],
        'fallback': 'Times-Roman',
        'reg_name': 'RegTimes',
    },
    'arial': {
        'paths': [
            'C:\\Windows\\Fonts\\arial.ttf',
            '/Library/Fonts/Arial.ttf',
        ],
        'fallback': 'Helvetica',
        'reg_name': 'RegArial',
    },
    'arial-bold': {
        'paths': [
            'C:\\Windows\\Fonts\\arialbd.ttf',
            '/Library/Fonts/Arial Bold.ttf',
        ],
        'fallback': 'Helvetica-Bold',
        'reg_name': 'RegArialBold',
    },
    'calibri': {
        'paths': [
            'C:\\Windows\\Fonts\\calibri.ttf',
        ],
        'fallback': 'Helvetica',
        'reg_name': 'RegCalibri',
    },
    'calibri-bold': {
        'paths': [
            'C:\\Windows\\Fonts\\calibrib.ttf',
        ],
        'fallback': 'Helvetica-Bold',
        'reg_name': 'RegCalibriBold',
    },
    'georgia': {
        'paths': [
            'C:\\Windows\\Fonts\\georgia.ttf',
            '/Library/Fonts/Georgia.ttf',
        ],
        'fallback': 'Times-Roman',
        'reg_name': 'RegGeorgia',
    },
    'georgia-bold': {
        'paths': [
            'C:\\Windows\\Fonts\\georgiab.ttf',
            '/Library/Fonts/Georgia Bold.ttf',
        ],
        'fallback': 'Times-Bold',
        'reg_name': 'RegGeorgiaBold',
    },
}

class CertificateAPI:
    def __init__(self):
        self.pdf_path = None
        self.excel_path = None
        self.names_list = []
        self._window = None
        self.upload_dir = tempfile.mkdtemp(prefix="autocertificate_")
        self._registered_fonts = set()

    def set_window(self, window):
        self._window = window

    def _build_pdf_preview_response(self, pdf_path):
        doc = fitz.open(pdf_path)
        page = doc.load_page(0)
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        img_data = pix.tobytes("png")
        base64_img = base64.b64encode(img_data).decode('utf-8')
        doc.close()

        return {
            "status": "success",
            "name": os.path.basename(pdf_path),
            "image": f"data:image/png;base64,{base64_img}"
        }

    def _load_excel_names(self, excel_path):
        _, ext = os.path.splitext(excel_path)
        engine = 'xlrd' if ext.lower() == '.xls' else 'openpyxl'
        df = pd.read_excel(excel_path, engine=engine)
        print(f"[DEBUG] Excel прочитаний (engine={engine}), колонок: {len(df.columns)}, рядків: {len(df)}")
        # Шукаємо першу колонку, що має хоча б одне непусте значення
        names_col = None
        for i in range(len(df.columns)):
            col = df.iloc[:, i].dropna().astype(str).str.strip()
            col = col[col != '']
            if len(col) > 0:
                names_col = col.tolist()
                print(f"[DEBUG] Знайдено імена у колонці {i}: {len(names_col)} записів")
                break
        self.names_list = names_col if names_col else []
        print(f"[DEBUG] Знайдено імен: {len(self.names_list)}")

        return {
            "status": "success",
            "file": os.path.basename(excel_path),
            "count": len(self.names_list),
            "first_name": self.names_list[0] if self.names_list else ""
        }

    def _save_uploaded_file(self, file_name, file_data, expected_extensions):
        if not file_name:
            raise ValueError("Не вказано назву файлу")

        _, extension = os.path.splitext(file_name)
        extension = extension.lower()
        if extension not in expected_extensions:
            allowed = ", ".join(expected_extensions)
            raise ValueError(f"Непідтримуваний формат. Дозволено: {allowed}")

        if "," in file_data:
            file_data = file_data.split(",", 1)[1]

        safe_name = os.path.basename(file_name)
        target_path = os.path.join(self.upload_dir, safe_name)

        with open(target_path, "wb") as uploaded_file:
            uploaded_file.write(base64.b64decode(file_data))

        return target_path

    def _validate_selected_file(self, selected_path, expected_extensions):
        if not selected_path:
            return {"status": "cancelled"}

        _, extension = os.path.splitext(selected_path)
        extension = extension.lower()
        if extension not in expected_extensions:
            allowed = ", ".join(sorted(expected_extensions))
            raise ValueError(f"Оберіть файл формату: {allowed}")

        return selected_path

    def _run_osascript(self, script):
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                check=False,
            )
        except Exception as error:
            raise RuntimeError(f"Не вдалося запустити системний діалог: {error}") from error

        if result.returncode != 0:
            stderr = (result.stderr or "").strip()
            if "User canceled" in stderr:
                return None
            raise RuntimeError(stderr or "Помилка системного діалогу")

        return result.stdout.strip() or None

    def _select_file_dialog(self, title):
        if sys.platform == "darwin":
            script = f'POSIX path of (choose file with prompt "{title}")'
            selected_path = self._run_osascript(script)
            return [selected_path] if selected_path else None

        return self._window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=False)

    def _select_folder_dialog(self, title):
        if sys.platform == "darwin":
            script = f'POSIX path of (choose folder with prompt "{title}")'
            selected_path = self._run_osascript(script)
            return [selected_path] if selected_path else None

        return self._window.create_file_dialog(webview.FOLDER_DIALOG)

    def _resolve_font(self, font_key):
        """Знаходить і реєструє шрифт за ключем, повертає зареєстровану назву."""
        font_info = FONT_MAP.get(font_key, FONT_MAP['times-bold'])
        reg_name = font_info['reg_name']
        if reg_name in self._registered_fonts:
            return reg_name
        for path in font_info['paths']:
            if os.path.exists(path):
                try:
                    pdfmetrics.registerFont(TTFont(reg_name, path))
                    self._registered_fonts.add(reg_name)
                    print(f"[DEBUG] Шрифт зареєстрований: {reg_name} ({path})")
                    return reg_name
                except Exception as e:
                    print(f"[DEBUG] Помилка реєстрації шрифту '{font_key}': {e}")
        print(f"[DEBUG] Шрифт '{font_key}' не знайдено, використовується: {font_info['fallback']}")
        return font_info['fallback']

    def uploadPdf(self, file_name, file_data):
        try:
            self.pdf_path = self._save_uploaded_file(file_name, file_data, {".pdf"})
            print(f"[DEBUG] PDF завантажений з фронтенду: {self.pdf_path}")
            return self._build_pdf_preview_response(self.pdf_path)
        except Exception as e:
            print(f"[ERROR] uploadPdf помилка: {str(e)}")
            return {"status": "error", "message": str(e)}

    def uploadExcel(self, file_name, file_data):
        try:
            self.excel_path = self._save_uploaded_file(file_name, file_data, {".xlsx", ".xls"})
            print(f"[DEBUG] Excel завантажений з фронтенду: {self.excel_path}")
            return self._load_excel_names(self.excel_path)
        except Exception as e:
            print(f"[ERROR] uploadExcel помилка: {str(e)}")
            return {"status": "error", "message": str(e)}

    def selectPdf(self):
        """Відкриває діалог вибору PDF та повертає base64 зображення для прев'ю."""
        try:
            result = self._select_file_dialog("Оберіть PDF шаблон")
            
            if result:
                self.pdf_path = self._validate_selected_file(result[0], {".pdf"})
                print(f"[DEBUG] PDF обраний: {self.pdf_path}")
                
                try:
                    response = self._build_pdf_preview_response(self.pdf_path)
                    print(f"[DEBUG] PDF preview успішно створено")
                    return response
                except Exception as e:
                    print(f"[ERROR] Помилка при створенні preview: {str(e)}")
                    return {"status": "error", "message": f"Помилка при створенні preview: {str(e)}"}
            else:
                print(f"[DEBUG] PDF не обраний")
                return {"status": "cancelled"}
        except Exception as e:
            print(f"[ERROR] selectPdf помилка: {str(e)}")
            return {"status": "error", "message": str(e)}

    def selectExcel(self):
        """Відкриває діалог вибору Excel та зчитує ПІБ."""
        try:
            result = self._select_file_dialog("Оберіть Excel файл")
            
            if result:
                self.excel_path = self._validate_selected_file(result[0], {".xlsx", ".xls"})
                print(f"[DEBUG] Excel обраний: {self.excel_path}")
                
                try:
                    response = self._load_excel_names(self.excel_path)
                    return response
                except Exception as e:
                    print(f"[ERROR] Помилка при читанні Excel: {str(e)}")
                    return {"status": "error", "message": f"Помилка при читанні Excel: {str(e)}"}
            else:
                print(f"[DEBUG] Excel не обраний")
                return {"status": "cancelled"}
        except Exception as e:
            print(f"[ERROR] selectExcel помилка: {str(e)}")
            return {"status": "error", "message": str(e)}

    def generateCertificates(self, x, y, font_size_fraction, font_key='times-bold', color='#68313a'):
        """Генерує сертифікати асинхронно; повертає статус 'started' і надсилає прогрес через JS."""
        try:
            print(f"[DEBUG] generateCertificates: x={x}, y={y}, font_size_fraction={font_size_fraction}, font_key={font_key}, color={color}")
            if not self.pdf_path or not self.names_list:
                msg = "Завантажте PDF та Excel файл!"
                print(f"[ERROR] {msg}")
                return {"status": "error", "message": msg}

            output_dir = self._select_folder_dialog("Оберіть папку для збереження сертифікатів")
            if not output_dir:
                print(f"[DEBUG] Папка для збереження не обрана")
                return {"status": "cancelled"}

            output_dir = output_dir[0]
            print(f"[DEBUG] Папка для збереження: {output_dir}")

            total = len(self.names_list)
            thread = threading.Thread(
                target=self._generate_certificates_thread,
                args=(x, y, font_size_fraction, font_key, color, output_dir),
                daemon=True,
            )
            thread.start()
            return {"status": "started", "total": total}

        except Exception as e:
            print(f"[ERROR] generateCertificates помилка: {str(e)}")
            import traceback
            traceback.print_exc()
            return {"status": "error", "message": str(e)}

    def _generate_certificates_thread(self, x, y, font_size_fraction, font_key, color, output_dir):
        """Фоновий потік генерації сертифікатів."""
        try:
            x = float(x)
            y = float(y)
            font_size_fraction = float(font_size_fraction)

            # Читаємо шаблон один раз у пам'ять
            with open(self.pdf_path, "rb") as f:
                template_bytes = f.read()

            template_reader = PdfReader(io.BytesIO(template_bytes))
            template_page = template_reader.pages[0]
            pdf_width = float(template_page.mediabox.width)
            pdf_height = float(template_page.mediabox.height)
            print(f"[DEBUG] Розміри PDF: {pdf_width}x{pdf_height}")

            font_size = font_size_fraction * pdf_width
            font_name = self._resolve_font(font_key)

            if not re.match(r'^#[0-9a-fA-F]{6}$', str(color)):
                color = '#68313a'

            real_x = (x / 100.0) * pdf_width
            real_y = (1.0 - y / 100.0) * pdf_height - font_size * 0.35
            print(f"[DEBUG] Real координати: x={real_x}, y={real_y}")

            total = len(self.names_list)
            count = 0
            for name in self.names_list:
                try:
                    packet = io.BytesIO()
                    can = canvas.Canvas(packet, pagesize=(pdf_width, pdf_height))
                    can.setFont(font_name, font_size)
                    can.setFillColor(HexColor(color))
                    can.drawCentredString(real_x, real_y, name)
                    can.save()

                    packet.seek(0)
                    text_pdf = PdfReader(packet)

                    # Щоразу створюємо новий reader із кешованих байт — без читання диска
                    fresh_template = PdfReader(io.BytesIO(template_bytes))
                    output = PdfWriter()
                    page = fresh_template.pages[0]
                    page.merge_page(text_pdf.pages[0])
                    output.add_page(page)

                    safe_name = "".join([c for c in name if c.isalnum() or c in ' -']).strip()
                    if not safe_name:
                        safe_name = f"Certificate_{count}"
                    save_path = os.path.join(output_dir, f"Certificate_{safe_name}.pdf")

                    with open(save_path, "wb") as out_stream:
                        output.write(out_stream)

                    count += 1
                    print(f"[DEBUG] Сертифікат {count}/{total}: {save_path}")
                    self._window.evaluate_js(f"updateProgress({count}, {total})")
                except Exception as e:
                    print(f"[ERROR] Помилка при обробці імені '{name}': {str(e)}")

            msg = f"Успішно! Згенеровано {count} сертифікатів з {total}"
            print(f"[DEBUG] {msg}")
            safe_msg = msg.replace("'", "\\'")
            self._window.evaluate_js(
                f"generationComplete({{\"status\":\"success\",\"message\":\"{safe_msg}\"}})"
            )
        except Exception as e:
            print(f"[ERROR] _generate_certificates_thread: {str(e)}")
            import traceback
            traceback.print_exc()
            safe_err = str(e).replace("'", "\\'").replace('"', '\\"')
            self._window.evaluate_js(
                f"generationComplete({{\"status\":\"error\",\"message\":\"{safe_err}\"}})"
            )

def resource_path(relative_path):
    """Повертає правильний шлях як для запуску зі скрипту, так і з .exe (PyInstaller)."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


if __name__ == '__main__':
    print("[START] Запуск Генератора Сертифікатів...")
    api = CertificateAPI()
    
    html_path = resource_path('index.html')
    
    print(f"[DEBUG] HTML файл: {html_path}")
    print(f"[DEBUG] HTML існує: {os.path.exists(html_path)}")
    
    window = webview.create_window(
        'CertificateForge',
        html_path,
        js_api=api,
        width=1200,
        height=800,
        min_size=(1000, 600)
    )
    api.set_window(window)
    print("[DEBUG] Вікно створено, запуск webview...")
    webview.start(debug=False)