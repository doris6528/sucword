import sys
import os
import re
import ctypes
from ctypes import c_int, byref, sizeof
from html.parser import HTMLParser

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QAction, QFileDialog,
    QFontComboBox, QComboBox, QToolBar, QColorDialog, QLabel,
    QVBoxLayout, QHBoxLayout, QWidget, QStatusBar, QMessageBox,
    QToolButton, QMenu, QScrollArea
)
from PyQt5.QtGui import (
    QFont, QIcon, QTextCharFormat, QKeySequence, QPainter, QColor, QPen, 
    QFontDatabase
)
from PyQt5.QtCore import Qt, QRectF

import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


# --- Trình chuyển đổi HTML từ PyQt5 sang chuẩn Microsoft Word (.docx) ---
class PyQtHtmlToDocxParser(HTMLParser):
    def __init__(self, doc):
        super().__init__()
        self.doc = doc
        self.current_p = None
        self.fmt_stack = [{}]

    def handle_starttag(self, tag, attrs):
        attr_dict = dict(attrs)
        style = attr_dict.get('style', '')
        align = attr_dict.get('align', '')

        if tag in ['p', 'div', 'h1', 'h2', 'h3']:
            self.current_p = self.doc.add_paragraph()
            if 'align:center' in style or 'text-align:center' in style or align == 'center':
                self.current_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif 'align:right' in style or 'text-align:right' in style or align == 'right':
                self.current_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif 'align:justify' in style or 'text-align:justify' in style or align == 'justify':
                self.current_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            else:
                self.current_p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        new_fmt = self.fmt_stack[-1].copy()

        if tag in ['b', 'strong']:
            new_fmt['bold'] = True
        if tag in ['i', 'em']:
            new_fmt['italic'] = True
        if tag in ['u']:
            new_fmt['underline'] = True

        if 'font-weight:600' in style or 'font-weight:bold' in style:
            new_fmt['bold'] = True
        if 'font-style:italic' in style:
            new_fmt['italic'] = True
        if 'text-decoration: underline' in style:
            new_fmt['underline'] = True

        c_match = re.search(r'color:\s*(#[0-9a-fA-F]{6})', style)
        if c_match:
            new_fmt['color'] = c_match.group(1)

        s_match = re.search(r'font-size:\s*(\d+(\.\d+)?)pt', style)
        if s_match:
            new_fmt['size'] = float(s_match.group(1))

        f_match = re.search(r"font-family:\s*'?([^;']+)'?", style)
        if f_match:
            new_fmt['font_name'] = f_match.group(1).replace("'", "").replace('"', '').strip()

        self.fmt_stack.append(new_fmt)

    def handle_endtag(self, tag):
        if len(self.fmt_stack) > 1:
            self.fmt_stack.pop()

    def handle_data(self, data):
        if not data:
            return
        if self.current_p is None:
            self.current_p = self.doc.add_paragraph()
        
        fmt = self.fmt_stack[-1]
        run = self.current_p.add_run(data)
        if fmt.get('bold'): 
            run.bold = True
        if fmt.get('italic'): 
            run.italic = True
        if fmt.get('underline'): 
            run.underline = True
        if fmt.get('color'):
            c = fmt['color']
            if c.startswith('#') and len(c) == 7:
                run.font.color.rgb = RGBColor(int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16))
        if fmt.get('size'):
            run.font.size = Pt(fmt['size'])
        if fmt.get('font_name'):
            run.font.name = fmt['font_name']


# --- Khung chứa 1 Trang giấy dài duy nhất tự động mở rộng theo nội dung ---
class PageCanvas(QWidget):
    PAGE_WIDTH = 816     # Chiều rộng chuẩn A4 (96 DPI)
    MIN_HEIGHT = 1056    # Chiều cao tối thiểu (1 trang A4)
    PADDING_TOP_BOTTOM = 30

    def __init__(self, parent=None):
        super().__init__(parent)
        self.paper_height = self.MIN_HEIGHT
        self.setMinimumWidth(self.PAGE_WIDTH + 100)
        self.setFixedHeight(self.MIN_HEIGHT + self.PADDING_TOP_BOTTOM * 2)

    def update_canvas_height(self, content_height):
        # Tính toán chiều cao trang giấy trắng mới
        new_paper_h = max(self.MIN_HEIGHT, int(content_height) + 120)
        if self.paper_height != new_paper_h:
            self.paper_height = new_paper_h
            self.setFixedHeight(self.paper_height + self.PADDING_TOP_BOTTOM * 2)
            self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Nền màu xám không gian làm việc
        painter.fillRect(self.rect(), QColor("#e9ecef"))
        
        x_offset = (self.width() - self.PAGE_WIDTH) // 2
        y_offset = self.PADDING_TOP_BOTTOM
        
        # Bóng đổ phía dưới trang giấy
        shadow_rect = QRectF(x_offset + 4, y_offset + 4, self.PAGE_WIDTH, self.paper_height)
        painter.fillRect(shadow_rect, QColor(0, 0, 0, 20))
        
        # Mặt trang giấy trắng A4 duy nhất
        page_rect = QRectF(x_offset, y_offset, self.PAGE_WIDTH, self.paper_height)
        painter.fillRect(page_rect, QColor("#ffffff"))
        painter.setPen(QPen(QColor("#ced4da"), 1))
        painter.drawRect(page_rect)


class SucWord(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_path = None
        self.load_custom_fonts()
        self.init_ui()

    def load_custom_fonts(self):
        """Tự động đăng ký Font Lora từ file .ttf nếu có sẵn trong thư mục ứng dụng"""
        font_files = ["Lora-Regular.ttf", "Lora.ttf", "lora.ttf"]
        for font_file in font_files:
            if os.path.exists(font_file):
                QFontDatabase.addApplicationFont(font_file)
                break

    def init_ui(self):
        self.setWindowTitle("SucWord - Trình soạn thảo văn bản (.docx)")
        self.setGeometry(100, 100, 1150, 850)
        
        logo_path = "logo.png"
        if os.path.exists(logo_path):
            self.setWindowIcon(QIcon(logo_path))

        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # --- Thanh công cụ Toolbar ---
        toolbar = QToolBar("Định dạng")
        toolbar.setStyleSheet("""
            QToolBar { 
                background-color: #f8f9fa; 
                border-bottom: 1px solid #dcdcdc; 
                padding: 6px; 
                spacing: 8px; 
            }
            QToolButton { 
                background-color: transparent; 
                border: 1px solid transparent; 
                border-radius: 4px; 
                padding: 4px 8px; 
                font-weight: bold; 
            }
            QToolButton:hover { 
                background-color: #e2e6ea; 
                border: 1px solid #dae0e5; 
            }
            QToolButton:checked { 
                background-color: #c8dede; 
                border: 1px solid #adcaca; 
            }
            QMenu {
                background-color: #ffffff;
                border: 1px solid #ced4da;
                padding: 4px;
            }
            QMenu::item {
                padding: 6px 20px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: #e9ecef;
            }
        """)
        self.addToolBar(toolbar)

        # File actions
        new_action = QAction("Tạo mới", self)
        new_action.setShortcut(QKeySequence.New)
        new_action.triggered.connect(self.file_new)
        toolbar.addAction(new_action)

        open_action = QAction("Mở tệp", self)
        open_action.setShortcut(QKeySequence.Open)
        open_action.triggered.connect(self.file_open)
        toolbar.addAction(open_action)

        save_action = QAction("Lưu (.docx)", self)
        save_action.setShortcut(QKeySequence.Save)
        save_action.triggered.connect(self.file_save)
        toolbar.addAction(save_action)

        toolbar.addSeparator()

        # Font & Size
        self.font_box = QFontComboBox()
        self.font_box.setCurrentFont(QFont("Lora"))
        self.font_box.currentFontChanged.connect(self.change_font_family)
        toolbar.addWidget(self.font_box)

        self.size_box = QComboBox()
        self.size_box.setEditable(True)
        for size in [8, 9, 10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72]:
            self.size_box.addItem(str(size))
        self.size_box.setCurrentText("13")
        self.size_box.currentTextChanged.connect(self.change_font_size)
        toolbar.addWidget(self.size_box)

        toolbar.addSeparator()

        # B, I, U
        self.bold_action = QAction("B", self)
        self.bold_action.setCheckable(True)
        self.bold_action.setShortcut(QKeySequence.Bold)
        self.bold_action.triggered.connect(self.toggle_bold)
        toolbar.addAction(self.bold_action)

        self.italic_action = QAction("I", self)
        self.italic_action.setCheckable(True)
        self.italic_action.setShortcut(QKeySequence.Italic)
        self.italic_action.triggered.connect(self.toggle_italic)
        toolbar.addAction(self.italic_action)

        self.underline_action = QAction("U", self)
        self.underline_action.setCheckable(True)
        self.underline_action.setShortcut(QKeySequence.Underline)
        self.underline_action.triggered.connect(self.toggle_underline)
        toolbar.addAction(self.underline_action)

        toolbar.addSeparator()

        # Text Color
        color_action = QAction("Màu chữ", self)
        color_action.triggered.connect(self.change_text_color)
        toolbar.addAction(color_action)

        toolbar.addSeparator()

        # Menu sổ xuống căn lề
        self.align_button = QToolButton(self)
        self.align_button.setText("Căn lề ▾")
        self.align_button.setPopupMode(QToolButton.InstantPopup)
        
        align_menu = QMenu(self.align_button)
        act_left = QAction("Căn trái", self)
        act_left.triggered.connect(lambda: self.set_alignment(Qt.AlignLeft, "Căn trái ▾"))
        align_menu.addAction(act_left)

        act_center = QAction("Căn giữa", self)
        act_center.triggered.connect(lambda: self.set_alignment(Qt.AlignCenter, "Căn giữa ▾"))
        align_menu.addAction(act_center)

        act_right = QAction("Căn phải", self)
        act_right.triggered.connect(lambda: self.set_alignment(Qt.AlignRight, "Căn phải ▾"))
        align_menu.addAction(act_right)

        act_justify = QAction("Căn đều", self)
        act_justify.triggered.connect(lambda: self.set_alignment(Qt.AlignJustify, "Căn đều ▾"))
        align_menu.addAction(act_justify)

        self.align_button.setMenu(align_menu)
        toolbar.addWidget(self.align_button)

        # --- VÙNG SOẠN THẢO TRANG DÀI (Scroll Area & Canvas) ---
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("QScrollArea { border: none; background-color: #e9ecef; }")

        self.canvas = PageCanvas()
        canvas_layout = QHBoxLayout(self.canvas)
        # Đặt lề trên 50px cho vùng nhập văn bản bên trong tờ giấy trắng
        canvas_layout.setContentsMargins(0, PageCanvas.PADDING_TOP_BOTTOM + 50, 0, PageCanvas.PADDING_TOP_BOTTOM + 50)
        canvas_layout.setSpacing(0)

        # Trình biên soạn Editor chuẩn
        self.editor = QTextEdit()
        self.editor.setFixedWidth(PageCanvas.PAGE_WIDTH - 100) # Rộng 716px, lề trái/phải 50px
        self.editor.setFont(QFont("Lora", 13))
        self.editor.setStyleSheet("""
            QTextEdit {
                background: transparent;
                border: none;
                font-size: 13pt;
                font-family: 'Lora', 'Segoe UI', serif;
            }
        """)

        # Tắt thanh cuộn riêng của Editor
        self.editor.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.editor.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self.editor.textChanged.connect(self.on_text_changed)
        self.editor.cursorPositionChanged.connect(self.on_cursor_changed)

        canvas_layout.addStretch()
        # Thêm cờ Qt.AlignTop để ghim vùng soạn thảo ở đỉnh trang giấy
        canvas_layout.addWidget(self.editor, 0, Qt.AlignTop)
        canvas_layout.addStretch()

        self.scroll_area.setWidget(self.canvas)
        main_layout.addWidget(self.scroll_area)

        # Status Bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_label = QLabel("Số từ: 0  |  Ký tự: 0")
        self.status_bar.addPermanentWidget(self.status_label)

        self.setCentralWidget(main_widget)

    def showEvent(self, event):
        super().showEvent(event)
        try:
            hwnd = int(self.winId())
            color_ref = 0x00FAF9F8
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 35, byref(c_int(color_ref)), sizeof(c_int)
            )
        except Exception:
            pass

    def on_text_changed(self):
        # Tự động mở rộng trang giấy theo chiều cao văn bản thực tế
        doc_height = self.editor.document().documentLayout().documentSize().height()
        self.editor.setFixedHeight(max(100, int(doc_height) + 20))
        self.canvas.update_canvas_height(doc_height)
        self.update_status_bar()
        self.auto_scroll_to_cursor()

    def on_cursor_changed(self):
        self.update_format_buttons()
        self.auto_scroll_to_cursor()

    def auto_scroll_to_cursor(self):
        """Tự động cuộn trang khi con trỏ gõ xuống gần mép dưới màn hình"""
        cursor_rect = self.editor.cursorRect()
        viewport = self.scroll_area.viewport()
        
        cursor_y_in_viewport = self.editor.mapTo(viewport, cursor_rect.bottomLeft()).y()
        viewport_height = viewport.height()
        
        margin_bottom = 60
        if cursor_y_in_viewport > viewport_height - margin_bottom:
            delta = cursor_y_in_viewport - (viewport_height - margin_bottom)
            v_bar = self.scroll_area.verticalScrollBar()
            v_bar.setValue(int(v_bar.value() + delta))

    def set_alignment(self, alignment, text):
        self.editor.setAlignment(alignment)
        self.align_button.setText(text)

    # --- Quản lý File .DOCX ---
    def file_new(self):
        self.editor.clear()
        self.current_path = None
        self.setWindowTitle("SucWord - Trình soạn thảo văn bản (.docx)")
        self.on_text_changed()

    def file_open(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Mở tệp Word", "", "Word Document (*.docx)"
        )
        if path:
            doc_obj = docx.Document(path)
            text = "\n".join([p.text for p in doc_obj.paragraphs])
            self.editor.setPlainText(text)
            self.current_path = path
            self.setWindowTitle(f"SucWord - {os.path.basename(path)}")
            self.on_text_changed()

    def file_save(self):
        if not self.current_path or not self.current_path.endswith('.docx'):
            path, _ = QFileDialog.getSaveFileName(self, "Lưu file Word", "", "Word Document (*.docx)")
            if not path:
                return
            if not path.endswith('.docx'):
                path += '.docx'
            self.current_path = path

        doc_obj = docx.Document()
        html_str = self.editor.toHtml()
        parser = PyQtHtmlToDocxParser(doc_obj)
        parser.feed(html_str)
        
        doc_obj.save(self.current_path)
        self.setWindowTitle(f"SucWord - {os.path.basename(self.current_path)}")
        QMessageBox.information(self, "Lưu thành công", f"Đã lưu tài liệu ra file Word:\n{self.current_path}")

    # --- Các thao tác định dạng Rich Text ---
    def toggle_bold(self):
        fmt = QTextCharFormat()
        fmt.setFontWeight(QFont.Bold if self.bold_action.isChecked() else QFont.Normal)
        self.editor.mergeCurrentCharFormat(fmt)

    def toggle_italic(self):
        fmt = QTextCharFormat()
        fmt.setFontItalic(self.italic_action.isChecked())
        self.editor.mergeCurrentCharFormat(fmt)

    def toggle_underline(self):
        fmt = QTextCharFormat()
        fmt.setFontUnderline(self.underline_action.isChecked())
        self.editor.mergeCurrentCharFormat(fmt)

    def change_font_family(self, font):
        fmt = QTextCharFormat()
        fmt.setFontFamily(font.family())
        self.editor.mergeCurrentCharFormat(fmt)

    def change_font_size(self, size_str):
        try:
            size = float(size_str)
            fmt = QTextCharFormat()
            fmt.setFontPointSize(size)
            self.editor.mergeCurrentCharFormat(fmt)
        except ValueError:
            pass

    def change_text_color(self):
        color = QColorDialog.getColor(self.editor.textColor(), self)
        if color.isValid():
            fmt = QTextCharFormat()
            fmt.setForeground(color)
            self.editor.mergeCurrentCharFormat(fmt)

    def update_format_buttons(self):
        fmt = self.editor.currentCharFormat()
        self.bold_action.setChecked(fmt.fontWeight() == QFont.Bold)
        self.italic_action.setChecked(fmt.fontItalic())
        self.underline_action.setChecked(fmt.fontUnderline())

    def update_status_bar(self):
        text = self.editor.toPlainText()
        words = len(text.split()) if text.strip() else 0
        chars = len(text)
        self.status_label.setText(f"Số từ: {words}  |  Ký tự: {chars}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    word_app = SucWord()
    word_app.show()
    sys.exit(app.exec_())