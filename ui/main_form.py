import os
import datetime
from PyQt6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QPushButton, QFrame, QSplitter,
    QToolButton, QTextEdit, QStatusBar, QMessageBox, QFileDialog, QProgressDialog, QApplication, QListWidgetItem, QMenu
)
from PyQt6.QtCore import (
    Qt, QThread, QObject, pyqtSignal, QSettings, QTimer, QDateTime, QDir, QSize, QLocale
)
from PyQt6.QtGui import QIcon, QFont, QColor
import pandas as pd
import json

from core.constants import SCRIPT_DIR
from core.document_generator import DocGenWorker, convert_docx_to_pdf, generate_docx_internal
from core.data_manager import load_session_file, load_sessions, save_session_file, save_session, delete_session_from_db
from core.utilities import OperationCanceledError, CustomTranslator
from ui.widgets.dialogs import show_message_box
from .sidebar import CollapsibleSidebar
from .widgets.merged_form import MergedFormWidget
from .widgets.dialogs import SettingsDialog, DetachedPreviewDialog

class MainForm(QWidget):
    request_job = pyqtSignal(dict, str, str)

    def __init__(self):
        super().__init__()
        self.worker = None
        self.thread = None
        self.settings = QSettings("BillManager", "ThemeSettings")
        self.active_session_info = None
        self.detached_preview_dialog = None
        self.is_preview_detached = False
        self.auto_save_timer = QTimer(self)
        self.auto_save_timer.timeout.connect(self.auto_save)
        self.preview_timer = QTimer(self)
        self.preview_timer.setSingleShot(True)
        self.preview_timer.setInterval(750)
        self.preview_timer.timeout.connect(self.trigger_auto_preview)
        self.last_session_data = load_session_file()
        self.init_ui()
        self.setup_worker_thread()
        self.load_settings()
        self.refresh_sidebar()

    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        self.status_bar = QStatusBar()
        self.status_bar.setSizeGripEnabled(False)
        main_content_widget = QWidget()
        main_content_layout = QVBoxLayout(main_content_widget)
        main_content_layout.setContentsMargins(15, 15, 15, 15)
        main_content_layout.setSpacing(15)
        self.sidebar = CollapsibleSidebar(parent=self)
        self.sidebar.listWidget().itemClicked.connect(self.handle_sidebar_click)
        self.sidebar.listWidget().customContextMenuRequested.connect(self.show_sidebar_context_menu)
        self.form_widget = MergedFormWidget()
        self.form_widget.something_changed.connect(self.start_preview_timer)
        form_card = QFrame(objectName="MainCard")
        form_card.setMaximumWidth(900)
        left_panel_layout = QVBoxLayout(form_card)
        left_panel_layout.addWidget(self.form_widget, 1)
        buttons_layout = QHBoxLayout()
        self.save_docx_btn = QPushButton(self.tr("Save DOCX"), objectName="SaveDocxButton")
        self.save_docx_btn.clicked.connect(self.save_docx)
        self.save_pdf_btn = QPushButton(self.tr("Save PDF"), objectName="SavePdfButton")
        self.save_pdf_btn.clicked.connect(self.save_pdf)
        self.save_pdf_btn.hide()
        self.preview_btn = QPushButton(self.tr("Refresh Preview"), objectName="PreviewButton")
        self.preview_btn.clicked.connect(self.preview_doc)
        self.quick_save_btn = QPushButton(self.tr("Quick Save"))
        self.quick_save_btn.clicked.connect(self.quick_save)
        self.export_btn = QPushButton(self.tr("Export Excel"))
        self.export_btn.clicked.connect(self.export_to_excel)
        self.settings_btn = QPushButton(QIcon(os.path.join(SCRIPT_DIR, "assets", "settings_icon.png")), self.tr("Settings"))
        self.settings_btn.clicked.connect(self.open_settings_dialog)
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.quick_save_btn)
        buttons_layout.addWidget(self.export_btn)
        buttons_layout.addWidget(self.settings_btn)
        buttons_layout.addWidget(self.preview_btn)
        buttons_layout.addWidget(self.save_docx_btn)
        buttons_layout.addWidget(self.save_pdf_btn)
        left_panel_layout.addLayout(buttons_layout)
        self.preview_card = QFrame(objectName="PreviewCard")
        self.preview_card.setMinimumSize(400, 400)
        self.preview_layout = QVBoxLayout(self.preview_card)
        self.preview_layout.setContentsMargins(5, 5, 5, 5)
        zoom_layout = QHBoxLayout()
        zoom_layout.addStretch()
        self.detach_btn = QToolButton(objectName="detachButton")
        self.detach_btn.setText("⏏️")
        self.detach_btn.setToolTip(self.tr("Detach Preview"))
        self.detach_btn.clicked.connect(self.toggle_preview_dock_state)
        zoom_layout.addWidget(self.detach_btn)
        zoom_in_btn = QToolButton(objectName="zoomInButton")
        zoom_in_btn.setText("➕")
        zoom_in_btn.setToolTip(self.tr("Zoom In"))
        zoom_in_btn.clicked.connect(self.zoom_in_preview)
        zoom_layout.addWidget(zoom_in_btn)
        zoom_out_btn = QToolButton(objectName="zoomOutButton")
        zoom_out_btn.setText("➖")
        zoom_out_btn.setToolTip(self.tr("Zoom Out"))
        zoom_out_btn.clicked.connect(self.zoom_out_preview)
        zoom_layout.addWidget(zoom_out_btn)
        self.preview_layout.addLayout(zoom_layout)
        self.preview_widget = QTextEdit()
        self.preview_widget.setReadOnly(True)
        self.preview_layout.addWidget(self.preview_widget, 1)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(form_card)
        splitter.addWidget(self.preview_card)
        splitter.setSizes([550, 950])
        splitter.setHandleWidth(10)
        splitter.setCollapsible(0, False)
        splitter.setCollapsible(1, False)
        main_content_layout.addWidget(splitter, 1)
        main_content_layout.addWidget(self.status_bar)
        main_layout.addWidget(self.sidebar)
        main_layout.addWidget(main_content_widget, 1)

    def retranslate(self):
        self.save_docx_btn.setText(self.tr("Save DOCX"))
        self.save_pdf_btn.setText(self.tr("Save PDF"))
        self.preview_btn.setText(self.tr("Refresh Preview"))
        self.quick_save_btn.setText(self.tr("Quick Save"))
        self.export_btn.setText(self.tr("Export Excel"))
        self.settings_btn.setText(self.tr("Settings"))
        detach_btn = self.findChild(QToolButton, "detachButton")
        if detach_btn:
            detach_btn.setToolTip(self.tr("Detach Preview") if not self.is_preview_detached else self.tr("Dock Preview"))
        zoom_in_btn = self.findChild(QToolButton, "zoomInButton")
        if zoom_in_btn:
            zoom_in_btn.setToolTip(self.tr("Zoom In"))
        zoom_out_btn = self.findChild(QToolButton, "zoomOutButton")
        if zoom_out_btn:
            zoom_out_btn.setToolTip(self.tr("Zoom Out"))
        self.sidebar.retranslate()
        self.form_widget.retranslate()
        self.update_status(self.tr("Ready"))

    def zoom_in_preview(self):
        self.preview_widget.zoomIn(2)

    def zoom_out_preview(self):
        self.preview_widget.zoomOut(2)
        
    def setup_worker_thread(self):
        self.worker = DocGenWorker()
        self.thread = QThread()
        self.worker.moveToThread(self.thread)
        self.request_job.connect(self.worker.run_job)
        self.worker.finished.connect(self.on_worker_finished)
        QApplication.instance().aboutToQuit.connect(self.thread.quit)
        self.thread.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()
        
    def show_sidebar_context_menu(self, pos):
        item = self.sidebar.listWidget().itemAt(pos)
        if not item or not item.data(Qt.ItemDataRole.UserRole):
            return
        session_id, _, _ = item.data(Qt.ItemDataRole.UserRole)
        menu = QMenu()
        delete_action = menu.addAction(self.tr("Delete Session"))
        action = menu.exec(self.sidebar.listWidget().mapToGlobal(pos))
        if action == delete_action:
            reply = QMessageBox.question(self, self.tr('Confirm Delete'), 
                                         self.tr(f"Are you sure you want to delete session '{item.text().splitlines()[0]}'?\nThis action cannot be undone."),
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, 
                                         QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                if delete_session_from_db(session_id):
                    self.refresh_sidebar()
                else:
                    show_message_box(self.tr("Error"), self.tr("Failed to delete the session from the database."))
    
    def load_settings(self):
        dark_mode = self.settings.value("dark_mode", True, type=bool)
        self.update_styles(dark_mode)
        language_code = self.settings.value("language", "en")
        self.apply_language(language_code)
        auto_save_interval = self.settings.value("auto_save_interval", 5, type=int)
        self.auto_save_timer.start(auto_save_interval * 60 * 1000)
        self.backup_location = self.settings.value("backup_location", "")

    def open_settings_dialog(self):
        dialog = SettingsDialog(self.settings, self)
        dialog.darkModeChanged.connect(self.update_styles)
        dialog.languageChanged.connect(self.apply_language)
        dialog.autoSaveChanged.connect(self.update_auto_save_interval)
        dialog.backupPathChanged.connect(self.update_backup_location)
        dialog.exec()

    def update_auto_save_interval(self, minutes):
        self.settings.setValue("auto_save_interval", minutes)
        self.auto_save_timer.setInterval(minutes * 60 * 1000)

    def update_backup_location(self, path):
        self.settings.setValue("backup_location", path)
        self.backup_location = path
        self.update_status(self.tr("Backup path updated."))
    
    def apply_language(self, language_code):
        app = QApplication.instance()
        if hasattr(app, 'current_translator'):
            app.removeTranslator(app.current_translator)
        translator = CustomTranslator(app, language_code)
        app.installTranslator(translator)
        app.current_translator = translator
        self.settings.setValue("language", language_code)
        self.retranslate()

    def update_styles(self, dark_mode):
        self.form_widget.update_styles(dark_mode)
        main_style_sheet = """
            QWidget {{ background-color: {bg}; color: {fg}; }}
            QFrame#MainCard, QFrame#PreviewCard {{ background-color: {bg_card}; border: 1px solid {border}; border-radius: 8px; }}
            QPushButton {{ padding: 11px 20px; border-radius: 4px; font-size: 14px; font-weight: bold; color: white; border: none; }}
            QPushButton:hover {{ background-color: {btn_hover}; }} QPushButton:pressed {{ background-color: {btn_press}; }}
            QPushButton:disabled {{ background-color: #555; color: #999; }}
            QPushButton#SaveDocxButton {{ background-color: #3F51B5; }}
            QPushButton#SavePdfButton {{ background-color: {pdf_bg}; }}
            QPushButton#PreviewButton {{ background-color: #607D8B; }}
            QPushButton#ExportAllButton {{ background-color: #00897B; }}
            QLabel#PreviewPlaceholder {{ color: {preview_fg}; font-size: 16px; padding: 20px; }}
            QSplitter::handle {{ background-color: {handle_bg}; }} QSplitter::handle:hover {{ background-color: {handle_hover}; }}
            QStatusBar {{ background-color: {bg}; color: {fg}; border-top: 1px solid {border}; }}
            QToolButton {{ border: none; }}
            QToolButton:hover {{ background-color: #444; }}
        """.format(
            bg="#353535" if dark_mode else "#F5F5F5",
            fg="#FFFFFF" if dark_mode else "#212121",
            bg_card="#252525" if dark_mode else "#FFFFFF",
            border="#444" if dark_mode else "#E0E0E0",
            btn_hover="#303F9F",
            btn_press="#1A237E",
            pdf_bg="#C62828" if dark_mode else "#D32F2F",
            preview_fg="#888" if dark_mode else "#AAA",
            handle_bg="#444" if dark_mode else "#E0E0E0",
            handle_hover="#3F51B5" if dark_mode else "#C5CAE9"
        )
        self.setStyleSheet(main_style_sheet)

    def find_item_by_data(self, data_to_find):
        if not data_to_find:
            return None
        target_id = data_to_find[0]
        for i in range(self.sidebar.listWidget().count()):
            item = self.sidebar.listWidget().item(i)
            item_data = item.data(Qt.ItemDataRole.UserRole)
            if item_data and item_data[0] == target_id:
                return item
        return None

    def load_session_data(self, item):
        try:
            session_data = item.data(Qt.ItemDataRole.UserRole)
            if not session_data: return
            sid, raw_json, timestamp = session_data
            data = json.loads(raw_json)
            self.form_widget.load_data(data)
            self.active_session_info = session_data
            self.form_widget.clear_dirty()
            self.update_status(self.tr("Loaded session: %s") % item.text().splitlines()[0])
        except RuntimeError:
            print("Warning: Attempted to load data from a deleted session item. Ignoring.")
            self.update_status(self.tr("Warning: Attempted to load data from a deleted session item. Ignoring."))
        except json.JSONDecodeError as e:
            show_message_box(self.tr("Load Error"), self.tr(f"Could not load session data: {e}"))
            self.active_session_info = None
            self.update_status(self.tr("Failed to load session data"))

    def refresh_sidebar(self):
        try:
            current_selection_info = self.active_session_info
            self.sidebar.listWidget().clear()
            new_bill_item = QListWidgetItem(self.tr("➕  New Bill"))
            new_bill_item.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            dark_mode = self.settings.value("dark_mode", True, type=bool)
            new_bill_item.setForeground(QColor("#3F51B5" if dark_mode else "#303F9F"))
            self.sidebar.listWidget().addItem(new_bill_item)
            sessions = load_sessions()
            item_to_select = None
            for sid, name, data, timestamp in sessions:
                try:
                    main_text = name if name else self.tr("Unnamed Session")
                    try:
                        sub_text = datetime.datetime.fromisoformat(timestamp).strftime("%d %b %Y, %I:%M %p")
                    except (ValueError, TypeError):
                        sub_text = self.tr("No date")
                    item = QListWidgetItem(f"{main_text}\n{sub_text}")
                    item.setData(Qt.ItemDataRole.UserRole, (sid, data, timestamp))
                    self.sidebar.listWidget().addItem(item)
                    if current_selection_info and sid == current_selection_info[0]:
                        item_to_select = item
                except Exception as e:
                    print(f"Error loading session {sid}: {str(e)}")
                    continue
            if item_to_select:
                self.sidebar.listWidget().setCurrentItem(item_to_select)
            elif self.active_session_info is None:
                if self.last_session_data and self.last_session_data.get('items'):
                    self.form_widget.load_data(self.last_session_data)
                self.sidebar.listWidget().setCurrentRow(0)
                self.update_status(self.tr("Ready"))
            elif self.sidebar.listWidget().count() > 1:
                self.sidebar.listWidget().setCurrentRow(1)
                self.load_session_data(self.sidebar.listWidget().item(1))
            else:
                self.clear_form()
            self.update_status(self.tr("Session list refreshed"))
        except Exception as e:
            print(f"Error refreshing sidebar: {str(e)}")
            self.update_status(self.tr("Error refreshing sessions"))

    def handle_sidebar_click(self, item):
        clicked_info = item.data(Qt.ItemDataRole.UserRole)
        is_redundant_click = False
        if clicked_info and self.active_session_info and clicked_info[0] == self.active_session_info[0]:
            is_redundant_click = True
        elif clicked_info is None and self.active_session_info is None:
            is_redundant_click = True
        if is_redundant_click:
            return
        if self.form_widget.is_dirty:
            reply = QMessageBox.question(self, self.tr('Unsaved Changes'),
                                         self.tr('You have unsaved changes. Do you want to save them before switching?'),
                                         QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Discard | QMessageBox.StandardButton.Cancel)
            if reply == QMessageBox.StandardButton.Save:
                self.quick_save()
            elif reply == QMessageBox.StandardButton.Cancel:
                item_to_reselect = self.find_item_by_data(self.active_session_info)
                if item_to_reselect:
                    self.sidebar.listWidget().setCurrentItem(item_to_reselect)
                else:
                    self.sidebar.listWidget().setCurrentRow(0)
                return
        if clicked_info is None:
            self.clear_form()
        else:
            item_to_load = self.find_item_by_data(clicked_info)
            if item_to_load:
                self.load_session_data(item_to_load)

    def clear_form(self):
        self.form_widget.clear_form()
        self.preview_widget.clear()
        self.sidebar.listWidget().setCurrentRow(0)
        self.active_session_info = None
        self.form_widget.clear_dirty()
        self.update_status(self.tr("New document ready"))

    def quick_save(self):
        try:
            data = self.form_widget.gather_data()
            session_name = data.get("name", "").strip()
            save_session_file(data)
            self.last_session_data = data
            if not session_name:
                session_name = self.tr("Unnamed Bill - %s") % datetime.datetime.now().strftime('%Y%m%d%H%M%S')
            save_session(session_name, data)
            sessions = load_sessions()
            for sid, name, s_data, timestamp in sessions:
                if name == session_name:
                    self.active_session_info = (sid, s_data, timestamp)
                    break
            self.refresh_sidebar()
            self.form_widget.clear_dirty()
            self.update_status(self.tr("Session saved: %s") % session_name)
        except Exception as e:
            print(f"Error during quick save: {str(e)}")
            self.update_status(self.tr("Save failed - check logs"))

    def auto_save(self):
        if not self.isVisible(): return
        if self.settings.value("auto_save_interval", 0, type=int) > 0 and self.form_widget.is_dirty:
            self.quick_save()

    def _trigger_worker(self, action_type, output_path=""):
        self.set_ui_enabled(False)
        data = self.form_widget.gather_data()
        if not data.get("name") and action_type != "fast_preview":
            show_message_box(self.tr("Missing Info"), self.tr("Please provide a 'Name' in the Document Details before generating a file."))
            self.set_ui_enabled(True)
            return
        self.request_job.emit(data, action_type, output_path)

    def on_worker_finished(self, success, message, result_data):
        self.set_ui_enabled(True)
        if not self.worker: return
        action_type = self.worker.action_type
        if not success:
            show_message_box(self.tr("Error"), self.tr(f"Operation failed: {message}"))
            self.update_status(self.tr(f"Error: {message}"))
            return
        preview_target = self.preview_widget
        if self.is_preview_detached and self.detached_preview_dialog:
            preview_target = self.detached_preview_dialog.preview_widget
        if action_type == "fast_preview":
            if preview_target:
                preview_target.setHtml(result_data)
                self.update_status(self.tr("Preview updated."))
        elif action_type == "preview":
            if preview_target:
                preview_target.setHtml(f"<h1>{self.tr('Slow preview not supported anymore. Use live preview.')}</h1>")
        elif action_type in ["save_docx", "save_pdf"]:
            show_message_box(self.tr("Success"), self.tr(f"File saved to:\n{result_data}"))
            self.quick_save()
            self.update_status(self.tr(f"File saved: {os.path.basename(result_data)}"))
    
    def toggle_preview_dock_state(self):
        if not self.is_preview_detached:
            self.detached_preview_dialog = DetachedPreviewDialog(self.preview_widget, self)
            self.detached_preview_dialog.closed.connect(self.toggle_preview_dock_state)
            self.preview_layout.removeWidget(self.preview_widget)
            self.detached_preview_dialog.layout().addWidget(self.preview_widget)
            self.is_preview_detached = True
            self.detach_btn.setToolTip(self.tr("Dock Preview"))
            self.detach_btn.setText("⬇️")
            self.detached_preview_dialog.show()
            self.preview_widget.update()
        else:
            if self.detached_preview_dialog:
                self.detached_preview_dialog.layout().removeWidget(self.preview_widget)
                self.detached_preview_dialog.close()
            self.preview_layout.addWidget(self.preview_widget)
            self.is_preview_detached = False
            self.detach_btn.setToolTip(self.tr("Detach Preview"))
            self.detach_btn.setText("⏏️")
            self.preview_widget.update()
            self.detached_preview_dialog = None

    def set_ui_enabled(self, enabled):
        for w in [self.save_docx_btn, self.save_pdf_btn, self.preview_btn, self.quick_save_btn, self.export_btn]:
            w.setEnabled(enabled)
        self.sidebar.setEnabled(enabled)
        self.form_widget.setEnabled(enabled)

    def get_default_filename(self, data, extension):
        work_name_raw = data.get('name', 'report')
        sanitized_name = "".join([c for c in work_name_raw if c.isalpha() or c.isdigit() or c==' ']).rstrip().replace(' ', '_')
        today_date = datetime.datetime.now().strftime("%Y-%m-%d")
        return f"{sanitized_name}_{today_date}.{extension}"

    def export_to_excel(self):
        data = self.form_widget.gather_data()
        rows = data.get("items", [])
        if not rows:
            return show_message_box(self.tr("Error"), self.tr("No construction items to export."))
        default_filename = self.get_default_filename(data, "xlsx")
        initial_dir = self.backup_location if self.backup_location else QDir.homePath()
        file_path, _ = QFileDialog.getSaveFileName(self, self.tr("Export to Excel"), os.path.join(initial_dir, default_filename), self.tr("Excel Files (*.xlsx)"))
        if not file_path: return
        try:
            pd.DataFrame(rows).to_excel(file_path, index=False)
            show_message_box(self.tr("Export Successful"), self.tr(f"Data exported to:\n{file_path}"))
            self.update_status(self.tr("Exported to Excel: %s") % os.path.basename(file_path))
        except Exception as e:
            show_message_box(self.tr("Export Failed"), self.tr(f"An error occurred: {str(e)}"))
            self.update_status(self.tr("Export failed: %s") % str(e))

    def start_preview_timer(self):
        self.preview_timer.start()

    def trigger_auto_preview(self):
        if not self.isVisible(): return
        self._trigger_worker("fast_preview")

    def preview_doc(self):
        self._trigger_worker("fast_preview")

    def save_docx(self):
        data = self.form_widget.gather_data()
        default_filename = self.get_default_filename(data, "docx")
        initial_dir = self.backup_location if self.backup_location else QDir.homePath()
        file_path, _ = QFileDialog.getSaveFileName(self, self.tr("Save DOCX"), os.path.join(initial_dir, default_filename), self.tr("Word Documents (*.docx)"))
        if file_path: self._trigger_worker("save_docx", output_path=file_path)

    def save_pdf(self):
        if not pypandoc:
            return show_message_box(self.tr("PDF Not Available"), self.tr("`pypandoc` library not installed."))
        data = self.form_widget.gather_data()
        default_filename = self.get_default_filename(data, "pdf")
        initial_dir = self.backup_location if self.backup_location else QDir.homePath()
        file_path, _ = QFileDialog.getSaveFileName(self, self.tr("Save PDF"), os.path.join(initial_dir, default_filename), self.tr("PDF Documents (*.pdf)"))
        if file_path: self._trigger_worker("save_pdf", output_path=file_path)

    def export_all_reports(self):
        data = self.form_widget.gather_data()
        if not data.get("name"):
            return show_message_box(self.tr("Missing Info"), self.tr("Please provide a 'Name' in the Document Details before exporting."))
        if not pypandoc:
            return show_message_box(self.tr("Dependency Missing"), self.tr("Cannot generate PDF because `pypandoc` is not installed."))
        dir_path = QFileDialog.getExistingDirectory(self, self.tr("Select Directory to Save Report Pack"), self.backup_location if self.backup_location else "")
        if not dir_path: return
        self.set_ui_enabled(False)
        progress = QProgressDialog(self.tr("Generating Report Pack..."), self.tr("Cancel"), 0, 3, self)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setWindowTitle(self.tr("Processing..."))
        progress.show()
        try:
            base_name_no_ext = os.path.splitext(self.get_default_filename(data, ""))[0]
            progress.setLabelText(self.tr("Generating DOCX file...")); progress.setValue(0)
            QApplication.processEvents()
            if progress.wasCanceled(): raise OperationCanceledError()
            docx_path = os.path.join(dir_path, f"{base_name_no_ext}.docx")
            success_docx, msg_docx = generate_docx_internal(data, docx_path)
            if not success_docx: raise Exception(self.tr(f"DOCX generation failed: {msg_docx}"))
            progress.setLabelText(self.tr("Generating PDF file...")); progress.setValue(1)
            QApplication.processEvents()
            if progress.wasCanceled(): raise OperationCanceledError()
            pdf_path = os.path.join(dir_path, f"{base_name_no_ext}.pdf")
            pypandoc_extra_args = [
                '--pdf-engine=xelatex',
                '-V', 'mainfont=Arial'
            ]
            pypandoc.convert_file(docx_path, 'pdf', outputfile=pdf_path, extra_args=pypandoc_extra_args)
            progress.setLabelText(self.tr("Generating Excel file...")); progress.setValue(2)
            QApplication.processEvents()
            if progress.wasCanceled(): raise OperationCanceledError()
            excel_path = os.path.join(dir_path, f"{base_name_no_ext}.xlsx")
            pd.DataFrame(data.get("items", [])).to_excel(excel_path, index=False)
            progress.setValue(3)
            show_message_box(self.tr("Success"), self.tr(f"Report pack saved successfully in:\n{dir_path}"))
            self.update_status(self.tr("Report pack generated."))
        except OperationCanceledError:
            self.update_status(self.tr("Report pack generation canceled."))
        except Exception as e:
            show_message_box(self.tr("Export Failed"), self.tr(f"An error occurred: {str(e)}"))
            self.update_status(self.tr("Report pack failed: %s") % str(e))
        finally:
            progress.close()
            self.set_ui_enabled(True)

    def update_status(self, message):
        timestamp = QDateTime.currentDateTime().toString("hh:mm:ss AP")
        self.status_bar.showMessage(f"{timestamp} | {message}")