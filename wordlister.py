import os
import pandas as pd
from collections import deque
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QPushButton,
    QMessageBox, QHBoxLayout, QFrame, QAction, QDialog, QFormLayout,
    QLineEdit, QDialogButtonBox, QSpinBox, QFileDialog, QFontDialog, QProgressBar
)
from PyQt5.QtCore import Qt, QSettings, QTimer
from PyQt5.QtGui import QFont, QPixmap

score_buckets = [61, 60, 50, 25, 0]
bucket_colors = {
    0: "#e74c3c",
    25: "#f39c12",
    50: "#3498db",
    60: "#2ecc71",
    61: "#9b59b6"
}

def map_score_to_bucket(score):
    return min(score_buckets, key=lambda b: abs(b - score))

class SettingsDialog(QDialog):
    def __init__(self, settings: QSettings, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Word Rescoring Settings")
        self.settings = settings

        main_layout = QVBoxLayout(self)

        title_label = QLabel("Adjust Word Rescoring Settings")
        title_label.setFont(QFont("Segoe UI", 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        instructions_label = QLabel(
            "Configure file paths, filters, timing, and display options.\n\n"
            "Master Wordlist: Original read-only source.\n"
            "Personal Wordlist & Tracker: In-memory updates, saved on 'S' or exit.\n"
            "Filters: Limit words.\n"
            "Timing: Minimal delays.\n"
            "Display: Font for main word.\n"
            "Progress: Uses an in-memory counter rather than scanning tracker repeatedly.\n\n"
            "Keys: Q/E to jump 2 buckets up/down."
        )
        instructions_label.setFont(QFont("Segoe UI", 12))
        instructions_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(instructions_label)

        master_wordlist_file = self.settings.value("master_wordlist_file", r"master_wordlist.txt")
        personal_wordlist_file = self.settings.value("personal_wordlist_file", "personal_wordlist.txt")
        tracker_file = self.settings.value("rescore_tracker_file", "rescore_tracker.txt")
        length_min = int(self.settings.value("length_min", 6))
        length_max = int(self.settings.value("length_max", 10))
        score_min = int(self.settings.value("score_min", 25))
        score_max = int(self.settings.value("score_max", 60))
        disappear_delay = int(self.settings.value("disappear_delay_ms", 200))
        font_family = self.settings.value("main_word_font_family", "Segoe UI")
        font_size = int(self.settings.value("main_word_font_size", 32))

        from PyQt5.QtWidgets import QGroupBox, QFormLayout, QHBoxLayout
        files_group = QGroupBox("Files")
        files_group_layout = QFormLayout(files_group)

        self.master_wordlist_edit = QLineEdit(master_wordlist_file)
        browse_master_btn = QPushButton("Browse...")
        browse_master_btn.clicked.connect(self.browse_master_wordlist)
        master_hbox = QHBoxLayout()
        master_hbox.addWidget(self.master_wordlist_edit)
        master_hbox.addWidget(browse_master_btn)
        files_group_layout.addRow("Master Wordlist File:", master_hbox)

        self.personal_wordlist_edit = QLineEdit(personal_wordlist_file)
        browse_personal_btn = QPushButton("Browse...")
        browse_personal_btn.clicked.connect(self.browse_personal_wordlist)
        personal_hbox = QHBoxLayout()
        personal_hbox.addWidget(self.personal_wordlist_edit)
        personal_hbox.addWidget(browse_personal_btn)
        files_group_layout.addRow("Personal Wordlist File:", personal_hbox)

        self.tracker_edit = QLineEdit(tracker_file)
        browse_tracker_btn = QPushButton("Browse...")
        browse_tracker_btn.clicked.connect(self.browse_tracker)
        tracker_hbox = QHBoxLayout()
        tracker_hbox.addWidget(self.tracker_edit)
        tracker_hbox.addWidget(browse_tracker_btn)
        files_group_layout.addRow("Tracker File:", tracker_hbox)

        filters_group = QGroupBox("Filters")
        filters_group_layout = QFormLayout(filters_group)
        
        self.length_min_spin = QSpinBox()
        self.length_min_spin.setMinimum(1)
        self.length_min_spin.setValue(length_min)

        self.length_max_spin = QSpinBox()
        self.length_max_spin.setMinimum(1)
        self.length_max_spin.setValue(length_max)

        self.score_min_spin = QSpinBox()
        self.score_min_spin.setRange(0, 100)
        self.score_min_spin.setValue(score_min)

        self.score_max_spin = QSpinBox()
        self.score_max_spin.setRange(0, 100)
        self.score_max_spin.setValue(score_max)

        filters_group_layout.addRow("Minimum Word Length:", self.length_min_spin)
        filters_group_layout.addRow("Maximum Word Length:", self.length_max_spin)
        filters_group_layout.addRow("Minimum Score:", self.score_min_spin)
        filters_group_layout.addRow("Maximum Score:", self.score_max_spin)

        timing_group = QGroupBox("Timing")
        timing_group_layout = QFormLayout(timing_group)

        self.delay_spin = QSpinBox()
        self.delay_spin.setRange(100, 10000)
        self.delay_spin.setValue(disappear_delay)
        timing_group_layout.addRow("Disappear Delay (ms):", self.delay_spin)

        font_group = QGroupBox("Main Word Font")
        font_group_layout = QFormLayout(font_group)

        self.font_button = QPushButton("Select Font")
        self.font_button.clicked.connect(self.select_font)
        self.font_family = font_family
        self.font_size = font_size
        font_group_layout.addRow("Font & Size:", self.font_button)

        main_layout.addWidget(files_group)
        main_layout.addWidget(filters_group)
        main_layout.addWidget(timing_group)
        main_layout.addWidget(font_group)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.save_settings)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)

    def browse_master_wordlist(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Select Master Wordlist File", "", "Text Files (*.txt);;All Files (*)")
        if fname:
            self.master_wordlist_edit.setText(fname)

    def browse_personal_wordlist(self):
        fname, _ = QFileDialog.getSaveFileName(self, "Select Personal Wordlist File", "", "Text Files (*.txt);;All Files (*)")
        if fname:
            self.personal_wordlist_edit.setText(fname)

    def browse_tracker(self):
        fname, _ = QFileDialog.getSaveFileName(self, "Select Tracker File", "", "Text Files (*.txt);;All Files (*)")
        if fname:
            self.tracker_edit.setText(fname)

    def select_font(self):
        current_font = QFont(self.font_family, self.font_size)
        font, ok = QFontDialog.getFont(current_font, self, "Select Font for Main Word")
        if ok:
            self.font_family = font.family()
            self.font_size = font.pointSize()

    def save_settings(self):
        self.settings.setValue("master_wordlist_file", self.master_wordlist_edit.text())
        self.settings.setValue("personal_wordlist_file", self.personal_wordlist_edit.text())
        self.settings.setValue("rescore_tracker_file", self.tracker_edit.text())
        self.settings.setValue("length_min", self.length_min_spin.value())
        self.settings.setValue("length_max", self.length_max_spin.value())
        self.settings.setValue("score_min", self.score_min_spin.value())
        self.settings.setValue("score_max", self.score_max_spin.value())
        self.settings.setValue("disappear_delay_ms", self.delay_spin.value())
        self.settings.setValue("main_word_font_family", self.font_family)
        self.settings.setValue("main_word_font_size", self.font_size)
        self.accept()

class RescoreApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Word Rescoring Tool")
        self.setGeometry(100, 100, 1200, 900)

        # Use native Windows style
        QApplication.setStyle("windows")
        QApplication.instance().setFont(QFont("Segoe UI", 9))

        self.settings = QSettings("config.ini", QSettings.IniFormat)
        self.master_wordlist_file = self.settings.value("master_wordlist_file", r"C:\Users\Dennis\OneDrive\XwiWordList.txt")
        self.personal_wordlist_file = self.settings.value("personal_wordlist_file", "personal_wordlist.txt")
        self.rescore_tracker_file = self.settings.value("rescore_tracker_file", "rescore_tracker.txt")
        self.length_min = int(self.settings.value("length_min", 6))
        self.length_max = int(self.settings.value("length_max", 10))
        self.score_min = int(self.settings.value("score_min", 25))
        self.score_max = int(self.settings.value("score_max", 60))
        self.disappear_delay_ms = int(self.settings.value("disappear_delay_ms", 200))
        self.main_word_font_family = self.settings.value("main_word_font_family", "Segoe UI")
        self.main_word_font_size = int(self.settings.value("main_word_font_size", 32))
        self.main_word_font_size = 64  # Larger font

        self.df = self.load_master_wordlist(self.master_wordlist_file)
        if self.df is None:
            QMessageBox.critical(self, "Error", f"Master wordlist file not found: {self.master_wordlist_file}")
            exit()

        self.ensure_file_exists(self.personal_wordlist_file)
        self.ensure_file_exists(self.rescore_tracker_file)

        self.personal_df = self.load_personal_wordlist(self.personal_wordlist_file)
        self.rescored_tracker = self.load_tracker(self.rescore_tracker_file)

        self.apply_personal_scores()

        self.filtered_df = self.apply_filters().sample(frac=1).reset_index(drop=True)

        self.current_index = 0
        self.current_label = None
        self.history_stack = deque(maxlen=5)
        self.ticker_items = deque(maxlen=5)
        self.scoring_in_progress = False

        # Build done_words set once from filtered words and rescored_tracker
        filtered_words_set = set(self.filtered_df['word'])
        self.done_words = set(row['word'] for _, row in self.rescored_tracker.iterrows()
                              if row['rescored'] == 1 and row['word'] in filtered_words_set)

        menubar = self.menuBar()
        settings_action = QAction("Settings", self)
        settings_action.triggered.connect(self.open_settings_dialog)
        menubar.addAction(settings_action)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        central_widget.setFocusPolicy(Qt.StrongFocus)
        central_widget.setFocus()

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(10)

        self.progress_label = QLabel("", self)
        self.progress_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.progress_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setFont(QFont("Segoe UI", 10))
        self.progress_bar.setRange(0, len(self.filtered_df))
        main_layout.addWidget(self.progress_bar)

        self.progress_detail_label = QLabel("", self)
        self.progress_detail_label.setFont(QFont("Segoe UI", 10))
        self.progress_detail_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.progress_detail_label)

        instructions = QLabel("Keybindings:\n[D] Increase | [A] Decrease | [Space] Keep | [U] Undo | [S] Save | [Q] Down 2 | [E] Up 2", self)
        instructions.setFont(QFont("Segoe UI", 10))
        instructions.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(instructions)

        self.word_area = QWidget(self)
        self.word_layout = QVBoxLayout(self.word_area)
        self.word_layout.setContentsMargins(50, 50, 50, 50)
        self.word_layout.setSpacing(10)

        self.current_score_label = QLabel("", self)
        self.current_score_label.setFont(QFont("Segoe UI", 14))
        self.current_score_label.setAlignment(Qt.AlignCenter)

        self.new_score_label = QLabel("", self)
        self.new_score_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        self.new_score_label.setAlignment(Qt.AlignCenter)
        self.new_score_label.setFixedHeight(40)
        self.new_score_label.setText("")

        self.word_layout.addWidget(self.new_score_label, alignment=Qt.AlignCenter)
        self.word_layout.addWidget(self.current_score_label, alignment=Qt.AlignCenter)
        main_layout.addWidget(self.word_area)

        button_layout = QHBoxLayout()
        self.exit_button = QPushButton("Save and Exit", self)
        self.exit_button.clicked.connect(self.export_and_exit)
        self.exit_button.setFocusPolicy(Qt.NoFocus)
        button_layout.addWidget(self.exit_button, alignment=Qt.AlignCenter)
        main_layout.addLayout(button_layout)

        self.ticker_layout = QHBoxLayout()
        main_layout.addLayout(self.ticker_layout)

        self.icons = {
            "increase_1": "icons/icon_increase_1.png",
            "increase_2": "icons/icon_increase_2.png",
            "decrease_1": "icons/icon_decrease_1.png",
            "decrease_2": "icons/icon_decrease_2.png",
            "keep":       "icons/icon_keep.png"
        }

        self.update_progress()
        self.show_next_word()

    def ensure_file_exists(self, path):
        if not os.path.exists(path):
            with open(path, 'w', encoding='utf-8') as f:
                pass

    def load_master_wordlist(self, path):
        if not os.path.exists(path):
            return None
        return pd.read_csv(path, delimiter=';', names=['word', 'score'], engine='python')

    def load_personal_wordlist(self, path):
        if os.path.getsize(path) > 0:
            return pd.read_csv(path, delimiter=';', names=['word','score'], engine='python')
        else:
            return pd.DataFrame(columns=['word','score'])

    def load_tracker(self, path):
        if os.path.getsize(path) > 0:
            return pd.read_csv(path, delimiter=';', names=['word','rescored'], engine='python')
        else:
            return pd.DataFrame(columns=['word','rescored'])

    def apply_personal_scores(self):
        if not self.personal_df.empty:
            personal_dict = dict(zip(self.personal_df['word'], self.personal_df['score']))
            self.df['score'] = self.df.apply(
                lambda row: personal_dict[row['word']] if row['word'] in personal_dict else row['score'], axis=1
            )

    def apply_filters(self):
        filtered = self.df[
            (self.df['score'].between(self.score_min, self.score_max)) &
            (self.df['word'].str.len().between(self.length_min, self.length_max))
        ]
        filtered = filtered.copy()
        filtered['score'] = filtered['score'].apply(map_score_to_bucket)
        return filtered

    def keyPressEvent(self, event):
        if self.current_index >= len(self.filtered_df) or self.scoring_in_progress:
            event.ignore()
            return

        key = event.key()
        if key == Qt.Key_D:
            self.rescore_word("increase")
        elif key == Qt.Key_A:
            self.rescore_word("decrease")
        elif key == Qt.Key_Space:
            self.rescore_word("keep")
        elif key == Qt.Key_U:
            self.undo_action()
        elif key == Qt.Key_S:
            self.save_changes()
        elif key == Qt.Key_Q:
            self.rescore_word("decrease_double")
        elif key == Qt.Key_E:
            self.rescore_word("increase_double")
        else:
            event.ignore()

    def rescore_word(self, action):
        if self.current_index >= len(self.filtered_df):
            return

        self.scoring_in_progress = True
        word_data = self.filtered_df.iloc[self.current_index]
        word = word_data["word"]
        old_score = word_data["score"]

        new_score = self.get_new_score_from_action(old_score, action)
        self.place_word(word, new_score)
        icon_key = self.get_icon_key(old_score, new_score)
        self.history_stack.append((self.current_index, word, old_score, new_score))

        self.ticker_items.appendleft((word, icon_key))
        self.update_ticker()

        self.update_personal_in_memory(word, new_score)
        self.update_tracker_in_memory(word)

        # Add to done_words if not already
        self.done_words.add(word)

        self.show_new_score_flash(old_score, new_score)

        QTimer.singleShot(self.disappear_delay_ms, self.remove_current_word)

    def get_new_score_from_action(self, old_score, action):
        idx = score_buckets.index(old_score)
        if action == "increase":
            new_idx = max(idx-1, 0)
        elif action == "decrease":
            new_idx = min(idx+1, len(score_buckets)-1)
        elif action == "increase_double":
            new_idx = max(idx-2, 0)
        elif action == "decrease_double":
            new_idx = min(idx+2, len(score_buckets)-1)
        else:
            new_idx = idx
        return score_buckets[new_idx]

    def remove_current_word(self):
        if self.current_label and self.current_label.parent():
            layout = self.current_label.parent().layout()
            layout.removeWidget(self.current_label)
            self.current_label.deleteLater()
        self.current_label = None
        self.current_index += 1
        self.show_next_word()

    def show_next_word(self):
        if self.current_index < len(self.filtered_df):
            word_data = self.filtered_df.iloc[self.current_index]
            word = word_data["word"]
            score = map_score_to_bucket(word_data["score"])
            self.filtered_df.at[self.current_index, "score"] = score
            self.place_word(word, score)
            self.update_progress()
            self.scoring_in_progress = False
        else:
            self.update_progress()
            QMessageBox.information(self, "Done", "All words have been rescored!")
            self.current_label = None
            self.scoring_in_progress = False

    def place_word(self, word, score):
        self.current_score_label.setText(f"Current Score: {score}")

        if self.current_label and self.current_label.parent():
            old_layout = self.current_label.parent().layout()
            old_layout.removeWidget(self.current_label)
            self.current_label.deleteLater()

        main_word_font = QFont(self.main_word_font_family, self.main_word_font_size, QFont.Bold)

        self.current_label = QLabel(word)
        self.current_label.setFont(main_word_font)
        self.current_label.setAlignment(Qt.AlignCenter)
        self.current_label.setMinimumHeight(150)
        self.current_label.setStyleSheet(f"""
            color: #000000;
            border: 3px solid {bucket_colors[score]};
            border-radius: 10px;
            padding: 20px;
            background-color: #ffffff;
        """)

        self.word_layout.insertWidget(0, self.current_label, alignment=Qt.AlignCenter)

    def show_new_score_flash(self, old_score, new_score):
        diff = new_score - old_score
        if diff > 0:
            text = f"New Score: {new_score}"
            color = "#2ecc71"
        elif diff < 0:
            text = f"New Score: {new_score}"
            color = "#e74c3c"
        else:
            text = f"Score: {new_score}"
            color = "#000000"

        self.new_score_label.setText(text)
        self.new_score_label.setStyleSheet(f"color: {color};")
        QTimer.singleShot(200, lambda: self.new_score_label.setText(""))

    def get_icon_key(self, old_score, new_score):
        if new_score == old_score:
            return "keep"
        old_idx = score_buckets.index(old_score)
        new_idx = score_buckets.index(new_score)
        diff = old_idx - new_idx
        if diff > 0:
            return "increase_2" if diff > 1 else "increase_1"
        elif diff < 0:
            return "decrease_2" if diff < -1 else "decrease_1"
        else:
            return "keep"

    def update_progress(self):
        total_count = len(self.filtered_df)
        done_count = len(self.done_words)  # use in-memory done_words set
        progress = 0.0
        if total_count > 0:
            progress = (done_count / total_count) * 100
        self.progress_label.setText(f"{progress:.1f}% Completed")
        self.progress_bar.setMaximum(total_count)
        self.progress_bar.setValue(done_count)
        self.progress_detail_label.setText(f"{done_count} out of {total_count} complete")

    def update_ticker(self):
        while self.ticker_layout.count() > 0:
            item = self.ticker_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        for word, icon_key in self.ticker_items:
            ticker_item_widget = QFrame()
            ticker_item_layout = QVBoxLayout(ticker_item_widget)
            ticker_item_layout.setContentsMargins(5,5,5,5)
            ticker_item_layout.setSpacing(2)

            icon_label = QLabel()
            icon_label.setFixedSize(48, 48)
            icon_label.setScaledContents(True)
            icon_path = self.icons.get(icon_key, None)
            if icon_path and os.path.exists(icon_path):
                pixmap = QPixmap(icon_path)
                if not pixmap.isNull():
                    icon_label.setPixmap(pixmap)
                else:
                    icon_label.setText(icon_key)
            else:
                icon_label.setText(icon_key)
            ticker_item_layout.addWidget(icon_label, alignment=Qt.AlignCenter)

            w_label = QLabel(word)
            w_label.setFont(QFont("Segoe UI", 10))
            w_label.setAlignment(Qt.AlignCenter)
            ticker_item_layout.addWidget(w_label, alignment=Qt.AlignCenter)

            self.ticker_layout.addWidget(ticker_item_widget)

    def update_personal_in_memory(self, word, score):
        if word in self.personal_df['word'].values:
            self.personal_df.loc[self.personal_df['word'] == word, 'score'] = score
        else:
            new_row = pd.DataFrame({'word': [word], 'score': [score]})
            self.personal_df = pd.concat([self.personal_df, new_row], ignore_index=True)

    def update_tracker_in_memory(self, word):
        if word in self.rescored_tracker['word'].values:
            self.rescored_tracker.loc[self.rescored_tracker['word'] == word, 'rescored'] = 1
        else:
            new_row = pd.DataFrame({'word': [word], 'rescored': [1]})
            self.rescored_tracker = pd.concat([self.rescored_tracker, new_row], ignore_index=True)

    def write_personal_wordlist_to_disk(self):
        self.personal_df.to_csv(self.personal_wordlist_file, sep=';', index=False, header=False)

    def write_tracker_to_disk(self):
        self.rescored_tracker.to_csv(self.rescore_tracker_file, sep=';', index=False, header=False)

    def save_changes(self):
        self.write_personal_wordlist_to_disk()
        self.write_tracker_to_disk()
        QMessageBox.information(self, "Saved", "Changes saved successfully.")

    def undo_action(self):
        if self.scoring_in_progress:
            return
        if not self.history_stack:
            QMessageBox.information(self, "Undo", "No more undo actions available.")
            return

        last_index, word, old_score, new_score = self.history_stack.pop()

        for i, (w, icon_key) in enumerate(self.ticker_items):
            if w == word:
                self.ticker_items.remove((w, icon_key))
                break
        self.update_ticker()

        self.current_index = last_index
        old_score = map_score_to_bucket(old_score)
        self.filtered_df.at[self.current_index, "score"] = old_score

        # Revert personal in memory
        if word in self.personal_df['word'].values:
            self.personal_df.loc[self.personal_df['word'] == word, 'score'] = old_score

        # Revert tracker in memory
        if word in self.rescored_tracker['word'].values:
            self.rescored_tracker.loc[self.rescored_tracker['word'] == word, 'rescored'] = 0

        # Remove from done_words
        if word in self.done_words:
            self.done_words.remove(word)

        self.show_next_word()

    def open_settings_dialog(self):
        dialog = SettingsDialog(self.settings, self)
        if dialog.exec_():
            self.master_wordlist_file = self.settings.value("master_wordlist_file", r"C:\Users\Dennis\OneDrive\XwiWordList.txt")
            self.personal_wordlist_file = self.settings.value("personal_wordlist_file", "personal_wordlist.txt")
            self.rescore_tracker_file = self.settings.value("rescore_tracker_file", "rescore_tracker.txt")
            self.length_min = int(self.settings.value("length_min", 6))
            self.length_max = int(self.settings.value("length_max", 10))
            self.score_min = int(self.settings.value("score_min", 25))
            self.score_max = int(self.settings.value("score_max", 60))
            self.disappear_delay_ms = int(self.settings.value("disappear_delay_ms", 200))
            self.main_word_font_family = self.settings.value("main_word_font_family", "Segoe UI")
            self.main_word_font_size = int(self.settings.value("main_word_font_size", 32))
            self.main_word_font_size = 64

            self.df = self.load_master_wordlist(self.master_wordlist_file)
            if self.df is None:
                QMessageBox.critical(self, "Error", f"Master wordlist file not found: {self.master_wordlist_file}")
                return

            self.ensure_file_exists(self.personal_wordlist_file)
            self.ensure_file_exists(self.rescore_tracker_file)

            self.personal_df = self.load_personal_wordlist(self.personal_wordlist_file)
            self.rescored_tracker = self.load_tracker(self.rescore_tracker_file)
            self.apply_personal_scores()
            self.filtered_df = self.apply_filters().sample(frac=1).reset_index(drop=True)
            self.current_index = 0
            self.history_stack.clear()
            self.ticker_items.clear()

            # Rebuild done_words based on new filtering
            filtered_words_set = set(self.filtered_df['word'])
            self.done_words = set(row['word'] for _, row in self.rescored_tracker.iterrows()
                                  if row['rescored'] == 1 and row['word'] in filtered_words_set)

            self.update_ticker()
            self.show_next_word()
            self.update_progress()
            QMessageBox.information(self, "Settings", "Settings updated.")

    def export_and_exit(self):
        self.write_personal_wordlist_to_disk()
        self.write_tracker_to_disk()
        QMessageBox.information(self, "Exit", "Progress saved.")
        self.close()

if __name__ == "__main__":
    app = QApplication([])
    QApplication.setStyle("windows")
    app.setFont(QFont("Segoe UI", 9))
    window = RescoreApp()
    window.show()
    app.exec_()
