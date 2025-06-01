import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import textwrap
import logging
import re
import csv
import configparser
import os

logging.basicConfig(level=logging.DEBUG, format='%(levelname)s:%(message)s')
JAPANESE_CHAR_PATTERN = re.compile(r'[\u3040-\u30ff\u4e00-\u9faf\uff66-\uff9f]')
THEME_FILE = "theme.ini"

class CSVTranslationTool:
    def __init__(self, root):
        self.root = root
        self.root.title("ButterCSV-Editor")
        self.root.geometry("900x700")
        self.data = None
        self.deduped_map = {}
        self.reverse_map = {}
        self.text_widgets = {}
        self.entries_frame = None
        self.list_widget = None
        self.canvas_window = None
        self.settings_frame = None

        self.current_page = 0
        self.entries_per_page = 50
        self.wrap_limit = 28
        self.max_lines = 3
        self.ordered_keys = []
        self.list_mode = False
        self.min_duplicates_filter = 0
        self.sort_descending = True
        self.temp_save_path = "_autosave_translation_cache.csv"
        self.current_view = "editor"

        self.style = ttk.Style()
        self.theme = configparser.ConfigParser()
        self.load_theme()

        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)
        self.setup_ui()
        self.root.bind("<Control-s>", lambda e: self.manual_save())

    def load_theme(self):
        if not os.path.exists(THEME_FILE):
            self.theme['colors'] = {
                'bg': '#0d1b2a',
                'fg': '#e9c46a',
                'frbg': '#0d1b2a',
                'topbar_bg': '#0d1b2a',
                'navbar_bg': '#0d1b2a',
                'button_bg': '#1b263b',
                'button_fg': '#eceff4',
                'entry_bg': '#1b263b',
                'entry_fg': '#eceff4',
                'list_bg': '#0d1b2a',
                'list_fg': '#e9c46a',
                'list_label_fg': '#778da9',
                'highlight': '#e9c46a',
                'accent': '#e9c46a'
            }
            self.theme['fonts'] = {
                'base': 'Calibri',
                'size': '12',
                'mono': 'Fira Code'
            }
            with open(THEME_FILE, 'w') as configfile:
                self.theme.write(configfile)
        else:
            self.theme.read(THEME_FILE)

        font_base = (self.theme['fonts'].get('base', 'Segoe UI'), int(self.theme['fonts'].get('size', 10)))
        font_mono = (self.theme['fonts'].get('mono', 'Consolas'), int(self.theme['fonts'].get('size', 10)))

        self.style.theme_use('clam')
        self.style.configure("TButton",
                             font=font_base,
                             padding=5,
                             background=self.theme['colors'].get('button_bg'),
                             foreground=self.theme['colors'].get('button_fg'))
        self.style.map("TButton",
                       background=[("!disabled", self.theme['colors'].get('button_bg'))],
                       foreground=[("!disabled", self.theme['colors'].get('button_fg'))])

        self.style.configure("TLabel",
                             font=font_base,
                             background=self.theme['colors'].get('bg'),
                             foreground=self.theme['colors'].get('fg'))

        self.style.configure("Accent.TLabel",
                             background=self.theme['colors'].get('bg'),
                             foreground=self.theme['colors'].get('accent'))

        self.style.configure("Highlight.TLabel",
                             background=self.theme['colors'].get('bg'),
                             foreground=self.theme['colors'].get('highlight'))

        self.style.configure("TEntry",
                             font=font_mono,
                             fieldbackground=self.theme['colors'].get('entry_bg'),
                             background=self.theme['colors'].get('entry_bg'),
                             foreground=self.theme['colors'].get('entry_fg'),
                             insertcolor=self.theme['colors'].get('entry_fg'))

        self.style.configure("TFrame",
                             background=self.theme['colors'].get('frbg'))

        self.style.configure("TopBar.TFrame",
                             background=self.theme['colors'].get('topbar_bg'))

        self.style.configure("NavBar.TFrame",
                             background=self.theme['colors'].get('navbar_bg'))

    def setup_ui(self):
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        options_menu = tk.Menu(self.menu_bar, tearoff=0)
        options_menu.add_command(label="Reload Theme", command=self.reload_theme)
        options_menu.add_command(label="Settings...", command=self.show_settings_view)
        self.menu_bar.add_cascade(label="Options", menu=options_menu)

        self.top_frame = ttk.Frame(self.root, style="TopBar.TFrame")
        self.top_frame.pack(side=tk.TOP, fill=tk.X, pady=10)

        self.load_button = ttk.Button(self.top_frame, text="Load CSV", command=self.load_csv)
        self.save_button = ttk.Button(self.top_frame, text="Save & Rebuild CSV", command=self.save_and_rebuild)
        self.fallback_save_button = ttk.Button(self.top_frame, text="Manual Save", command=self.manual_save)
        self.toggle_list_button = ttk.Button(self.top_frame, text="Toggle List Mode", command=self.toggle_list_mode)

        self.filter_label = ttk.Label(self.top_frame, text="Min Duplicates:", style="Highlight.TLabel")
        self.filter_entry = ttk.Entry(self.top_frame, width=5)
        self.sort_button = ttk.Button(self.top_frame, text="Sort: Desc", command=self.toggle_sort_order)
        self.apply_filter_button = ttk.Button(self.top_frame, text="Apply Filter", command=self.apply_filter)

        for widget in [self.load_button, self.save_button, self.fallback_save_button,
                       self.toggle_list_button, self.filter_label, self.filter_entry,
                       self.sort_button, self.apply_filter_button]:
            widget.pack(side=tk.LEFT, padx=5)

        self.filter_entry.insert(0, "0")

        self.nav_frame = ttk.Frame(self.root, style="NavBar.TFrame")
        self.nav_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.prev_button = ttk.Button(self.nav_frame, text="Previous", command=self.prev_page)
        self.page_label = ttk.Label(self.nav_frame, text="Page 1", style="Accent.TLabel")
        self.next_button = ttk.Button(self.nav_frame, text="Next", command=self.next_page)
        self.prev_button.pack(side=tk.LEFT, padx=10, pady=5)
        self.page_label.pack(side=tk.LEFT, padx=10)
        self.page_seek_entry = ttk.Entry(self.nav_frame, width=5)
        self.page_seek_entry.pack(side=tk.LEFT, padx=5)
        self.page_seek_entry.bind("<Return>", self.seek_page)
        self.next_button.pack(side=tk.LEFT, padx=10, pady=5)

        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(self.main_frame, bg=self.theme['colors'].get('bg'), highlightthickness=0)
        self.scroll_y = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scroll_y.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.refresh_page()

    def refresh_page(self):
        start = self.current_page * self.entries_per_page
        end = start + self.entries_per_page
        page_keys = self.ordered_keys[start:end]
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.clear_main_canvas()

        if not self.ordered_keys:
            self.entries_frame = ttk.Frame(self.canvas)
            self.canvas_window = self.canvas.create_window((0, 0), window=self.entries_frame, anchor="nw")

            ttk.Label(
                self.entries_frame,
                text="Welcome to ButterCSV-Editor!",
                justify="center",
                font=(self.theme['fonts'].get('base'), 18, 'bold'),
                foreground=self.theme['colors'].get('list_fg')
            ).pack(expand=True, fill=tk.BOTH, padx=270, pady=(70, 20))

            ttk.Label(
                self.entries_frame,
                text="Load a CSV file to get started.\n\nSaves/'pre-rebuild' are cached in the same dir as the script.\n\nUse 'List Mode' for mass edit and copy.\n\nUse the 'Options' menu to reload the theme or adjust settings.\n\nYou can stylize the theme in the 'theme.ini' same dir as script.\n\nKeyboard Shortcuts:\n\nCTRL+S - Manual Save\n\nCTRL+C - Copy\n\nCTRL+V - Paste",
                justify="center",
                font=(self.theme['fonts'].get('base'), 14),
                foreground=self.theme['colors'].get('entry_fg')
            ).pack(expand=True, fill=tk.BOTH, padx=185, pady=(10, 0))

            # Disable scrolling
            self.canvas.unbind_all("<MouseWheel>")
            self.canvas.config(scrollregion=(0, 0, 0, 0))
            return

        if self.list_mode:
            self.canvas.pack_forget()
            self.scroll_y.pack_forget()

            self.list_widget = tk.Text(self.main_frame,
                                       bg=self.theme['colors'].get('list_bg'),
                                       fg=self.theme['colors'].get('list_fg'),
                                       insertbackground=self.theme['colors'].get('list_fg'),
                                       wrap=tk.WORD,
                                       font=(self.theme['fonts'].get('mono'), int(self.theme['fonts'].get('size'))),
                                       borderwidth=0,
                                       highlightthickness=0)
            self.list_widget.pack(fill=tk.BOTH, expand=True, padx=(80, 180), pady=(20, 20))
            self.list_widget.config(state=tk.NORMAL)
            self.list_widget.delete("1.0", tk.END)
            self.protected_label_lines = []

            for i, key in enumerate(page_keys):
                content = self.deduped_map[key]
                issues = self.check_text_limits(content)
                flag = f" ⚠ ({'; '.join(str(x) for x in issues)})" if issues else ""
                true_index = self.reverse_map[key][0] + 1
                label = f"____Entry {true_index}{flag}:\n"

                label_index = self.list_widget.index(tk.INSERT)
                entry_tag = f"label_{i}"

                self.list_widget.insert(tk.END, label)
                label_end_index = self.list_widget.index(tk.INSERT)

                self.list_widget.tag_add(entry_tag, label_index, label_end_index)
                self.list_widget.tag_config(entry_tag,
                                            foreground=self.theme['colors'].get('list_label_fg', '#89c2d9'),
                                            font=(self.theme['fonts'].get('mono'),
                                                  int(self.theme['fonts'].get('size')), 'bold'))

                self.protected_label_lines.append(entry_tag)

                self.list_widget.insert(tk.END, content.strip() + "\n\n")

            self.list_widget.bind("<Key>", self._block_label_edit)
            self.list_widget.bind("<BackSpace>", self._block_label_edit)
            self.list_widget.bind("<Delete>", self._block_label_edit)
            self.list_widget.bind("<<Modified>>", self.on_list_mode_change)
            self.bind_clipboard_shortcuts(self.list_widget)
            self.list_widget.edit_modified(False)
            self.add_context_menu(self.list_widget)

        else:
            self.entries_frame = ttk.Frame(self.canvas)
            self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
            self.canvas_window = self.canvas.create_window((0, 0), window=self.entries_frame, anchor="nw")

            for i, key in enumerate(page_keys):
                content = self.deduped_map[key]
                issues = self.check_text_limits(content)
                flag = f" ⚠ ({'; '.join(str(x) for x in issues)})" if issues else ""
                true_index = self.reverse_map[key][0] + 1
                label_text = f"Entry {true_index} ({len(self.reverse_map[key])}x){flag}:"
                lbl = ttk.Label(self.entries_frame, text=label_text)
                lbl.pack(anchor='w', padx=5, pady=(5, 0))

                txt = tk.Text(self.entries_frame,
                              height=4,
                              width=100,
                              bg=self.theme['colors'].get('entry_bg'),
                              fg=self.theme['colors'].get('entry_fg'),
                              insertbackground=self.theme['colors'].get('entry_fg'),
                              wrap=tk.WORD,
                              font=(self.theme['fonts'].get('mono'), int(self.theme['fonts'].get('size'))))
                txt.insert(tk.END, content.strip())
                txt.pack(padx=5, pady=(0, 10), fill=tk.X)
                txt.bind("<<Modified>>", lambda e, k=key, l=lbl, t=txt: self.on_text_change(e, k, l))
                txt.edit_modified(False)

                self.text_widgets[key] = txt
                self.add_context_menu(txt)
                self.bind_clipboard_shortcuts(txt)

        self.canvas.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        pages = (len(self.ordered_keys) - 1) // self.entries_per_page + 1
        self.page_label.config(text=f"Page {self.current_page+1} of {pages}")

    def on_text_change(self, event, key, label):
        widget = event.widget
        widget.edit_modified(False)
        content = widget.get("1.0", tk.END).strip()
        self.deduped_map[key] = content
        issues = self.check_text_limits(content)
        if issues:
            label.config(text=f"{label.cget('text').split(' ⚠')[0]} ⚠ ({'; '.join(issues)})")
        else:
            label.config(text=label.cget('text').split(' ⚠')[0])

    def on_list_mode_change(self, event=None):
        if not self.list_mode or not self.list_widget:
            return

        if not self.list_widget.edit_modified():
            return
        self.list_widget.edit_modified(False)

        try:
            self.list_widget.unbind("<<Modified>>")

            index = self.list_widget.index("insert")
            current_line = int(index.split('.')[0])

            target_idx = None
            label_line_indices = []
            for tag in self.protected_label_lines:
                tag_indices = self.list_widget.tag_ranges(tag)
                if tag_indices:
                    label_line = int(str(tag_indices[0]).split('.')[0])
                    label_line_indices.append((label_line, tag))

            label_line_indices.sort()
            for i, (label_line, tag) in enumerate(label_line_indices):
                next_line = label_line_indices[i + 1][0] if i + 1 < len(label_line_indices) else float('inf')
                if current_line > label_line and current_line < next_line:
                    target_idx = i
                    break

            if target_idx is None:
                return

            key = self.ordered_keys[self.current_page * self.entries_per_page + target_idx]
            tag = self.protected_label_lines[target_idx]

            tag_indices = self.list_widget.tag_ranges(tag)
            if not tag_indices:
                return

            label_line = int(str(tag_indices[0]).split('.')[0])
            body_start = f"{label_line + 1}.0"
            body_end = f"{label_line_indices[target_idx + 1][0]}.0" if target_idx + 1 < len(label_line_indices) else tk.END
            body = self.list_widget.get(body_start, body_end).strip()

            self.deduped_map[key] = body
            issues = self.check_text_limits(body)
            flag = f" ⚠ ({'; '.join(issues)})" if issues else ""
            true_index = self.reverse_map[key][0] + 1
            new_label = f"____Entry {true_index}{flag}:"

            self.list_widget.delete(f"{label_line}.0", f"{label_line}.end")
            self.list_widget.insert(f"{label_line}.0", new_label)
            self.list_widget.tag_remove(tag, f"{label_line}.0", f"{label_line}.end")
            self.list_widget.tag_add(tag, f"{label_line}.0", f"{label_line}.end")

        except Exception as e:
            logging.error(f"on_list_mode_change failed: {e}")

        finally:
            self.list_widget.edit_modified(False)
            self.list_widget.bind("<<Modified>>", self.on_list_mode_change)

    def check_text_limits(self, text):
        issues = []
        lines = text.splitlines()

        def clean_color_codes(line):
            return re.sub(r'‾C[0-9A-F]{2}', '', line)

        for i, line in enumerate(lines, start=1):
            cleaned = clean_color_codes(line.strip())
            if len(cleaned) > self.wrap_limit:
                issues.append(f"line {i} > {self.wrap_limit} chars")

        if len(lines) > self.max_lines:
            issues.append(f"{len(lines)} lines > {self.max_lines}")

        # Check for unclosed color codes
        color_codes = re.findall(r'‾C([0-9A-F]{2})', text)
        opened = [code for code in color_codes if code != '00']
        closed = color_codes.count('00')
        if closed < len(opened):
            issues.append("unmatched color tag(s): " + ", ".join(sorted(set(f"‾C{code}" for code in opened))))

        return issues

    def protect_label(self, index_start, index_end):
        def block_edit(event):
            index = self.list_widget.index("insert")
            if self.list_widget.compare(index, ">=", index_start) and self.list_widget.compare(index, "<", index_end):
                return "break"
            return None

        self.list_widget.tag_bind("protected_label", "<Key>", block_edit)
        self.list_widget.tag_bind("protected_label", "<BackSpace>", block_edit)
        self.list_widget.tag_bind("protected_label", "<Delete>", block_edit)
        self.list_widget.tag_bind("protected_label", "<Return>", block_edit)

    def show_settings_view(self):
        self.save_current_page()
        self.clear_main_canvas()
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.current_view = "settings"

        self.settings_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.settings_frame, anchor="nw")

        ttk.Label(self.settings_frame, text="Entries Per Page:").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.page_count_entry = ttk.Entry(self.settings_frame, width=5)
        self.page_count_entry.insert(0, str(self.entries_per_page))
        self.page_count_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(self.settings_frame, text="Max Characters per Line:").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.char_wrap_entry = ttk.Entry(self.settings_frame, width=5)
        self.char_wrap_entry.insert(0, str(self.wrap_limit))
        self.char_wrap_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(self.settings_frame, text="Max Lines per Entry:").grid(row=2, column=0, sticky='e', padx=5, pady=5)
        self.max_lines_entry = ttk.Entry(self.settings_frame, width=5)
        self.max_lines_entry.insert(0, str(self.max_lines))
        self.max_lines_entry.grid(row=2, column=1, padx=5, pady=5)

        btn_frame = ttk.Frame(self.settings_frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=15)

        apply_btn = ttk.Button(btn_frame, text="Apply", command=self.apply_settings)
        save_btn = ttk.Button(btn_frame, text="Save and Return", command=self.save_settings_and_return)

        apply_btn.pack(side=tk.LEFT, padx=10)
        save_btn.pack(side=tk.LEFT, padx=10)

        self.canvas.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def apply_settings(self):
        try:
            self.entries_per_page = max(1, int(self.page_count_entry.get()))
        except ValueError:
            pass
        try:
            self.wrap_limit = max(1, int(self.char_wrap_entry.get()))
        except ValueError:
            self.wrap_limit = 28
        try:
            self.max_lines = max(1, int(self.max_lines_entry.get()))
        except ValueError:
            self.max_lines = 3

    def save_settings_and_return(self):
        self.apply_settings()
        self.refresh_page()

    def clear_main_canvas(self):
        if self.entries_frame:
            self.entries_frame.destroy()
            self.entries_frame = None
        if self.list_widget:
            self.list_widget.destroy()
            self.list_widget = None
        if self.settings_frame:
            self.settings_frame.destroy()
            self.settings_frame = None
        if self.canvas_window:
            self.canvas.delete(self.canvas_window)
            self.canvas_window = None
        self.text_widgets.clear()

    def seek_page(self, event=None):
        try:
            page = int(self.page_seek_entry.get()) - 1
            max_page = (len(self.ordered_keys) - 1) // self.entries_per_page
            if 0 <= page <= max_page:
                if self.list_mode:
                    page_keys = self.ordered_keys[self.current_page * self.entries_per_page:
                                                  (self.current_page + 1) * self.entries_per_page]
                    self.save_current_page(page_keys)
                else:
                    self.save_current_page()
                self.current_page = page
                self.refresh_page()
        except ValueError:
            pass

    def _block_label_edit(self, event):
        if not self.list_widget or not hasattr(self, 'protected_label_lines'):
            return

        try:
            index = self.list_widget.index("insert")
            for tag in self.protected_label_lines:
                if tag in self.list_widget.tag_names(index):
                    return "break"
        except Exception:
            return

    def toggle_sort_order(self):
        self.sort_descending = not self.sort_descending
        self.sort_button.config(text="Sort: Desc" if self.sort_descending else "Sort: Asc")
        self.apply_filter()

    def toggle_list_mode(self):
        if self.list_mode:
            page_keys = self.ordered_keys[self.current_page * self.entries_per_page:
                                          (self.current_page + 1) * self.entries_per_page]
            self.save_current_page(page_keys)
        else:
            self.save_current_page()

        self.list_mode = not self.list_mode
        self.refresh_page()

    def next_page(self):
        if self.list_mode:
            page_keys = self.ordered_keys[self.current_page * self.entries_per_page:
                                          (self.current_page + 1) * self.entries_per_page]
            self.save_current_page(page_keys)
        else:
            self.save_current_page()

        if (self.current_page + 1) * self.entries_per_page < len(self.ordered_keys):
            self.current_page += 1
            self.refresh_page()

    def prev_page(self):
        if self.list_mode:
            page_keys = self.ordered_keys[self.current_page * self.entries_per_page:
                                          (self.current_page + 1) * self.entries_per_page]
            self.save_current_page(page_keys)
        else:
            self.save_current_page()

        if self.current_page > 0:
            self.current_page -= 1
            self.refresh_page()

    def save_current_page(self, page_keys=None):
        try:
            if self.list_mode and self.list_widget and page_keys:
                content = self.list_widget.get("1.0", tk.END).strip()
                pattern = r'^____Entry \d+(?: ⚠ \([^)]+\))?:$'
                split_entries = re.split(pattern, content, flags=re.MULTILINE)[1:]
                matches = re.findall(pattern, content, flags=re.MULTILINE)

                if len(split_entries) == len(page_keys):
                    for key, val in zip(page_keys, split_entries):
                        cleaned = val.strip()
                        self.deduped_map[key] = cleaned
                        issues = self.check_text_limits(cleaned)
                        if issues:
                            logging.warning(f"Entry {key} issues: {issues}")
                else:
                    logging.warning("List mode save aborted: mismatch in entry count.")
                    logging.warning(f"Expected: {len(page_keys)}, Got: {len(split_entries)}")
                    logging.debug(f"Raw content preview:\n{content[:500]}")
            else:
                for key, widget in self.text_widgets.items():
                    self.deduped_map[key] = widget.get("1.0", tk.END).strip()

            self.autosave_temp()

        except Exception as e:
            logging.error(f"Error during save_current_page: {e}")

    def autosave_temp(self):
        try:
            with open(self.temp_save_path, 'w', encoding='utf-8', newline='') as f:
                csv.writer(f).writerows(self.deduped_map.items())
        except Exception as e:
            logging.error(f"Autosave failed: {e}")

    def manual_save(self):
        self.save_current_page()
        messagebox.showinfo("Manual Save", f"Saved to {self.temp_save_path}")

    def load_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not path:
            return
        try:
            with open(path, newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                self.data = list(reader)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read CSV: {e}")
            return

        self.deduped_map.clear()
        self.reverse_map.clear()

        DUMMY_KEYWORDS = {"dummy", "ダミー", "ダミー。", "※開発用"}

        for idx, row in enumerate(self.data):
            target = row.get('target', '').strip()
            if not target or target in DUMMY_KEYWORDS:
                continue

            if JAPANESE_CHAR_PATTERN.search(target) or re.search(r'[a-zA-Z0-9]', target):
                self.deduped_map.setdefault(target, target)
                self.reverse_map.setdefault(target, []).append(idx)

        self.apply_filter()

    def apply_filter(self):
        try:
            self.min_duplicates_filter = int(self.filter_entry.get())
        except ValueError:
            self.min_duplicates_filter = 0
        self.ordered_keys = [k for k in self.deduped_map if len(self.reverse_map[k]) >= self.min_duplicates_filter]
        self.ordered_keys.sort(key=lambda k: len(self.reverse_map[k]), reverse=self.sort_descending)
        self.current_page = 0
        self.refresh_page()

    def wrap_text(self, text):
        lines = []
        for paragraph in text.split("\n"):
            wrapped_lines = textwrap.wrap(paragraph,
                                          width=self.wrap_limit,
                                          break_long_words=False,
                                          break_on_hyphens=False)
            lines.extend(wrapped_lines if wrapped_lines else [""])
        return (lines, len(lines) > self.max_lines)

    def save_and_rebuild(self):
        self.save_current_page()
        lines = []
        warnings = []

        DUMMY_KEYWORDS = {"dummy", "ダミー", "ダミー。", "※開発用"}

        for idx, row in enumerate(self.data):
            loc, src, tgt = row['location'], row['source'], row['target']
            stripped = tgt.strip()

            if stripped in self.deduped_map:
                new_txt = self.deduped_map[stripped]
                wrapped, over = self.wrap_text(new_txt)
                final = "\n".join(wrapped)
                quoted_final = f'"{final.replace(chr(34), chr(34)*2)}"'
                lines.append([loc, src, quoted_final])

                if over:
                    warnings.append(f"{loc},{src} exceeded line limit with {len(wrapped)} lines")

            else:
                if ',' in tgt or '\n' in tgt or '"' in tgt:
                    tgt = f'"{tgt.replace(chr(34), chr(34)*2)}"'
                lines.append([loc, src, tgt])

        path = filedialog.asksaveasfilename(defaultextension=".csv")
        if not path:
            return

        try:
            with open(path, 'w', newline='', encoding='utf-8') as f:
                f.write("location,source,target\n")
                for row in lines:
                    f.write(",".join(row) + "\n")
            if warnings:
                messagebox.showwarning("Line Limit Warnings", "\n".join(warnings))
            else:
                messagebox.showinfo("Saved", "CSV rebuilt and saved.")
        except Exception as e:
            messagebox.showerror("Error", f"Save failed: {e}")

    def reload_theme(self):
        self.load_theme()
        self.refresh_page()

    def on_exit(self):
        self.save_current_page()
        self.root.destroy()

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def bind_clipboard_shortcuts(self, widget):
        def on_paste(event):
            # Prevent double-paste
            if event.state & 0x0004:
                return "break"

        widget.bind("<Control-c>", lambda e: widget.event_generate("<<Copy>>"), add=True)
        widget.bind("<Control-v>", lambda e: widget.event_generate("<<Paste>>"), add=True)
        widget.bind("<Control-v>", on_paste, add=True)
        widget.bind("<Control-x>", lambda e: widget.event_generate("<<Cut>>"), add=True)

    def add_context_menu(self, widget):
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="Copy", command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label="Paste", command=lambda: widget.event_generate("<<Paste>>"))
        menu.add_command(label="Cut", command=lambda: widget.event_generate("<<Cut>>"))
        widget.bind("<Button-3>", lambda e: menu.tk_popup(e.x_root, e.y_root))

        self.bind_clipboard_shortcuts(widget)

if __name__ == "__main__":
    root = tk.Tk()
    app = CSVTranslationTool(root)
    root.mainloop()