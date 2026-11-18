# Taken from Linpad++ (GPL-3.0) for the LinBASIC Studio IDE.
"""Persistent preferences (the equivalent of Notepad++'s config.xml GUIConfigs)."""
import json
import os

from . import paths

DEFAULTS = {
    # LinBASIC Studio
    "codepage": "",                    # code page of VB6 source files; "" = from the locale (cp1253 for Greek)
    "msdn_path": "",                   # the user's own MSDN Library (VB6), for Help
    "auto_list_members": True,         # Tools > Options > Editor of VB6
    "auto_quick_info": True,
    "recent_projects": [],
    "mdi_tabbed": True,                # code windows as tabs (False: MDI windows like VB6)
    "make_target": "linux",            # File > Make: linux | win32 | win64
    "error_trapping": 1,               # Tools > Options > General: 0 Break on All Errors, 1 Break in Class
                                       # Module, 2 Break on Unhandled Errors
    "show_grid": True,                 # Tools > Options > General: Form Grid Settings (VB6: 120 x 120 twips)
    "grid_width": 120,
    "grid_height": 120,
    "align_grid": True,                # Align Controls to Grid
    "layout_guides": True,             # Form Layout: Resolution Guides (VB6: on, DEV-8/fd1 stage 20)
    "dockable": {},                     # Tools > Options > Docking: {"immediate": True, ... "objbrowser": False}
    "layout_dockable": True,           # Form Layout: Dockable (off: a window among the code and form windows)
    "addin_reseditor_startup": False,  # Add-In Manager: VB 6 Resource Editor, Load on Startup
    # General / UI
    "localization": "english.xml",
    "toolbar_hidden": False,
    "toolbar_icons": "small",          # small | large
    "statusbar_hidden": False,
    "menubar_hidden": False,
    "tabbar_hidden": False,
    "tab_close_button": True,
    "tab_dblclick_close": False,
    "tab_exit_on_last": False,
    "tab_lock": False,
    "tab_max_len": 0,
    "tab_reduce": False,
    "tab_multiline": False,
    "add_mdi_nodialog": False,         # Project > Add MDI Form: "Don't show this dialog in the future"
    "custom_colors": [],               # View > Color Palette: the last row (Define Colors...), 0xRRGGBB
    "toolbox_tabs": [],                # the Toolbox: the user's tabs [{"name": "MyTab", "tools": ["Timer", ...]}]
    "dark_mode": "system",             # light | dark | system (Tools > Options > Theme; system: as the OS is)
    "theme": "",                       # "" = stylers.xml
    "light_theme": "",                 # theme remembered for light mode
    "dark_theme": "DarkModeDefault.xml",  # theme remembered for dark mode
    # Editing
    "caretline_mode": 1,               # 0 none, 1 background, 2 frame
    "caretline_frame_width": 1,
    "caret_width": 1,
    "caret_blink": 600,
    "wrap_mode": 0,                    # 0 default, 1 aligned, 2 indent
    "virtual_space": False,
    "scroll_beyond_last": False,
    "fold_style": "box",               # simple | arrow | circle | box | none
    "line_numbers": True,
    "line_numbers_dynamic": True,
    "bookmark_margin": True,
    "change_history": True,
    "edge_mode": 0,                    # 0 none, 1 line, 2 background
    "edge_columns": "80",
    "padding_left": 0,
    "padding_right": 0,
    "multi_edit": True,
    "column_to_multi": True,
    "smooth_font": True,
    "border_width": 2,
    "npc_appearance": "abbrev",        # abbrev | codepoint
    # New document
    "new_eol": 2,                      # SC_EOL_CRLF=0, CR=1, LF=2
    "new_encoding": "utf8",
    "new_lang": "normal",
    "apply_utf8_to_ansi": False,
    # Default directory
    "open_dir_mode": "follow",         # follow | remember | custom
    "open_dir_custom": "",
    "last_dir": "",
    # Recent files
    "recent_max": 10,
    "recent_submenu": False,
    "recent_display": "full",          # name | full
    "recent_maxlen": 100,
    "recent_dont_check": False,
    "recent_files": [],
    # Language
    "lang_compact": False,
    "tab_size": 4,
    "use_spaces": False,
    "lang_tabs": {},                   # lang -> [size, use_spaces]
    "backspace_unindent": False,
    "auto_indent": "advanced",         # none | basic | advanced
    "disabled_langs": [],
    # Highlighting
    "token_match_case": False,
    "token_whole_word": True,
    "smart_hl": True,
    "smart_hl_case": False,
    "smart_hl_word": True,
    "smart_hl_use_find": False,
    "smart_hl_other_view": False,
    "tag_match": True,
    "tag_attr": True,
    "tag_comment": False,
    # Print
    "print_linenums": False,
    "print_color": "wysiwyg",          # wysiwyg | invert | bw | nobg
    "print_margins": [5, 5, 5, 5],     # top left right bottom (mm)
    "print_header": "$(FULL_CURRENT_PATH)",
    "print_footer": "",
    # Searching
    "find_fill_selected": True,
    "find_fill_word": True,
    "find_maxchars": 1024,
    "find_in_sel_auto": True,
    "find_in_sel_min": 1024,
    "find_stay_open": False,
    "confirm_replace_all_docs": True,
    "replace_no_move": False,
    "find_mono": False,
    "find_transparency": "lose",       # none | lose | always
    "find_alpha": 150,
    "find_history": [], "replace_history": [], "filter_history": [], "dir_history": [],
    "find_opts": {},
    # Backup / session
    "remember_session": True,
    "snapshot": True,
    "snapshot_interval": 7,
    "backup_mode": "none",             # none | simple | verbose
    "backup_dir": "",
    # Auto-completion
    "ac_enable": True,
    "ac_mode": "both",                 # func | word | both
    "ac_from": 1,
    "ac_ignore_numbers": True,
    "ac_brief": False,
    "param_hint": True,
    "ai_paren": False, "ai_bracket": False, "ai_brace": False,
    "ai_dquote": False, "ai_squote": False, "ai_tag": False,
    # Multi-instance / date
    "multi_instance": "mono",          # mono | multi | session
    "date_format": "yyyy-MM-dd HH:mm:ss",
    "date_reverse": False,
    # Delimiter
    "extra_word_chars": "",
    "delim_open": "(",
    "delim_close": ")",
    "delim_multiline": False,
    # Performance
    "large_file_restrict": True,
    "large_file_mb": 200,
    "large_no_wrap": True,
    "large_ac": False,
    "large_smart": False,
    "large_brace": False,
    "large_links": False,
    # Cloud & link
    "link_enable": True,
    "link_underline": True,
    "link_fullbox": False,
    "uri_schemes": "svn:// cvs:// git:// imap:// irc:// irc6:// ircs:// ldap:// ldaps:// news: telnet:// gopher:// ssh:// sftp:// smb:// skype: snmp:// spotify: steam:// sms: slack:// chrome:// bitcoin:",
    # Search engine
    "search_engine": "https://duckduckgo.com/?q=$(CURRENT_WORD)",
    # MISC
    "file_autodetect": "enabled",      # enabled | silent | disabled
    "autodetect_scroll_end": False,
    "docswitcher_mru": True,
    "minimize_to_tray": False,
    "close_to_tray": False,
    "title_filename_only": False,
    "autodetect_encoding": True,
    "ansi_codepage": "",               # "" = derived from the locale
    "session_ext": "",
    "workspace_ext": "",
    "mute_sounds": True,
    # View toggles
    "always_on_top": False,
    "word_wrap": False,
    "show_ws": False,
    "show_eol": False,
    "show_npc": False,
    "show_cc": False,
    "show_indent": True,
    "show_wrap_symbol": False,
    "sync_v": False,
    "sync_h": False,
    "zoom_sync": False,
    "props_description": True,         # Properties > Description (the right button): the box under the list
    "code_zoom": 0,                    # zoom of the code windows (Ctrl+Mouse Wheel), kept between sessions
    "panels": {},
    "geometry": "",
    "window_state": "",
    "splitter": "",
    "split_orientation": 1,            # Qt.Orientation.Horizontal
    "workspace_roots": [],
    "projects": ["", "", ""],
    "run_last": "",
    "column_editor": {},
    "hash_treat_lines": False,
}


class Settings:
    def __init__(self):
        self.path = paths.cfg("config.json")
        self.data = {}
        self.load()

    def load(self):
        try:
            with open(self.path, encoding="utf-8") as f:
                self.data = json.load(f)
        except (OSError, ValueError):
            self.data = {}

    def save(self):
        tmp = self.path + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=1, ensure_ascii=False)
            os.replace(tmp, self.path)
        except OSError:
            pass

    def __getitem__(self, key):
        if key in self.data:
            return self.data[key]
        v = DEFAULTS.get(key)
        if isinstance(v, (list, dict)):
            v = json.loads(json.dumps(v))
            self.data[key] = v
        return v

    def get(self, key, default=None):
        if key in self.data:
            return self.data[key]
        return DEFAULTS.get(key, default)

    def __setitem__(self, key, value):
        self.data[key] = value


S = None


def init():
    global S
    S = Settings()
    return S
