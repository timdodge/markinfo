"""Tkinter interface for MarkInfo."""

from __future__ import annotations

import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Optional

from markinfo import __version__
from markinfo.config import (
    AppConfig,
    UI_SCALE_STEPS,
    normalize_ui_scale,
    save_config,
    step_ui_scale,
)
from markinfo.database import DBManager
from markinfo.models import (
    AGE_MAX,
    AGE_MIN,
    KEEPER,
    LEVEL_MAX,
    LEVEL_MIN,
    NAME_MAX,
    NO_CAT,
    NOTES_MAX,
    Player,
    SESSION_MAX,
    SESSION_MIN,
    UNKNOWN,
    notes_indicator,
)
from markinfo.parser import ParseError, process_input_file

ASSET_DIR = Path(__file__).resolve().parent / "assets"
SQUAD_COLUMNS = ("name", "age", "level", "cat", "pos", "type", "sess", "notes")
SQUAD_HEADINGS = ("Name", "Age", "Level", "Cat", "Pos", "Type", "Session", "Notes")
SEARCH_COLUMNS = ("league", "team") + SQUAD_COLUMNS
SEARCH_HEADINGS = ("League", "Team") + SQUAD_HEADINGS
NO_LEAGUES = "NO LEAGUES"
NO_TEAMS = "NO TEAMS"
ALL = "(All)"
BASE_MIN_WIDTH = 760
BASE_MIN_HEIGHT = 560
TREE_COL_WIDTHS = {
    "league": 110,
    "team": 110,
    "name": 140,
    "age": 50,
    "level": 50,
    "cat": 50,
    "pos": 60,
    "type": 60,
    "sess": 70,
    "notes": 70,
}
NAMED_FONTS = (
    "TkDefaultFont",
    "TkTextFont",
    "TkFixedFont",
    "TkMenuFont",
    "TkHeadingFont",
    "TkCaptionFont",
    "TkSmallCaptionFont",
    "TkIconFont",
    "TkTooltipFont",
)


def set_window_icon(win: tk.Misc) -> None:
    ico = ASSET_DIR / "markinfo.ico"
    png = ASSET_DIR / "markinfo.png"
    try:
        if ico.is_file():
            win.iconbitmap(str(ico))
    except tk.TclError:
        pass
    try:
        if png.is_file():
            image = tk.PhotoImage(file=str(png))
            win.iconphoto(True, image)
            win._markinfo_icon = image  # type: ignore[attr-defined]
    except tk.TclError:
        pass


def ask_name(parent: tk.Misc, title: str, prompt: str, initial: str = "") -> Optional[str]:
    value = simpledialog.askstring(title, prompt, parent=parent, initialvalue=initial)
    if value is None:
        return None
    value = value.strip()
    if not value:
        return None
    if len(value) > NAME_MAX:
        messagebox.showerror(
            "Name too long",
            f"Names are limited to {NAME_MAX} characters.",
            parent=parent,
        )
        return None
    return value.upper()


class PlayerDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        db: DBManager,
        league: str,
        teams: list[str],
        current_team: str,
        player: Optional[Player] = None,
        mode: str = "save",
    ) -> None:
        super().__init__(parent)
        self.db = db
        self.league = league
        self.current_team = current_team
        self.mode = mode
        self.result: Optional[Player] = None
        self.title(f"{current_team} - {league} League")
        set_window_icon(self)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.name_var = tk.StringVar(value=player.name if player else "")
        self.age_var = tk.IntVar(value=player.age if player else AGE_MIN)
        self.level_var = tk.IntVar(value=player.level if player else LEVEL_MIN)
        self.cat_var = tk.StringVar(value=player.cat if player else "")
        self.pos_var = tk.StringVar(value=player.pos if player else "")
        self.type_var = tk.StringVar(value=player.type if player else "")
        self.team_var = tk.StringVar(value=current_team)

        pad = {"padx": 8, "pady": 6}
        body = ttk.Frame(self, padding=8)
        body.grid(row=0, column=0, sticky="nsew")

        name_fr = ttk.LabelFrame(body, text="Name")
        name_fr.grid(row=0, column=0, columnspan=2, sticky="ew", **pad)
        self.name_entry = ttk.Entry(name_fr, textvariable=self.name_var, width=24)
        self.name_entry.pack(fill="x", padx=6, pady=6)
        self.name_entry.bind("<FocusOut>", lambda _e: self._upper_name())
        if mode in ("save", "move") and player:
            self.name_entry.state(["disabled"])

        age_fr = ttk.LabelFrame(body, text="Age")
        age_fr.grid(row=0, column=2, sticky="ew", **pad)
        self.age_spin = ttk.Spinbox(
            age_fr,
            from_=AGE_MIN,
            to=AGE_MAX,
            textvariable=self.age_var,
            width=4,
            command=self._age_changed,
        )
        self.age_spin.pack(padx=6, pady=6)
        self.age_spin.bind("<FocusOut>", lambda _e: self._age_changed())

        level_fr = ttk.LabelFrame(body, text="Level")
        level_fr.grid(row=0, column=3, sticky="ew", **pad)
        self.level_spin = ttk.Spinbox(
            level_fr,
            from_=LEVEL_MIN,
            to=LEVEL_MAX,
            textvariable=self.level_var,
            width=4,
            command=self._level_changed,
        )
        self.level_spin.pack(padx=6, pady=6)
        self.level_spin.bind("<FocusOut>", lambda _e: self._level_changed())

        cat_fr = ttk.LabelFrame(body, text="Category")
        cat_fr.grid(row=1, column=0, sticky="ew", **pad)
        self.cat_combo = ttk.Combobox(
            cat_fr, textvariable=self.cat_var, values=self.db.get_cats(),
            state="readonly", width=10,
        )
        self.cat_combo.pack(padx=6, pady=6)
        self.cat_combo.bind("<<ComboboxSelected>>", lambda _e: self._cat_changed())

        pos_fr = ttk.LabelFrame(body, text="Position")
        pos_fr.grid(row=1, column=1, sticky="ew", **pad)
        self.pos_combo = ttk.Combobox(
            pos_fr, textvariable=self.pos_var, values=self.db.get_posns(),
            state="readonly", width=10,
        )
        self.pos_combo.pack(padx=6, pady=6)
        self.pos_combo.bind("<<ComboboxSelected>>", lambda _e: self._pos_changed())

        type_fr = ttk.LabelFrame(body, text="Type")
        type_fr.grid(row=1, column=2, columnspan=2, sticky="ew", **pad)
        self.type_combo = ttk.Combobox(
            type_fr, textvariable=self.type_var, values=self.db.get_types(),
            state="readonly", width=10,
        )
        self.type_combo.pack(padx=6, pady=6)
        self.type_combo.bind("<<ComboboxSelected>>", lambda _e: self._type_changed())

        notes_fr = ttk.LabelFrame(body, text="Notes")
        notes_fr.grid(row=2, column=0, columnspan=4, sticky="ew", **pad)
        self.notes_text = tk.Text(notes_fr, height=4, width=48, wrap="word")
        self.notes_text.pack(fill="both", expand=True, padx=6, pady=6)
        if player and player.notes:
            self.notes_text.insert("1.0", player.notes)

        team_fr = ttk.LabelFrame(body, text="New Team")
        team_fr.grid(row=3, column=0, columnspan=2, sticky="ew", **pad)
        team_values = list(teams)
        if mode == "move" and current_team in team_values:
            team_values = [t for t in team_values if t != current_team]
        self.team_combo = ttk.Combobox(
            team_fr, textvariable=self.team_var, values=team_values,
            state="readonly", width=22,
        )
        self.team_combo.pack(fill="x", padx=6, pady=6)
        if mode != "move":
            self.team_combo.state(["disabled"])
            self.team_var.set(current_team)
        elif team_values:
            self.team_var.set(team_values[0])

        btns = ttk.Frame(body)
        btns.grid(row=3, column=2, columnspan=2, sticky="e", **pad)
        captions = {"add": "Add", "move": "Move", "save": "Save"}
        ttk.Button(btns, text=captions.get(mode, "Save"), command=self._save).pack(
            side="left", padx=4
        )
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="left", padx=4)

        if not self.cat_var.get() and self.cat_combo["values"]:
            self.cat_var.set(self.cat_combo["values"][0])
        if not self.pos_var.get() and self.pos_combo["values"]:
            self.pos_var.set(self.pos_combo["values"][0])
        if not self.type_var.get() and self.type_combo["values"]:
            self.type_var.set(self.type_combo["values"][0])

        self._age_changed()
        self.bind("<Escape>", lambda _e: self.destroy())
        if mode == "add":
            self.name_entry.focus_set()
        else:
            self.age_spin.focus_set()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.wait_window(self)

    def _upper_name(self) -> None:
        self.name_var.set(self.name_var.get().upper())

    def _age(self) -> int:
        try:
            return int(self.age_var.get())
        except (tk.TclError, ValueError):
            return AGE_MIN

    def _level(self) -> int:
        try:
            return int(self.level_var.get())
        except (tk.TclError, ValueError):
            return LEVEL_MIN

    def _age_changed(self) -> None:
        age = self._age()
        ptype = self.type_var.get()
        if age == 17:
            self.type_var.set("SBY")
        if age > 18 and ptype == "SBY":
            self.type_var.set(NO_CAT)
        if age != 18 and ptype in ("APP", "FUT"):
            self.type_var.set(NO_CAT)
        if age < 19:
            if self._level() > 12:
                self.level_var.set(12)
                self.type_var.set(NO_CAT)
            self.level_spin.configure(to=12)
        else:
            self.level_spin.configure(to=LEVEL_MAX)
        self._type_changed()

    def _level_changed(self) -> None:
        level = self._level()
        if level > 12:
            self.type_var.set("STAR")
        if level < 12 and self.type_var.get() == "STAR":
            self.type_var.set(NO_CAT)

    def _cat_changed(self) -> None:
        if self.cat_var.get() == NO_CAT:
            self.pos_var.set(KEEPER)

    def _pos_changed(self) -> None:
        if self.pos_var.get() == KEEPER:
            self.cat_var.set(NO_CAT)
        elif self.cat_var.get() == NO_CAT:
            self.cat_var.set(UNKNOWN)

    def _type_changed(self) -> None:
        ptype = self.type_var.get()
        age = self._age()
        level = self._level()
        if ptype == "STAR" and level < 12:
            self.type_var.set("SBY" if age == 17 else NO_CAT)
        if ptype == "SBY" and age > 18:
            self.type_var.set(NO_CAT)
        if ptype in ("APP", "FUT") and age != 18:
            self.type_var.set(NO_CAT)

    def _save(self) -> None:
        self._upper_name()
        self._age_changed()
        self._level_changed()
        self._pos_changed()
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror("Name required", "Please name the player!", parent=self)
            return
        action = {"add": "Add", "move": "Move", "save": "Save"}.get(self.mode, "Save")
        if not messagebox.askyesno("Confirm Changes", f"{action} player {name}?", parent=self):
            return
        notes = self.notes_text.get("1.0", "end").strip()[:NOTES_MAX]
        self.result = Player(
            league=self.league,
            team=self.team_var.get(),
            name=name,
            age=self._age(),
            level=self._level(),
            cat=self.cat_var.get(),
            pos=self.pos_var.get(),
            type=self.type_var.get(),
            sess=0,
            notes=notes,
        )
        self.destroy()


class NewSeasonDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, league: str) -> None:
        super().__init__(parent)
        self.league = league
        self.result: Optional[int] = None
        self.title(f"New Season - {league} League")
        set_window_icon(self)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.limit_var = tk.IntVar(value=2)
        body = ttk.LabelFrame(self, text="Player Pruning", padding=12)
        body.pack(fill="both", expand=True, padx=10, pady=10)
        ttk.Label(body, text="Remove players whose details").pack(anchor="w")
        ttk.Label(body, text="have not been updated for").pack(anchor="w")
        row = ttk.Frame(body)
        row.pack(anchor="w", pady=8)
        ttk.Label(row, text="more than").pack(side="left", padx=(0, 6))
        ttk.Spinbox(row, from_=1, to=16, textvariable=self.limit_var, width=4).pack(
            side="left"
        )
        ttk.Label(row, text="sessions").pack(side="left", padx=6)
        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(btns, text="OK", command=self._ok).pack(side="right", padx=4)
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="right")
        self.bind("<Escape>", lambda _e: self.destroy())
        self.wait_window(self)

    def _ok(self) -> None:
        try:
            self.result = int(self.limit_var.get())
        except (tk.TclError, ValueError):
            self.result = 2
        self.destroy()


class MarkInfoApp:
    def __init__(self, db: DBManager, config: Optional[AppConfig] = None) -> None:
        self.db = db
        self.config = config or AppConfig()
        self.squad: list[Player] = []
        self.search_results: list[Player] = []
        self._tree_sort: dict[int, dict] = {}
        self._trees: list[ttk.Treeview] = []
        self._ui_scale = self.config.ui_scale
        self.root = tk.Tk()
        self.root.title("MarkInfo 3.0")
        self.root.minsize(BASE_MIN_WIDTH, BASE_MIN_HEIGHT)
        set_window_icon(self.root)
        try:
            ttk.Style().theme_use("clam")
        except tk.TclError:
            pass
        self._ui_scale_var = tk.IntVar(value=int(round(self._ui_scale * 100)))
        self._capture_scale_baselines()
        self._apply_ui_scale()
        self._build_menu()
        self._build_body()
        if abs(self._ui_scale - 1.0) > 1e-9:
            self.root.geometry(
                f"{int(round(BASE_MIN_WIDTH * self._ui_scale))}x"
                f"{int(round(BASE_MIN_HEIGHT * self._ui_scale))}"
            )
        self.session_var.set(self.db.session)
        self.refresh_leagues()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.bind("<Control-n>", lambda _e: self._new_season())
        self.root.bind("<Control-i>", lambda _e: self._import_shomatch())
        self.root.bind("<Control-l>", lambda _e: self._add_league())
        self.root.bind("<Control-t>", lambda _e: self._add_team())
        self.root.bind("<Control-plus>", lambda _e: self._nudge_ui_scale(1))
        self.root.bind("<Control-equal>", lambda _e: self._nudge_ui_scale(1))
        self.root.bind("<Control-KP_Add>", lambda _e: self._nudge_ui_scale(1))
        self.root.bind("<Control-minus>", lambda _e: self._nudge_ui_scale(-1))
        self.root.bind("<Control-KP_Subtract>", lambda _e: self._nudge_ui_scale(-1))
        self.root.bind("<Control-0>", lambda _e: self._set_ui_scale(1.0))
        self.root.bind("<Control-KP_0>", lambda _e: self._set_ui_scale(1.0))

    def run(self) -> None:
        self.root.mainloop()

    # -- construction -------------------------------------------------------

    def _build_menu(self) -> None:
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Season", command=self._new_season, accelerator="Ctrl+N")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._on_close)
        menubar.add_cascade(label="File", menu=file_menu)

        maint = tk.Menu(menubar, tearoff=0)
        maint.add_command(
            label="Import Shomatch File",
            command=self._import_shomatch,
            accelerator="Ctrl+I",
        )
        maint.add_separator()
        maint.add_command(label="New League", command=self._add_league, accelerator="Ctrl+L")
        maint.add_command(label="Delete League", command=self._del_league)
        maint.add_command(label="Rename League", command=self._rename_league)
        maint.add_separator()
        maint.add_command(label="New Team", command=self._add_team, accelerator="Ctrl+T")
        maint.add_command(label="Delete Team", command=self._del_team)
        maint.add_command(label="Rename Team", command=self._rename_team)
        menubar.add_cascade(label="Maintenance", menu=maint)

        view = tk.Menu(menubar, tearoff=0)
        size_menu = tk.Menu(view, tearoff=0)
        for step in UI_SCALE_STEPS:
            percent = int(round(step * 100))
            size_menu.add_radiobutton(
                label=f"{percent}%",
                value=percent,
                variable=self._ui_scale_var,
                command=self._on_ui_scale_menu,
            )
        view.add_cascade(label="UI Size", menu=size_menu)
        view.add_separator()
        view.add_command(
            label="Increase UI Size",
            command=lambda: self._nudge_ui_scale(1),
            accelerator="Ctrl++",
        )
        view.add_command(
            label="Decrease UI Size",
            command=lambda: self._nudge_ui_scale(-1),
            accelerator="Ctrl+-",
        )
        view.add_command(
            label="Reset UI Size",
            command=lambda: self._set_ui_scale(1.0),
            accelerator="Ctrl+0",
        )
        menubar.add_cascade(label="View", menu=view)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self._about)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.root.config(menu=menubar)

    def _build_body(self) -> None:
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill="x")

        team_fr = ttk.LabelFrame(top, text="Selected Team", padding=8)
        team_fr.pack(side="left", fill="x", expand=True)
        ttk.Label(team_fr, text="League :").grid(row=0, column=0, sticky="e", padx=4, pady=2)
        self.league_var = tk.StringVar()
        self.league_combo = ttk.Combobox(
            team_fr, textvariable=self.league_var, state="readonly", width=28
        )
        self.league_combo.grid(row=0, column=1, sticky="ew", padx=4, pady=2)
        self.league_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_teams())
        ttk.Label(team_fr, text="Team :").grid(row=1, column=0, sticky="e", padx=4, pady=2)
        self.team_var = tk.StringVar()
        self.team_combo = ttk.Combobox(
            team_fr, textvariable=self.team_var, state="readonly", width=28
        )
        self.team_combo.grid(row=1, column=1, sticky="ew", padx=4, pady=2)
        self.team_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_squad())
        team_fr.columnconfigure(1, weight=1)

        sess_fr = ttk.LabelFrame(top, text="Session", padding=8)
        sess_fr.pack(side="left", padx=8, fill="y")
        self.session_var = tk.IntVar(value=1)
        ttk.Spinbox(
            sess_fr,
            from_=SESSION_MIN,
            to=SESSION_MAX,
            textvariable=self.session_var,
            width=4,
            command=self._save_session,
        ).pack(padx=8, pady=8)

        ttk.Button(
            top, text="Import\nShomatch File", command=self._import_shomatch, width=16
        ).pack(side="left", fill="y", padx=(0, 4))

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.notebook = notebook

        squad_tab = ttk.Frame(notebook, padding=4)
        notebook.add(squad_tab, text="Squad")
        self.squad_tree = self._make_tree(squad_tab, SQUAD_COLUMNS, SQUAD_HEADINGS)
        self.squad_tree.bind("<Double-1>", lambda _e: self._edit_player())
        self.squad_tree.bind("<Return>", lambda _e: self._edit_player())

        btn_fr = ttk.LabelFrame(squad_tab, text="Player Maintenance", padding=8)
        btn_fr.pack(fill="x", pady=(6, 0))
        for label, cmd in (
            ("Edit Player", self._edit_player),
            ("Move Player", self._move_player),
            ("Add Player", self._add_player),
            ("Delete Player", self._del_player),
        ):
            ttk.Button(btn_fr, text=label, command=cmd).pack(side="left", expand=True, padx=4)

        search_tab = ttk.Frame(notebook, padding=4)
        notebook.add(search_tab, text="Search")
        self._build_search(search_tab)

        self.status = tk.StringVar(value="Ready")
        ttk.Label(self.root, textvariable=self.status, relief="sunken", anchor="w").pack(
            fill="x", side="bottom"
        )

    def _build_search(self, parent: ttk.Frame) -> None:
        filters = ttk.LabelFrame(parent, text="Query", padding=8)
        filters.pack(fill="x")
        self.q_name = tk.StringVar()
        self.q_league = tk.StringVar(value=ALL)
        self.q_team = tk.StringVar(value=ALL)
        self.q_pos = tk.StringVar(value=ALL)
        self.q_cat = tk.StringVar(value=ALL)
        self.q_type = tk.StringVar(value=ALL)
        self.q_min_age = tk.StringVar()
        self.q_max_age = tk.StringVar()
        self.q_min_level = tk.StringVar()
        self.q_max_level = tk.StringVar()

        def labelled(row: int, col: int, text: str, widget: tk.Widget) -> None:
            ttk.Label(filters, text=text).grid(row=row, column=col, sticky="e", padx=4, pady=2)
            widget.grid(row=row, column=col + 1, sticky="ew", padx=4, pady=2)

        labelled(0, 0, "Name :", ttk.Entry(filters, textvariable=self.q_name, width=20))
        self.q_league_combo = ttk.Combobox(
            filters, textvariable=self.q_league, state="readonly", width=18
        )
        labelled(0, 2, "League :", self.q_league_combo)
        self.q_team_combo = ttk.Combobox(
            filters, textvariable=self.q_team, state="readonly", width=18
        )
        labelled(0, 4, "Team :", self.q_team_combo)
        self.q_pos_combo = ttk.Combobox(
            filters, textvariable=self.q_pos, state="readonly", width=10
        )
        labelled(1, 0, "Pos :", self.q_pos_combo)
        self.q_cat_combo = ttk.Combobox(
            filters, textvariable=self.q_cat, state="readonly", width=10
        )
        labelled(1, 2, "Cat :", self.q_cat_combo)
        self.q_type_combo = ttk.Combobox(
            filters, textvariable=self.q_type, state="readonly", width=10
        )
        labelled(1, 4, "Type :", self.q_type_combo)
        labelled(2, 0, "Age from :", ttk.Entry(filters, textvariable=self.q_min_age, width=6))
        labelled(2, 2, "to :", ttk.Entry(filters, textvariable=self.q_max_age, width=6))
        labelled(2, 4, "Level from :", ttk.Entry(filters, textvariable=self.q_min_level, width=6))
        ttk.Label(filters, text="to :").grid(row=2, column=6, sticky="e", padx=4)
        ttk.Entry(filters, textvariable=self.q_max_level, width=6).grid(
            row=2, column=7, sticky="w", padx=4, pady=2
        )
        ttk.Button(filters, text="Search", command=self._run_search).grid(
            row=0, column=6, rowspan=2, padx=8, sticky="ns"
        )
        self.q_league_combo.bind("<<ComboboxSelected>>", lambda _e: self._search_league_changed())
        self.search_tree = self._make_tree(parent, SEARCH_COLUMNS, SEARCH_HEADINGS)
        self.search_tree.bind("<Double-1>", lambda _e: self._edit_search_player())
        self._refresh_search_filters()

    def _make_tree(
        self, parent: ttk.Frame, columns: tuple[str, ...], headings: tuple[str, ...]
    ) -> ttk.Treeview:
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        yscroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=yscroll.set)
        tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")
        factor = self._ui_scale
        for col, heading in zip(columns, headings):
            tree.heading(col, text=heading, command=lambda c=col: self._sort_tree(tree, c))
            tree.column(
                col,
                width=int(TREE_COL_WIDTHS.get(col, 80) * factor),
                anchor="center" if col != "name" else "w",
            )
        tree.tag_configure("old", foreground="#666666")
        self._tree_sort[id(tree)] = {
            "col": None,
            "reverse": False,
            "headings": dict(zip(columns, headings)),
        }
        self._trees.append(tree)
        return tree

    def _sort_tree(self, tree: ttk.Treeview, col: str) -> None:
        state = self._tree_sort[id(tree)]
        if state["col"] == col:
            state["reverse"] = not state["reverse"]
        else:
            state["col"] = col
            state["reverse"] = False
        self._apply_sort(tree)

    def _apply_sort(self, tree: ttk.Treeview) -> None:
        state = self._tree_sort.get(id(tree))
        if not state or not state["col"]:
            return
        col = state["col"]
        reverse = state["reverse"]
        rows = [(tree.set(item, col), item) for item in tree.get_children("")]
        numeric = all(not value or _is_int(value) for value, _item in rows)
        if numeric:
            rows.sort(key=lambda pair: int(pair[0] or 0), reverse=reverse)
        else:
            rows.sort(key=lambda pair: str(pair[0]), reverse=reverse)
        for index, (_value, item) in enumerate(rows):
            tree.move(item, "", index)
        for name, heading in state["headings"].items():
            suffix = ""
            if name == col:
                suffix = " ↓" if reverse else " ↑"
            tree.heading(name, text=heading + suffix)

    # -- UI scale -----------------------------------------------------------

    def _capture_scale_baselines(self) -> None:
        self._base_fonts: dict[str, dict] = {}
        for name in NAMED_FONTS:
            try:
                named = tkfont.nametofont(name)
            except tk.TclError:
                continue
            size = int(named.cget("size") or 0)
            if size == 0:
                size = int(named.actual("size") or 10)
            self._base_fonts[name] = {
                "family": named.cget("family"),
                "size": size,
                "weight": named.cget("weight"),
                "slant": named.cget("slant"),
            }

    def _on_ui_scale_menu(self) -> None:
        self._set_ui_scale(self._ui_scale_var.get() / 100.0)

    def _nudge_ui_scale(self, delta: int) -> None:
        self._set_ui_scale(step_ui_scale(self._ui_scale, delta))

    def _set_ui_scale(self, factor: float) -> None:
        factor = normalize_ui_scale(factor)
        if abs(factor - self._ui_scale) < 1e-9:
            self._ui_scale_var.set(int(round(factor * 100)))
            return
        self._ui_scale = factor
        self.config.ui_scale = factor
        self._ui_scale_var.set(int(round(factor * 100)))
        self._apply_ui_scale()
        try:
            save_config(self.config)
        except OSError as exc:
            messagebox.showwarning("UI size", f"Could not save the UI size setting:\n{exc}")
        if hasattr(self, "status"):
            self.status.set(f"UI size set to {int(round(factor * 100))}%")

    def _apply_ui_scale(self) -> None:
        factor = self._ui_scale
        for name, info in self._base_fonts.items():
            try:
                named = tkfont.nametofont(name)
            except tk.TclError:
                continue
            size = int(info["size"])
            new_size = max(1, int(round(abs(size) * factor)))
            if size < 0:
                new_size = -new_size
            named.configure(size=new_size)
        style = ttk.Style(self.root)
        default_font = tkfont.nametofont("TkDefaultFont")
        heading_font = tkfont.nametofont("TkHeadingFont")
        line = int(default_font.metrics("linespace") or 16)
        rowheight = max(line + 8, int(round(22 * factor)))
        arrow = max(10, int(round(12 * factor)))
        pad_x = max(4, int(round(6 * factor)))
        pad_y = max(2, int(round(4 * factor)))
        style.configure(".", font=default_font)
        style.configure("Treeview", font=default_font, rowheight=rowheight)
        style.configure("Treeview.Heading", font=heading_font)
        style.configure("TButton", padding=(pad_x, pad_y))
        style.configure("TNotebook.Tab", padding=(pad_x, pad_y))
        try:
            style.configure("TSpinbox", arrowsize=arrow)
            style.configure("TCombobox", arrowsize=arrow)
            style.configure("Vertical.TScrollbar", arrowsize=arrow, width=arrow)
            style.configure("Horizontal.TScrollbar", arrowsize=arrow, width=arrow)
        except tk.TclError:
            pass
        min_w = int(round(BASE_MIN_WIDTH * factor))
        min_h = int(round(BASE_MIN_HEIGHT * factor))
        self.root.minsize(min_w, min_h)
        self._scale_tree_columns()
        if self.root.winfo_ismapped():
            self.root.update_idletasks()
            width = max(self.root.winfo_width(), min_w)
            height = max(self.root.winfo_height(), min_h)
            self.root.geometry(f"{width}x{height}")

    def _scale_tree_columns(self) -> None:
        factor = self._ui_scale
        for tree in self._trees:
            for col in tree["columns"]:
                tree.column(col, width=int(TREE_COL_WIDTHS.get(col, 80) * factor))

    # -- data refresh -------------------------------------------------------

    def _save_session(self) -> None:
        try:
            self.db.session = int(self.session_var.get())
        except (tk.TclError, ValueError):
            pass

    def current_league(self) -> Optional[str]:
        value = self.league_var.get()
        if not value or value == NO_LEAGUES:
            return None
        return value

    def current_team(self) -> Optional[str]:
        value = self.team_var.get()
        if not value or value == NO_TEAMS:
            return None
        return value

    def refresh_leagues(self, select: Optional[str] = None) -> None:
        leagues = self.db.get_leagues()
        if not leagues:
            self.league_combo["values"] = [NO_LEAGUES]
            self.league_var.set(NO_LEAGUES)
        else:
            self.league_combo["values"] = leagues
            if select in leagues:
                self.league_var.set(select)
            elif self.league_var.get() not in leagues:
                self.league_var.set(leagues[0])
        self.refresh_teams()
        self._refresh_search_filters()

    def refresh_teams(self, select: Optional[str] = None) -> None:
        league = self.current_league()
        if not league:
            self.team_combo["values"] = [NO_TEAMS]
            self.team_var.set(NO_TEAMS)
            self.refresh_squad()
            return
        teams = self.db.get_teams(league)
        if not teams:
            self.team_combo["values"] = [NO_TEAMS]
            self.team_var.set(NO_TEAMS)
        else:
            self.team_combo["values"] = teams
            if select in teams:
                self.team_var.set(select)
            elif self.team_var.get() not in teams:
                self.team_var.set(teams[0])
        self.refresh_squad()

    def refresh_squad(self) -> None:
        for item in self.squad_tree.get_children():
            self.squad_tree.delete(item)
        league = self.current_league()
        team = self.current_team()
        if not league or not team:
            self.squad = []
            self.status.set("No team selected")
            return
        self.squad = list(self.db.get_squad(team, league))
        for player in self.squad:
            tags = ("old",) if player.sess < 0 else ()
            self.squad_tree.insert(
                "",
                "end",
                values=(
                    player.name,
                    player.age,
                    player.level,
                    player.cat,
                    player.pos,
                    player.type,
                    player.sess,
                    notes_indicator(player.notes),
                ),
                tags=tags,
            )
        self._apply_sort(self.squad_tree)
        self.status.set(f"{len(self.squad)} player(s) in {team}")

    def _selected_squad_player(self) -> Optional[Player]:
        selection = self.squad_tree.selection()
        if not selection:
            return None
        values = self.squad_tree.item(selection[0], "values")
        if not values:
            return None
        name = values[0]
        for player in self.squad:
            if player.name == name:
                return player
        return None

    def _refresh_search_filters(self) -> None:
        leagues = [ALL] + self.db.get_leagues()
        self.q_league_combo["values"] = leagues
        if self.q_league.get() not in leagues:
            self.q_league.set(ALL)
        self._search_league_changed()
        self.q_pos_combo["values"] = [ALL] + self.db.get_posns()
        self.q_cat_combo["values"] = [ALL] + self.db.get_cats()
        self.q_type_combo["values"] = [ALL] + self.db.get_types()
        for combo, var in (
            (self.q_pos_combo, self.q_pos),
            (self.q_cat_combo, self.q_cat),
            (self.q_type_combo, self.q_type),
        ):
            if var.get() not in combo["values"]:
                var.set(ALL)

    def _search_league_changed(self) -> None:
        league = self.q_league.get()
        if league == ALL:
            teams = [ALL]
            for name in self.db.get_leagues():
                teams.extend(self.db.get_teams(name))
            teams = [ALL] + sorted(set(teams) - {ALL})
        else:
            teams = [ALL] + self.db.get_teams(league)
        self.q_team_combo["values"] = teams
        if self.q_team.get() not in teams:
            self.q_team.set(ALL)

    def _optional_int(self, raw: str) -> Optional[int]:
        raw = raw.strip()
        if not raw:
            return None
        return int(raw)

    def _run_search(self) -> None:
        try:
            min_age = self._optional_int(self.q_min_age.get())
            max_age = self._optional_int(self.q_max_age.get())
            min_level = self._optional_int(self.q_min_level.get())
            max_level = self._optional_int(self.q_max_level.get())
        except ValueError:
            messagebox.showerror("Invalid query", "Age and level filters must be numbers.")
            return
        self.search_results = self.db.search_players(
            name=self.q_name.get(),
            league="" if self.q_league.get() == ALL else self.q_league.get(),
            team="" if self.q_team.get() == ALL else self.q_team.get(),
            pos="" if self.q_pos.get() == ALL else self.q_pos.get(),
            cat="" if self.q_cat.get() == ALL else self.q_cat.get(),
            ptype="" if self.q_type.get() == ALL else self.q_type.get(),
            min_age=min_age,
            max_age=max_age,
            min_level=min_level,
            max_level=max_level,
        )
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)
        for player in self.search_results:
            tags = ("old",) if player.sess < 0 else ()
            self.search_tree.insert(
                "",
                "end",
                values=(
                    player.league,
                    player.team,
                    player.name,
                    player.age,
                    player.level,
                    player.cat,
                    player.pos,
                    player.type,
                    player.sess,
                    notes_indicator(player.notes),
                ),
                tags=tags,
            )
        self._apply_sort(self.search_tree)
        self.status.set(f"{len(self.search_results)} player(s) matched")

    def _selected_search_player(self) -> Optional[Player]:
        selection = self.search_tree.selection()
        if not selection:
            return None
        values = self.search_tree.item(selection[0], "values")
        if not values:
            return None
        league, team, name = values[0], values[1], values[2]
        for player in self.search_results:
            if player.league == league and player.team == team and player.name == name:
                return player
        return None

    # -- player actions -----------------------------------------------------

    def _need_team(self) -> bool:
        if not self.current_league():
            messagebox.showinfo("MarkInfo", "You have to add a league first!")
            return False
        if not self.current_team():
            messagebox.showinfo("MarkInfo", "No Teams!")
            return False
        return True

    def _add_player(self) -> None:
        if not self._need_team():
            return
        self._save_session()
        league = self.current_league()
        team = self.current_team()
        assert league and team
        dialog = PlayerDialog(
            self.root, self.db, league, list(self.team_combo["values"]), team, mode="add"
        )
        if dialog.result:
            self.db.add_player(dialog.result)
            self.refresh_squad()

    def _edit_player(self) -> None:
        player = self._selected_squad_player()
        if not player or not self._need_team():
            return
        self._save_session()
        league = self.current_league()
        team = self.current_team()
        assert league and team
        dialog = PlayerDialog(
            self.root,
            self.db,
            league,
            list(self.team_combo["values"]),
            team,
            player=player,
            mode="save",
        )
        if dialog.result:
            self.db.add_player(dialog.result)
            self.refresh_squad()

    def _move_player(self) -> None:
        player = self._selected_squad_player()
        if not player or not self._need_team():
            return
        teams = [t for t in self.team_combo["values"] if t not in (NO_TEAMS, "")]
        if len(teams) < 2:
            messagebox.showinfo("MarkInfo", "There is no other team to move to.")
            return
        self._save_session()
        league = self.current_league()
        team = self.current_team()
        assert league and team
        dialog = PlayerDialog(
            self.root, self.db, league, teams, team, player=player, mode="move"
        )
        if dialog.result:
            self.db.move_player(dialog.result, team)
            self.refresh_squad()

    def _del_player(self) -> None:
        player = self._selected_squad_player()
        if not player or not self._need_team():
            return
        if messagebox.askyesno("Confirm Delete", f"Really delete {player.name}?"):
            self.db.del_player(player.name, player.team, player.league)
            self.refresh_squad()

    def _edit_search_player(self) -> None:
        player = self._selected_search_player()
        if not player:
            return
        self._save_session()
        teams = self.db.get_teams(player.league) or [player.team]
        dialog = PlayerDialog(
            self.root, self.db, player.league, teams, player.team, player=player, mode="save"
        )
        if dialog.result:
            self.db.add_player(dialog.result)
            self.refresh_squad()
            self._run_search()

    # -- import / season ----------------------------------------------------

    def _import_shomatch(self) -> None:
        league = self.current_league()
        if not league:
            messagebox.showinfo("MarkInfo", "You have to add a league first!")
            return
        self._save_session()
        path = filedialog.askopenfilename(
            title="Choose Shomatch File",
            filetypes=[
                ("ShoMatch files", "*.zip"),
                ("Text files", "*.txt"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        self.root.config(cursor="watch")
        self.root.update_idletasks()
        try:
            added, updated = process_input_file(
                path,
                league,
                self.db.add_player,
                self.db.move_player,
            )
        except ParseError as exc:
            messagebox.showerror("Import error", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("Import error", str(exc))
            return
        finally:
            self.root.config(cursor="")
        self.refresh_teams()
        messagebox.showinfo(
            "Import complete", f"{added} players added and {updated} updated"
        )

    def _new_season(self) -> None:
        league = self.current_league()
        if not league:
            messagebox.showinfo("MarkInfo", "You have to add a league first!")
            return
        dialog = NewSeasonDialog(self.root, league)
        if dialog.result is None:
            return
        self.root.config(cursor="watch")
        self.root.update_idletasks()
        try:
            self.db.new_season(league, dialog.result)
        finally:
            self.root.config(cursor="")
        self.refresh_squad()

    # -- league / team maintenance -----------------------------------------

    def _add_league(self) -> None:
        name = ask_name(self.root, "New League", "Enter the name of the new League")
        if not name:
            return
        self.db.add_league(name)
        self.refresh_leagues(select=name)

    def _del_league(self) -> None:
        league = self.current_league()
        if not league:
            messagebox.showinfo("MarkInfo", "You have to add a league first!")
            return
        if messagebox.askyesno("Confirm Delete", f"Really delete {league}?"):
            self.db.del_league(league)
            self.refresh_leagues()

    def _rename_league(self) -> None:
        league = self.current_league()
        if not league:
            messagebox.showinfo("MarkInfo", "No Leagues!")
            return
        name = ask_name(
            self.root, f"Rename {league}", f"Enter the new name for {league}", league
        )
        if not name:
            return
        self.db.rename_league(league, name)
        self.refresh_leagues(select=name)

    def _add_team(self) -> None:
        league = self.current_league()
        if not league:
            messagebox.showinfo("MarkInfo", "You have to add a league first!")
            return
        name = ask_name(
            self.root, "New Team", f"Enter the name of the new Team in {league}"
        )
        if not name:
            return
        self.db.add_team(name, league)
        self.refresh_teams(select=name)

    def _del_team(self) -> None:
        if not self._need_team():
            return
        team = self.current_team()
        league = self.current_league()
        assert team and league
        if messagebox.askyesno("Confirm Delete", f"Really delete {team}?"):
            self.db.del_team(team, league)
            self.refresh_teams()

    def _rename_team(self) -> None:
        if not self._need_team():
            return
        team = self.current_team()
        league = self.current_league()
        assert team and league
        name = ask_name(
            self.root, f"Rename {team}", f"Enter the new name for {team}", team
        )
        if not name:
            return
        self.db.rename_team(team, name, league)
        self.refresh_teams(select=name)

    def _about(self) -> None:
        messagebox.showinfo(
            f"MarkInfo v{__version__}",
            "Copyright 2000, 2001 Tim Dodge\n\n"
            "Python port of MarkInfo V2, a player database for "
            "Spellbinder Games' Kickabout.\n\n"
            "MarkInfo comes with ABSOLUTELY NO WARRANTY.\n\n"
            "This is free software; you are welcome to redistribute it "
            "under the GNU General Public License version 2 or later.",
        )

    def _on_close(self) -> None:
        self._save_session()
        self.db.close()
        self.root.destroy()


def _is_int(value: str) -> bool:
    try:
        int(value)
        return True
    except ValueError:
        return False
